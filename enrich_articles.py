"""
enrich_articles.py
Collect richer metadata for the journal articles on the workbook's "Metadata for Deposit" tab:

  Scopus Abstract Retrieval (by DOI)  -> abstract, each author's own affiliation(s), ORCIDs, Scopus author IDs,
                                         author keywords, index terms, subject areas, funding/award numbers
  PubMed (NCBI E-utilities, by PMID)  -> MeSH headings (descriptor + UI + qualifiers + major topic)
                                         (also the fallback for abstract/authors/affiliations/grants)
  Crossref (by DOI)                   -> the publisher's license, funders/awards, dates, fallback authors/abstract
  Unpaywall (by DOI)                  -> backup license and OA status
  Prism (optional, read-only)         -> how Prism's subject vocabulary stores each MeSH heading

Nothing is written to the deposit workbook. Results go to an Excel file for review/merging, and every raw
API response is saved so nothing has to be fetched twice.

Used from Enrich_Articles.ipynb. Needs: requests, openpyxl.
"""
import difflib
import json
import re
import time
import xml.etree.ElementTree as ET
from pathlib import Path

import requests

# ------------------------------------------------------------------ helpers
def aslist(x):
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def txt(x):
    """Scopus JSON values are strings, {'$': value} dicts, or lists of those."""
    if x is None:
        return ""
    if isinstance(x, str):
        return x.strip()
    if isinstance(x, dict):
        return txt(x.get("$") or x.get("#text") or "")
    if isinstance(x, list):
        return "; ".join(t for t in (txt(i) for i in x) if t)
    return str(x)


def load_config(path):
    """Read settings from a config.py (variables like SCOPUS_API_KEY = "...") or a .env file.
    Returns a dict of NAME -> value. Values are never printed."""
    p = Path(path)
    if p.suffix.lower() == ".py":
        import importlib.util
        spec = importlib.util.spec_from_file_location("_prism_config", p)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return {k: str(v) for k, v in vars(mod).items() if k.isupper() and isinstance(v, (str, int))}
    return load_env(p)


def load_env(path):
    """Minimal .env reader (KEY=VALUE lines). Values are never printed."""
    env = {}
    p = Path(path)
    if p.exists():
        for line in p.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def norm_title(t):
    return re.sub(r"[^a-z0-9 ]", " ", (t or "").lower()).split()


def title_similarity(a, b):
    return round(difflib.SequenceMatcher(None, " ".join(norm_title(a)), " ".join(norm_title(b))).ratio(), 2)


def clean_abstract(a):
    """Remove publisher copyright lines that Scopus appends/prepends to abstracts."""
    a = re.sub(r"\s+", " ", a or "").strip()
    # leading: "Copyright © 2021 Elsevier Inc. All rights reserved." / "© 2021 The Authors."
    a = re.sub(r"^(Copyright\s*)?(©|\(c\))?\s*(Copyright\s*)?\d{4}\b[^.]*\.(\s*All rights reserved\.)?\s*", "", a, flags=re.I) \
        if re.match(r"^(Copyright|©|\(c\))", a, flags=re.I) else a
    # trailing: "... results. © 2024 The Author(s)." / "Copyright © 2021 ..."
    a = re.sub(r"\s*(Copyright\s*)?(©|\(c\))\s*\d{4}.*$", "", a, flags=re.I)
    a = re.sub(r"\s*Copyright\s+\d{4}\b.*$", "", a)
    return a.strip()


LICENSE_MAP = [   # (pattern in license URL or label, SPDX id used by InvenioRDM, display title)
    (r"creativecommons\.org/licenses/by-nc-nd/4\.0|cc-by-nc-nd", "cc-by-nc-nd-4.0", "Creative Commons Attribution Non Commercial No Derivatives 4.0 International"),
    (r"creativecommons\.org/licenses/by-nc-sa/4\.0|cc-by-nc-sa", "cc-by-nc-sa-4.0", "Creative Commons Attribution Non Commercial Share Alike 4.0 International"),
    (r"creativecommons\.org/licenses/by-nc/4\.0|cc-by-nc\b", "cc-by-nc-4.0", "Creative Commons Attribution Non Commercial 4.0 International"),
    (r"creativecommons\.org/licenses/by-nd/4\.0|cc-by-nd\b", "cc-by-nd-4.0", "Creative Commons Attribution No Derivatives 4.0 International"),
    (r"creativecommons\.org/licenses/by-sa/4\.0|cc-by-sa\b", "cc-by-sa-4.0", "Creative Commons Attribution Share Alike 4.0 International"),
    (r"creativecommons\.org/licenses/by/4\.0|^cc-by$", "cc-by-4.0", "Creative Commons Attribution 4.0 International"),
    (r"creativecommons\.org/licenses/by-nc-nd/3\.0", "cc-by-nc-nd-3.0", "Creative Commons Attribution Non Commercial No Derivatives 3.0"),
    (r"creativecommons\.org/licenses/by-nc/3\.0", "cc-by-nc-3.0", "Creative Commons Attribution Non Commercial 3.0"),
    (r"creativecommons\.org/licenses/by/3\.0", "cc-by-3.0", "Creative Commons Attribution 3.0"),
    (r"creativecommons\.org/licenses/by/2\.0", "cc-by-2.0", "Creative Commons Attribution 2.0"),
    (r"creativecommons\.org/publicdomain/zero", "cc0-1.0", "Creative Commons Zero v1.0 Universal"),
    (r"public-domain|publicdomain/mark", "", "Public domain (US Government work)"),
]


def map_license(url_or_label):
    s = (url_or_label or "").strip().lower()
    for pat, spdx, title in LICENSE_MAP:
        if re.search(pat, s):
            return spdx, title
    return "", ""


# ------------------------------------------------------------------ input
def read_articles(workbook, sheet="Metadata for Deposit"):
    import openpyxl
    ws = openpyxl.load_workbook(workbook, data_only=True)[sheet]
    head = [c.value for c in ws[1]]
    ix = {h: i for i, h in enumerate(head) if h}
    out = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[ix["Resource type"]] != "journal article" or not r[ix["record_id"]]:
            continue
        doi = re.sub(r"^https?://(dx\.)?doi\.org/", "", str(r[ix["DOI"]] or "").strip(), flags=re.I)
        pmid = ""
        if str(r[ix["Alternate identifiers: Scheme"]] or "").upper() == "PMID":
            pmid = str(r[ix["Alternate Identifiers: Identifier"]] or "").split(".")[0]
        out.append({"record_id": r[ix["record_id"]], "doi": doi, "pmid": pmid, "title": r[ix["Title"]],
                    "subjects_now": r[ix["Subject (term)"]] or ""})
    return out


# ------------------------------------------------------------------ API client
class Client:
    def __init__(self, scopus_key, insttoken="", email="", ncbi_key="", raw_dir="enrichment_raw", pause=0.4):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = f"Galter-Prism-enrichment/1.0 (mailto:{email})" if email else "Galter-Prism-enrichment/1.0"
        self.key, self.inst, self.email, self.ncbi_key = scopus_key, insttoken, email, ncbi_key
        self.raw = Path(raw_dir)
        self.pause = pause
        self.log = []

    def _save(self, sub, name, content):
        d = self.raw / sub
        d.mkdir(parents=True, exist_ok=True)
        (d / name).write_text(content if isinstance(content, str) else json.dumps(content, indent=1), encoding="utf-8")

    def _cached(self, sub, name):
        p = self.raw / sub / name
        return p.read_text(encoding="utf-8") if p.exists() else None

    def _get(self, url, **kw):
        time.sleep(self.pause)
        return self.s.get(url, timeout=60, **kw)

    # Scopus -----------------------------------------------------------------
    def scopus(self, rid, doi, use_cache=True):
        name = f"{rid}.json"
        if use_cache and (c := self._cached("scopus", name)):
            return json.loads(c), "cache"
        if not doi:
            return None, "no DOI"
        h = {"X-ELS-APIKey": self.key, "Accept": "application/json"}
        if self.inst:
            h["X-ELS-Insttoken"] = self.inst
        for view in ("FULL", "META_ABS"):
            r = self._get(f"https://api.elsevier.com/content/abstract/doi/{requests.utils.quote(doi, safe='/')}",
                          headers=h, params={"view": view})
            if r.status_code == 200:
                j = r.json()
                j["_view"] = view
                self._save("scopus", name, j)
                return j, f"ok ({view})"
            if r.status_code == 404:
                return None, "not in Scopus (404)"
            if r.status_code in (401, 403) and view == "FULL":
                continue                       # not entitled to FULL: try the smaller view
            return None, f"HTTP {r.status_code}: {r.text[:120]}"
        return None, "not entitled to FULL or META_ABS view (check VPN / insttoken)"

    # PubMed -----------------------------------------------------------------
    def pmid_for_doi(self, doi):
        params = {"db": "pubmed", "term": f"{doi}[doi]", "retmode": "json"}
        if self.ncbi_key:
            params["api_key"] = self.ncbi_key
        if self.email:
            params["email"] = self.email
        try:
            r = self._get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi", params=params)
            ids = r.json().get("esearchresult", {}).get("idlist", [])
            return ids[0] if len(ids) == 1 else ""
        except Exception:
            return ""

    def pubmed(self, rid, pmid, use_cache=True):
        name = f"{rid}.xml"
        if use_cache and (c := self._cached("pubmed", name)):
            return c, "cache"
        if not pmid:
            return None, "no PMID"
        params = {"db": "pubmed", "id": pmid, "retmode": "xml"}
        if self.ncbi_key:
            params["api_key"] = self.ncbi_key
        if self.email:
            params["email"] = self.email
        r = self._get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi", params=params)
        if r.status_code == 200 and "<PubmedArticle" in r.text:
            self._save("pubmed", name, r.text)
            return r.text, "ok"
        return None, f"HTTP {r.status_code}"

    # Crossref / Unpaywall ---------------------------------------------------
    def crossref(self, rid, doi, use_cache=True):
        name = f"{rid}.json"
        if use_cache and (c := self._cached("crossref", name)):
            return json.loads(c), "cache"
        if not doi:
            return None, "no DOI"
        r = self._get(f"https://api.crossref.org/works/{requests.utils.quote(doi, safe='/')}",
                      params={"mailto": self.email} if self.email else None)
        if r.status_code == 200:
            j = r.json().get("message", {})
            self._save("crossref", name, j)
            return j, "ok"
        return None, f"HTTP {r.status_code}"

    def unpaywall(self, rid, doi, use_cache=True):
        name = f"{rid}.json"
        if use_cache and (c := self._cached("unpaywall", name)):
            return json.loads(c), "cache"
        if not doi or not self.email:
            return None, "no DOI" if not doi else "skipped (CONTACT_EMAIL not set)"
        r = self._get(f"https://api.unpaywall.org/v2/{doi}", params={"email": self.email})
        if r.status_code == 200:
            j = r.json()
            self._save("unpaywall", name, j)
            return j, "ok"
        return None, f"HTTP {r.status_code}"


# ------------------------------------------------------------------ parsers
def parse_scopus(j):
    """-> dict with abstract, authors[], keywords, index_terms, subject_areas, funding[], ids."""
    if not j:
        return {}
    a = j.get("abstracts-retrieval-response", j)
    core = a.get("coredata", {}) or {}
    out = {
        "view": j.get("_view", ""),
        "title": txt(core.get("dc:title")),
        "doi": txt(core.get("prism:doi")),
        "eid": txt(core.get("eid")),
        "pmid": txt(core.get("pubmed-id")),
        "journal": txt(core.get("prism:publicationName")),
        "publisher": txt(core.get("dc:publisher")),
        "cover_date": txt(core.get("prism:coverDate")),
        "volume": txt(core.get("prism:volume")),
        "issue": txt(core.get("prism:issueIdentifier")),
        "pages": txt(core.get("prism:pageRange")),
        "abstract": clean_abstract(txt(core.get("dc:description"))),
    }
    item = a.get("item", {}) or {}
    head = (item.get("bibrecord", {}) or {}).get("head", {}) or {}
    if not out["abstract"]:
        out["abstract"] = clean_abstract(txt((head.get("abstracts"))))

    # Affiliation names from the top-level list (by afid)
    aff_names = {}
    for af in aslist(a.get("affiliation")):
        name = ", ".join(x for x in (txt(af.get("affilname")), txt(af.get("affiliation-city")),
                                     txt(af.get("affiliation-country"))) if x)
        aff_names[str(af.get("@id"))] = name

    # Authors: start from the top-level list (order, names, Scopus IDs, afids)
    authors = {}
    for au in aslist((a.get("authors") or {}).get("author")):
        seq = str(au.get("@seq") or len(authors) + 1)
        pref = au.get("preferred-name") or {}
        authors[seq] = {
            "seq": int(seq) if seq.isdigit() else len(authors) + 1,
            "family": txt(au.get("ce:surname")) or txt(pref.get("ce:surname")),
            "given": txt(au.get("ce:given-name")) or txt(pref.get("ce:given-name")) or txt(au.get("ce:initials")),
            "scopus_id": txt(au.get("@auid")),
            "orcid": "",
            "affiliations": [aff_names.get(str(x.get("@id")), "") for x in aslist(au.get("affiliation")) if isinstance(x, dict)],
        }
    # FULL view: author-groups carry ORCIDs and the full affiliation text as printed in the paper
    for grp in aslist(head.get("author-group")):
        aff = grp.get("affiliation") or {}
        full = txt(aff.get("ce:source-text"))
        if not full:
            parts = [txt(o) for o in aslist(aff.get("organization"))] + [txt(aff.get("city")), txt(aff.get("country"))]
            full = ", ".join(p for p in parts if p)
        for au in aslist(grp.get("author")):
            seq = str(au.get("@seq") or "")
            rec = authors.get(seq) or next((v for v in authors.values() if v["scopus_id"] and v["scopus_id"] == txt(au.get("@auid"))), None)
            if rec is None:
                rec = authors.setdefault(seq or str(len(authors) + 1), {
                    "seq": int(seq) if seq.isdigit() else len(authors) + 1,
                    "family": txt(au.get("ce:surname")), "given": txt(au.get("ce:given-name")) or txt(au.get("ce:initials")),
                    "scopus_id": txt(au.get("@auid")), "orcid": "", "affiliations": []})
            if au.get("@orcid"):
                rec["orcid"] = txt(au.get("@orcid"))
            if full:
                if "_full" not in rec:
                    rec["_full"] = []
                if full not in rec["_full"]:
                    rec["_full"].append(full)
    for rec in authors.values():
        if rec.get("_full"):
            rec["affiliations"] = rec.pop("_full")        # prefer the affiliation text printed in the paper
        rec["affiliations"] = [x for x in dict.fromkeys(rec["affiliations"]) if x]
    out["authors"] = sorted(authors.values(), key=lambda r: r["seq"])

    out["author_keywords"] = [txt(k) for k in aslist((a.get("authkeywords") or {}).get("author-keyword")) if txt(k)]
    out["index_terms"] = [txt(k) for k in aslist((a.get("idxterms") or {}).get("mainterm")) if txt(k)]
    out["subject_areas"] = [txt(k) for k in aslist((a.get("subject-areas") or {}).get("subject-area")) if txt(k)]

    funding = []
    for g in aslist((head.get("grantlist") or {}).get("grant")):
        agency = txt(g.get("grant-agency"))
        for gid in aslist(g.get("grant-id")) or [""]:
            funding.append({"funder": agency, "award": txt(gid), "source": "Scopus grantlist"})
    xocs = (item.get("xocs:meta") or {}).get("xocs:funding-list") or {}
    for f in aslist(xocs.get("xocs:funding")):
        agency = txt(f.get("xocs:funding-agency-matched-string")) or txt(f.get("xocs:funding-agency"))
        for gid in aslist(f.get("xocs:funding-id")) or [""]:
            funding.append({"funder": agency, "award": txt(gid), "source": "Scopus funding list"})
    seen, uniq = set(), []
    for f in funding:
        k = (f["funder"].lower(), f["award"].lower())
        if f["funder"] and k not in seen:
            seen.add(k); uniq.append(f)
    out["funding"] = uniq
    return out


def parse_pubmed(xml_text):
    if not xml_text:
        return {}
    root = ET.fromstring(xml_text)
    art = root.find(".//PubmedArticle")
    if art is None:
        return {}
    mc = art.find("MedlineCitation")
    a = mc.find("Article")
    out = {"pmid": (mc.findtext("PMID") or "").strip(),
           "title": "".join(a.find("ArticleTitle").itertext()).strip() if a.find("ArticleTitle") is not None else ""}
    parts = []
    for t in a.findall("Abstract/AbstractText"):
        body = "".join(t.itertext()).strip()
        label = t.get("Label")
        parts.append(f"{label.title()}: {body}" if label else body)
    out["abstract"] = clean_abstract(" ".join(parts))
    mesh = []
    for mh in mc.findall("MeshHeadingList/MeshHeading"):
        d = mh.find("DescriptorName")
        mesh.append({"descriptor": (d.text or "").strip(), "ui": d.get("UI", ""), "major": d.get("MajorTopicYN") == "Y"
                     or any(q.get("MajorTopicYN") == "Y" for q in mh.findall("QualifierName")),
                     "qualifiers": [(q.text or "").strip() for q in mh.findall("QualifierName")]})
    out["mesh"] = mesh
    out["keywords"] = [(k.text or "").strip() for k in mc.findall("KeywordList/Keyword") if (k.text or "").strip()]
    authors = []
    for i, au in enumerate(a.findall("AuthorList/Author"), 1):
        orcid = next(((x.text or "").strip() for x in au.findall("Identifier") if x.get("Source") == "ORCID"), "")
        orcid = re.sub(r"^https?://orcid\.org/", "", orcid)
        authors.append({"seq": i, "family": (au.findtext("LastName") or au.findtext("CollectiveName") or "").strip(),
                        "given": (au.findtext("ForeName") or au.findtext("Initials") or "").strip(),
                        "orcid": orcid, "scopus_id": "",
                        "affiliations": [(x.text or "").strip() for x in au.findall("AffiliationInfo/Affiliation")]})
    out["authors"] = authors
    out["grants"] = [{"funder": (g.findtext("Agency") or "").strip(), "award": (g.findtext("GrantID") or "").strip(),
                      "source": "PubMed"} for g in a.findall("GrantList/Grant")]
    return out


def parse_crossref(j):
    if not j:
        return {}
    lic = None
    for want in ("vor", "unspecified", "am", "tdm", None):
        cands = [l for l in j.get("license", []) if want is None or l.get("content-version") == want]
        cc = [l for l in cands if "creativecommons.org" in (l.get("URL") or "")]
        if cc:
            lic = cc[0]; break
    if lic is None and j.get("license"):
        lic = next((l for l in j["license"] if l.get("content-version") == "vor"), j["license"][0])
    date = (j.get("published") or j.get("published-print") or j.get("published-online") or j.get("issued") or {}).get("date-parts", [[None]])[0]
    return {
        "title": (j.get("title") or [""])[0],
        "journal": (j.get("container-title") or [""])[0],
        "publisher": j.get("publisher", ""),
        "date": "-".join(f"{p:02d}" if i else str(p) for i, p in enumerate(date) if p),
        "license_url": (lic or {}).get("URL", ""),
        "abstract": clean_abstract(re.sub(r"<[^>]+>", " ", j.get("abstract", "") or "")),
        "funding": [{"funder": f.get("name", ""), "award": aw, "funder_doi": f.get("DOI", ""), "source": "Crossref"}
                    for f in j.get("funder", []) for aw in (f.get("award") or [""])],
        "authors": [{"seq": i, "family": au.get("family", au.get("name", "")), "given": au.get("given", ""),
                     "orcid": re.sub(r"^https?://orcid\.org/", "", au.get("ORCID", "")), "scopus_id": "",
                     "affiliations": [x.get("name", "") for x in au.get("affiliation", []) if x.get("name")]}
                    for i, au in enumerate(j.get("author", []), 1)],
    }


# ------------------------------------------------------------------ main collection
def enrich(articles, client, use_cache=True, progress=print):
    results = []
    for n, art in enumerate(articles, 1):
        rid, doi = art["record_id"], art["doi"]
        rec = {"record_id": rid, "doi": doi, "workbook_title": art["title"], "subjects_now": art["subjects_now"]}
        log = lambda step, status: client.log.append({"record_id": rid, "step": step, "status": status})

        sj, st = client.scopus(rid, doi, use_cache); log("Scopus", st)
        sc = parse_scopus(sj)
        pmid = art["pmid"] or sc.get("pmid") or (client.pmid_for_doi(doi) if doi else "")
        px, pt = client.pubmed(rid, pmid, use_cache); log("PubMed", pt)
        pm = parse_pubmed(px)
        cj, ct = client.crossref(rid, doi, use_cache); log("Crossref", ct)
        cr = parse_crossref(cj)
        uj, ut = client.unpaywall(rid, doi, use_cache); log("Unpaywall", ut)

        rec["in_scopus"] = bool(sc)
        rec["scopus_view"] = sc.get("view", "")
        rec["scopus_eid"] = sc.get("eid", "")
        rec["pmid"] = pm.get("pmid") or pmid
        best_title = sc.get("title") or pm.get("title") or cr.get("title")
        rec["source_title"] = best_title
        rec["title_match"] = title_similarity(art["title"], best_title) if best_title else None

        # Abstract: Scopus, else PubMed, else Crossref
        for src, val in (("Scopus", sc.get("abstract")), ("PubMed", pm.get("abstract")), ("Crossref", cr.get("abstract"))):
            if val:
                rec["abstract"], rec["abstract_source"] = val, src
                break
        else:
            rec["abstract"], rec["abstract_source"] = "", ""

        # Authors: Scopus, else PubMed, else Crossref. Fill missing ORCIDs from PubMed/Crossref by position+surname.
        authors, a_src = (sc.get("authors"), "Scopus") if sc.get("authors") else \
                         (pm.get("authors"), "PubMed") if pm.get("authors") else (cr.get("authors") or [], "Crossref")
        for au in authors:
            for other in (pm.get("authors") or [], cr.get("authors") or []):
                match = next((o for o in other if o["family"].lower() == au["family"].lower()), None)
                if match and not au.get("orcid") and match.get("orcid"):
                    au["orcid"] = match["orcid"]
                if match and not au.get("affiliations") and match.get("affiliations"):
                    au["affiliations"] = match["affiliations"]
        rec["authors"], rec["authors_source"] = authors, a_src

        rec["author_keywords"] = sc.get("author_keywords") or pm.get("keywords") or []
        rec["index_terms"] = sc.get("index_terms", [])
        rec["subject_areas"] = sc.get("subject_areas", [])
        rec["mesh"] = pm.get("mesh", [])

        # License: Crossref (publisher-registered), else Unpaywall
        lic_url, lic_src = cr.get("license_url", ""), "Crossref" if cr.get("license_url") else ""
        spdx, title = map_license(lic_url)
        if not spdx and uj and (uj.get("best_oa_location") or {}).get("license"):
            ul = uj["best_oa_location"]["license"]
            spdx, title = map_license(ul)
            if spdx or not lic_url:
                lic_url, lic_src = (lic_url or ul), "Unpaywall"
        rec.update({"license_id": spdx, "license_title": title, "license_url": lic_url, "license_source": lic_src,
                    "oa_status": (uj or {}).get("oa_status", "")})

        # Funding (kept separate; never required for upload)
        funding = sc.get("funding") or []
        funding += [f for f in (pm.get("grants") or []) + (cr.get("funding") or [])
                    if f["funder"] and (f["funder"].lower(), (f["award"] or "").lower())
                    not in {(x["funder"].lower(), (x["award"] or "").lower()) for x in funding}]
        rec["funding"] = funding

        rec["journal"] = sc.get("journal") or cr.get("journal")
        rec["publisher"] = sc.get("publisher") or cr.get("publisher")
        rec["pub_date"] = sc.get("cover_date") or cr.get("date")
        rec["volume"], rec["issue"], rec["pages"] = sc.get("volume", ""), sc.get("issue", ""), sc.get("pages", "")
        results.append(rec)
        progress(f"[{n:>2}/{len(articles)}] {rid}  Scopus: {st:<28} PubMed: {pt:<10} Crossref: {ct:<6} "
                 f"authors:{len(authors)} MeSH:{len(rec['mesh'])} license:{spdx or '-'}")
    return results


# ------------------------------------------------------------------ Prism MeSH lookup (read-only)
def prism_subject_lookup(results, base_url, token="", verify=True, progress=print):
    """For each MeSH descriptor, ask Prism's subjects vocabulary how it stores that term.
    Adds rec['mesh'][i]['prism_id'] / ['prism_scheme'] when an exact match is found. Returns schemes seen."""
    s = requests.Session()
    if token:
        s.headers["Authorization"] = f"Bearer {token}"
    terms = sorted({m["descriptor"] for r in results for m in r["mesh"] if m["descriptor"]})
    found, schemes = {}, {}
    for i, t in enumerate(terms, 1):
        try:
            r = s.get(f"{base_url.rstrip('/')}/api/subjects", params={"suggest": t, "size": 20},
                      timeout=30, verify=verify)
            hits = r.json().get("hits", {}).get("hits", []) if r.status_code == 200 else []
        except Exception as e:
            progress(f"   lookup failed for '{t}': {e.__class__.__name__}")
            hits = []
        for h in hits:
            schemes[h.get("scheme", "")] = schemes.get(h.get("scheme", ""), 0) + 1
        exact = [h for h in hits if (h.get("subject") or "").casefold() == t.casefold()]
        mesh_exact = [h for h in exact if "mesh" in (h.get("scheme") or "").lower()] or exact
        if mesh_exact:
            found[t] = (mesh_exact[0].get("id", ""), mesh_exact[0].get("scheme", ""))
        if i % 25 == 0:
            progress(f"   looked up {i}/{len(terms)} terms")
        time.sleep(0.1)
    for r in results:
        for m in r["mesh"]:
            m["prism_id"], m["prism_scheme"] = found.get(m["descriptor"], ("", ""))
    progress(f"Prism matched {len(found)} of {len(terms)} distinct MeSH headings. Schemes seen: {schemes}")
    return schemes


# ------------------------------------------------------------------ output
def save_excel(results, log, path):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    hdr = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="4E2A84")

    def sheet(name, cols, rows, widths=None):
        ws = wb.create_sheet(name)
        for c, h in enumerate(cols, 1):
            cell = ws.cell(1, c, h); cell.font = hdr; cell.fill = fill
        for r, row in enumerate(rows, 2):
            for c, v in enumerate(row, 1):
                ws.cell(r, c, v).alignment = Alignment(wrap_text=False, vertical="top")
        for c, w in enumerate(widths or [], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        ws.freeze_panes = "B2"
        return ws

    wb.remove(wb.active)
    sheet("Articles",
          ["record_id", "doi", "in_scopus", "scopus_view", "scopus_eid", "pmid", "workbook_title", "source_title",
           "title_match", "journal", "publisher", "pub_date", "volume", "issue", "pages", "abstract", "abstract_source",
           "n_authors", "authors_source", "author_keywords", "mesh_headings", "index_terms", "subject_areas",
           "license_id", "license_title", "license_url", "license_source", "oa_status", "n_funding", "subjects_now"],
          [[r["record_id"], r["doi"], r["in_scopus"], r["scopus_view"], r["scopus_eid"], r["pmid"], r["workbook_title"],
            r["source_title"], r["title_match"], r["journal"], r["publisher"], r["pub_date"], r["volume"], r["issue"],
            r["pages"], r["abstract"], r["abstract_source"], len(r["authors"]), r["authors_source"],
            "; ".join(r["author_keywords"]), "; ".join(m["descriptor"] for m in r["mesh"]), "; ".join(r["index_terms"]),
            "; ".join(r["subject_areas"]), r["license_id"], r["license_title"], r["license_url"], r["license_source"],
            r["oa_status"], len(r["funding"]), r["subjects_now"]] for r in results],
          [11, 28, 8, 9, 18, 10, 40, 40, 8, 25, 20, 11, 6, 6, 9, 60, 9, 8, 9, 40, 50, 40, 25, 14, 30, 40, 10, 9, 8, 40])
    sheet("Authors", ["record_id", "seq", "family", "given", "orcid", "scopus_author_id", "affiliations", "source"],
          [[r["record_id"], a["seq"], a["family"], a["given"], a.get("orcid", ""), a.get("scopus_id", ""),
            " | ".join(a.get("affiliations") or []), r["authors_source"]] for r in results for a in r["authors"]],
          [11, 5, 18, 18, 20, 14, 90, 9])
    sheet("MeSH", ["record_id", "descriptor", "descriptor_ui", "qualifiers", "major_topic", "prism_subject_id", "prism_scheme"],
          [[r["record_id"], m["descriptor"], m["ui"], "; ".join(m["qualifiers"]), m["major"],
            m.get("prism_id", ""), m.get("prism_scheme", "")] for r in results for m in r["mesh"]],
          [11, 40, 12, 40, 8, 45, 12])
    sheet("Author keywords", ["record_id", "keyword"],
          [[r["record_id"], k] for r in results for k in r["author_keywords"]], [11, 50])
    sheet("Funding (optional)", ["record_id", "funder", "award_number", "funder_doi", "source"],
          [[r["record_id"], f["funder"], f.get("award", ""), f.get("funder_doi", ""), f["source"]]
           for r in results for f in r["funding"]], [11, 50, 22, 35, 18])
    sheet("Log", ["record_id", "step", "status"], [[x["record_id"], x["step"], x["status"]] for x in log], [11, 10, 60])
    wb.save(path)
