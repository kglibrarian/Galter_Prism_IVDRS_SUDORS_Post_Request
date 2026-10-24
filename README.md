# Galter Prism IVDRS & SUDORS Post Request

Notebooks and scripts used by Galter Health Sciences Library to deposit the reports, data briefs, fact sheets,
presentation and open access journal articles of the **Illinois Violent Death Reporting System (IVDRS)** and the
**Illinois State Unintentional Drug Overdose Reporting System (SUDORS)** into
[Prism](https://prism.northwestern.edu), Northwestern's InvenioRDM repository.

Everything goes into one parent community, *Illinois Violent Death and Overdose Reporting Systems (IVDRS & SUDORS)*.
Each record's keywords include its proposed child community (*IVDRS Reports and Articles*, *SUDORS Reports and
Articles*, *Peer-reviewed Publications*) so the child communities can be created and filled later.

## What's in the repository

| File | What it does |
|---|---|
| `01_Download_Files.ipynb` + `download_files.py` | Downloads every PDF listed on the deposit workbook's **Files** tab (direct links, Google Drive, publisher pages), falling back to a headless browser. Writes `output/logs/download_log.csv`. |
| `02_Capture_Web_Reports.ipynb` + `capture_web_reports.py` | Saves Infogram reports as PDFs, including every tab and dropdown view and slideshow pages. Cell 7b re-captures only the records you list. |
| `03_Check_Files_Status.ipynb` | Read-only check of the `data/files/` folder against the Files tab: missing, unreadable or suspicious files, near-misses in file names, and extra files. Also moves any PNG backups out of `data/files/`. |
| `04_Enrich_Articles.ipynb` + `enrich_articles.py` | Collects article metadata from Scopus, PubMed, Crossref and Unpaywall: publisher abstracts, every author with affiliations and ORCID, author keywords, MeSH headings (with the ids Prism uses), licenses and funding. Writes `output/enrichment/article_enrichment.xlsx`. |
| `05_Prism_Staging_Deposit.ipynb` + `prism_deposit.py` | Deposits the workbook to **Prism staging**: checks vocabularies against Prism, creates the community, invites members, creates drafts, uploads files and publishes into the community. |
| `06_Prism_Production_Deposit.ipynb` | The same for **Prism production**, with an extra typed confirmation on every step that changes Prism. |
| `07_Prism_Community_Report.ipynb` | Read-only: lists every published record in the community into a spreadsheet, one tab per child community, with the DOI to cite (the parent DOI for reports, the publisher's DOI for articles). |
| `config.example.py` | Template for your keys and tokens. |
| `paths.py` | Where the notebooks read from (`data/`) and write to (`output/`). |
| `organize_folders.py` | One-time move of an older, flat project folder into `data/` and `output/`. |

Run the notebooks in number order. Each notebook explains its cells at the top.

## Setup

1. Install Python 3.10 or later, then the packages:
   ```
   python -m pip install -r requirements.txt
   python -m playwright install chromium
   ```
2. Copy `config.example.py` to `config.py` and fill in your keys and tokens. **`config.py` is ignored by git; never commit it.**
   The notebooks read it without printing or saving any value.
3. Put the deposit workbook and the PDFs in `data/` (see **Folders** below). Everything the notebooks write goes to
   `output/`. All folder locations are set in one place, `paths.py`.

## Folders

```
config.py                  your keys (not committed)
paths.py                   every folder location, used by all notebooks
data/                      inputs, kept locally (not committed)
  2026_10-22 IVDRS_SUDORS_Reconciled_Deposit_List.xlsx
  files/                   PDFs: ivdrs_reports/, sudors_reports/, articles/, not_open_access/
  canva_pages/             pages downloaded by hand from Canva
output/                    everything the notebooks write (not committed)
  logs/                    download logs (01), files_status.csv (03)
  captures/                PNG backups and test captures (02)
  enrichment/              article_enrichment.xlsx and cached API responses (04)
  prism/                   payloads, registries, deposit logs and status (05, 06), records report (07)
```

**Moving an older project folder into this layout:** close Jupyter, then run `python organize_folders.py` to see
what would move, and `python organize_folders.py --move` to move it. It never deletes or overwrites anything.

## The deposit workbook

The notebooks read the reconciled deposit workbook directly. Its tabs:

- **Metadata for Deposit**: one row per record (Prism template columns plus staff columns). Only rows with
  Action = Deposit are sent.
- **Creators**: one row per article author (ORCIDs are sent; the Scopus author ID and source columns are staff-only and are **not** sent, because Prism holds ORCIDs only).
- **Subjects**: every keyword and controlled subject, with the exact entry Prism stores (controlled terms by id, e.g.
  `{"id": "https://id.nlm.nih.gov/mesh/D013405"}`; keywords as `{"subject": "..."}`).
- **Files**: one row per file, with its path inside `data/files/`.
- **Funding (optional)**: reference only; never sent to Prism.
- **Community Members**: people to invite to the community, with their role.

## Deposit rules used for this collection

- **Reports and the presentation:** Prism mints the DOI. The publisher is
  *Prism. Galter Health Sciences Library. Northwestern University*.
- **Journal articles:** only gold or hybrid open access is deposited. Each keeps the publisher's DOI (sent as an
  external DOI so Prism does not mint a second one) and the publisher's license.
- **Lurie-era IVDRS reports:** IVDRS is the creator, and Lurie Children's is an organizational contributor
  (role: hosting institution).
- **Contributor roles:** Prism requires a role on every contributor.
- **Resource types:** these use Prism's ids: `text-report`, `article-journal_article`, `learning-presentation`. Bare
  category ids such as `article` are rejected.
- **Citing records:** cite the parent DOI. It always resolves to the latest version, while Prism's "Cite as" box
  shows the DOI of one version.

## Safety

- **No accidental changes:** every notebook cell that changes Prism has its own True/False switch, and all of them
  are off by default, so *Run All* changes nothing. In production, each of those steps also needs
  `CONFIRM_PRODUCTION = "DEPOSIT TO PRODUCTION"`.
- **Safe to re-run:**
  - `output/prism/<env>_registry.json` records each record's Prism draft id.
  - An existing draft is updated rather than duplicated.
  - Files already uploaded with the same checksum are skipped.
  - Published records are never changed by these notebooks. New versions are a separate step.
- **Keep the registry:** keep your local `output/prism/production_registry.json`. It links each `record_id` to its Prism
  record and is needed for any later update.

## Not in the repository

`.gitignore` keeps `config.py` and the contents of `data/` and `output/` out of git (the two folders themselves are
committed, empty). It also ignores workbooks, CSVs, PDFs and PNGs anywhere, Python and Jupyter caches, and leftover
`.tmp` files.
