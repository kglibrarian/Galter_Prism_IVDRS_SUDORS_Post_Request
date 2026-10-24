"""
capture_web_reports.py
Save interactive web reports (Infogram, and optionally Looker Studio / Canva) as full-page PDFs
that keep every chart and graphic, laid out as they look on screen.

Why the normal "Print to PDF" looks bad: Infogram charts are drawn by JavaScript only when they
scroll into view, and the site's print styles reflow/hide parts of the page. This script loads
the page in a real browser, scrolls through it so every chart renders, switches to *screen*
styling, and saves ONE tall PDF page sized to the whole report (plus a PNG backup).

Tabbed charts: many Infogram charts have tabs (or a dropdown) that switch between views of the
data. The script finds those controls, clicks each tab / picks each option, and adds a labelled
page for every view after the main page, so the final PDF is:
    page 1      the whole report as it first appears
    pages 2..n  "View: <tab name>" for each tab/option of each tabbed chart
Run with --inspect first to see which tabs it finds on a page without saving anything.

Infogram links are captured from Infogram's embed version (e.infogram.com/...?src=embed), which is
the report alone, without Infogram's own header, template gallery and footer. A web page that just
wraps one Infogram report (e.g. some sites.northwestern.edu/ivdrs pages) is captured the same way.
Slideshow reports (Previous/Next page arrows) are captured slide by slide, one PDF page per slide. Every
scrolling capture is checked: if any Infogram page's text is missing from the PDF, it is redone page by page.

SETUP (once, in a terminal / VS Code terminal):
    pip install playwright openpyxl pillow pypdf requests
    python -m playwright install chromium

USE
  One page:
    python capture_web_reports.py --url "https://infogram.com/..." --out "Suicide_Deaths_Females.pdf"

  Rows from a CSV you saved from the Files tab (any rows; Link type column optional):
    python capture_web_reports.py --csv infogram_rows.csv --files-dir files

  Every Infogram row on the workbook's Files tab (saves to <files_dir>/<file_name>):
    python capture_web_reports.py --workbook "2026_10-22 IVDRS_SUDORS_Reconciled_Deposit_List.xlsx" --files-dir files

  Add --include-looker-canva to also try Looker Studio and Canva rows (results vary; check them).
  Add --letter to ALSO write a version split into US Letter pages (for printing).
  Existing files are skipped unless you add --overwrite.
  --inspect          list the tabs/dropdowns found on each page; saves nothing
  --no-views         skip clicking through tabs (main page only)
  --tab-selector CSS use your own CSS selector for one set of tab buttons if detection misses them
  --only IVSU-0001 IVSU-0010   just these records (with --workbook/--csv); add --overwrite to redo them
"""
import argparse
import sys
import time
from pathlib import Path

import io

from playwright.sync_api import sync_playwright

from urllib.parse import urlparse

# --------------------------------------------------------------------------------------------
# Infogram specifics (checked on the live IVDRS pages, Oct 2026):
#  * Tabs are  div.igc-tab[role="tab"]  inside  div.igc-tab-texts[role="tablist"]; on infogram.com
#    they sit inside a shadow DOM (Playwright's locators see into it; plain querySelector does not).
#  * Dropdown views are  select.igc-tab-select ; options are picked by position (their values repeat).
#  * Each chart lives in  .InfographicEditor-Contents-Item ; blocks load a second or two after they
#    scroll into view (a spinner shows until then).
#  * The embed version  https://e.infogram.com/<slug>?src=embed  is the report alone, without
#    Infogram's own header, template gallery and footer.
# --------------------------------------------------------------------------------------------

def to_embed(url):
    """infogram.com/<slug> -> e.infogram.com/<slug>?src=embed (report only). Other URLs unchanged."""
    u = urlparse(url)
    if u.netloc.lower() in ("infogram.com", "www.infogram.com"):
        slug = u.path.strip("/")
        if slug and "/" not in slug and slug != "app" and not u.fragment:
            return f"https://e.infogram.com/{slug}?src=embed"
    return url


def find_embedded_infogram(page, url):
    """If a normal web page (e.g. sites.northwestern.edu) just wraps ONE Infogram report in an iframe,
    return that report's embed URL so it can be captured on its own, with tabs. Otherwise None."""
    if "infogram.com" in urlparse(url).netloc:
        return None
    try:
        page.goto(url, wait_until="load", timeout=90_000)
        page.wait_for_timeout(3000)
        srcs = page.evaluate("[...document.querySelectorAll('iframe')].map(f => f.src || '')"
                             ".filter(s => /(^|\\.)e\\.infogram\\.com\\//.test(s.replace(/^https?:\\/\\//, '')))")
    except Exception:
        return None
    srcs = list(dict.fromkeys(s.split("?")[0] for s in srcs))
    if len(srcs) == 1:
        return srcs[0] + "?src=embed"
    if len(srcs) > 1:
        print(f"      note: page embeds {len(srcs)} Infogram reports; capturing the page as a whole")
    return None


def resolve_url(page, url):
    """URL actually captured: Infogram's embed version, or the Infogram report inside a wrapper page."""
    emb = to_embed(url)
    if emb != url:
        return emb
    inner = find_embedded_infogram(page, url)
    if inner:
        print(f"      page wraps an Infogram report; capturing it directly: {inner}")
        return inner
    return url


# Shadow-DOM-aware helpers run in the page.
ALL_ELS_JS = r"""const allEls = root => { const r = []; const walk = n => { for (const c of n.querySelectorAll('*'))
  { r.push(c); if (c.shadowRoot) walk(c.shadowRoot); } }; walk(root); return r; };
const deepText = el => { let t = (el.innerText || '');
  allEls(el).forEach(c => { if (c.shadowRoot) t += ' ' + (c.shadowRoot.textContent || ''); }); return t.trim(); };
const deepRich = el => { const all = [el, ...allEls(el)];
  return all.some(c => c.tagName === 'IMG' || c.tagName === 'CANVAS' ||
         (c.tagName && c.tagName.toLowerCase() === 'svg' && allEls(c).length > 10)); };"""

DEEP_TEXT_JS = "(el) => {" + ALL_ELS_JS + " return deepText(el); }"

# From a control, find the chart block it belongs to and tag it so a locator can find it again.
MARK_BOX_JS = r"""(el, key) => {
  const up = x => x.parentElement || (x.getRootNode && x.getRootNode().host) || null;
  let c = el;
  while (c && !(c.classList && c.classList.contains('InfographicEditor-Contents-Item'))) c = up(c);
  if (!c) { c = el; while (c) { const r = c.getBoundingClientRect(); if (r.height >= 250 && r.width >= 300) break; c = up(c); } }
  if (!c) c = el.parentElement;
  c.setAttribute('data-cap-box', key); return true; }"""


def find_views(page, tab_selector=None, root=None):
    """Returns [{'kind': 'tabs'|'dropdown', 'n': index, 'labels': [...]}, ...] within root (default: whole page)."""
    root = root or page
    out = []
    if tab_selector:
        loc = root.locator(tab_selector)
        if loc.count() > 1:
            out.append({"kind": "tabs", "n": 0, "labels": [t.strip() for t in loc.all_inner_texts()], "custom": True})
    else:
        lists = root.locator('[role="tablist"]')
        for gi in range(lists.count()):
            tabs = lists.nth(gi).locator('[role="tab"]')
            if tabs.count() > 1:
                out.append({"kind": "tabs", "n": gi, "labels": [t.strip() for t in tabs.all_inner_texts()]})
    sels = root.locator("select.igc-tab-select")
    for si in range(sels.count()):
        labels = [t.strip() for t in sels.nth(si).locator("option").all_inner_texts()]
        if len(labels) > 1:
            out.append({"kind": "dropdown", "n": si, "labels": labels})

    # Number charts top-to-bottom as they appear in the report.
    def ypos(v):
        loc = (root.locator(tab_selector).first if v.get("custom") else
               root.locator('[role="tablist"]').nth(v["n"]) if v["kind"] == "tabs" else
               root.locator("select.igc-tab-select").nth(v["n"]))
        try:
            return loc.evaluate("e => e.getBoundingClientRect().top + window.scrollY")
        except Exception:
            return 1e9
    out.sort(key=ypos)
    return out


def _box_text(box):
    try:
        return box.evaluate(DEEP_TEXT_JS, timeout=3000)
    except Exception:
        return ""


def _wait_for_change(page, box, before, max_s=8.0):
    """Wait until the chart's content changes (or max_s passes), then let animations finish."""
    waited = 0.0
    while waited < max_s:
        page.wait_for_timeout(250); waited += 0.25
        if _box_text(box) != before:
            break
    page.wait_for_timeout(1500)


def capture_views(page, tab_selector=None, root=None, start=1):
    """Click every tab / choose every dropdown option; return [(label, png_bytes), ...]."""
    from PIL import Image  # noqa: F401  (checked early so a missing package fails clearly)
    root = root or page
    shots = []
    for k, v in enumerate(find_views(page, tab_selector, root), start):
        key = f"v{k}"
        if v["kind"] == "tabs":
            items = root.locator(tab_selector) if v.get("custom") else \
                root.locator('[role="tablist"]').nth(v["n"]).locator('[role="tab"]')
            first = items.nth(0)
        else:
            items = None
            first = root.locator("select.igc-tab-select").nth(v["n"])
        try:
            first.scroll_into_view_if_needed(timeout=10000)
            first.evaluate(MARK_BOX_JS, key)
        except Exception as e:
            print(f"        (could not locate chart {k}: {e})"); continue
        box = page.locator(f'[data-cap-box="{key}"]').first
        for i, label in enumerate(v["labels"]):
            try:
                before = _box_text(box)
                if v["kind"] == "tabs":
                    items.nth(i).click(timeout=8000)
                else:
                    root.locator("select.igc-tab-select").nth(v["n"]).select_option(index=i, timeout=8000)
                _wait_for_change(page, box, before)
                box.scroll_into_view_if_needed(timeout=5000)
                page.wait_for_timeout(300)
                kind = "tab" if v["kind"] == "tabs" else "option"
                shots.append((f"Chart {k}, {kind}: {label}", box.screenshot(timeout=15000)))
            except Exception as e:
                print(f"        (could not capture '{label}': {e})")
        # put the chart back on its first view
        try:
            if v["kind"] == "tabs":
                items.nth(0).click(timeout=5000)
            else:
                root.locator("select.igc-tab-select").nth(v["n"]).select_option(index=0, timeout=5000)
        except Exception:
            pass
    return shots


def views_to_pdf_pages(shots):
    """Turn (label, png) pairs into one PDF (bytes): each view on its own page with a caption."""
    from PIL import Image, ImageDraw, ImageFont
    pages = []
    for label, png in shots:
        img = Image.open(io.BytesIO(png)).convert("RGB")
        cap_h = 70
        canvas = Image.new("RGB", (img.width + 80, img.height + cap_h + 40), "white")
        canvas.paste(img, (40, cap_h))
        d = ImageDraw.Draw(canvas)
        font = None
        for name in ("arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "Helvetica.ttc"):
            try:
                font = ImageFont.truetype(name, 32); break
            except Exception:
                pass
        if font is None:
            try:
                font = ImageFont.load_default(size=32)   # Pillow 10.1+
            except TypeError:
                font = ImageFont.load_default()
        d.text((40, 20), "View: " + label, fill="black", font=font)
        pages.append(canvas)
    buf = io.BytesIO()
    pages[0].save(buf, format="PDF", save_all=True, append_images=pages[1:], resolution=144)
    return buf.getvalue()


def append_pdf(main_pdf: Path, extra_pdf_bytes: bytes):
    from pypdf import PdfReader, PdfWriter
    w = PdfWriter()
    for src in (PdfReader(str(main_pdf)), PdfReader(io.BytesIO(extra_pdf_bytes))):
        for pg in src.pages:
            w.add_page(pg)
    tmp = main_pdf.with_suffix(".tmp.pdf")
    with open(tmp, "wb") as f:
        w.write(f)
    tmp.replace(main_pdf)

VIEW_WIDTH = 1200          # px; wide enough for Infogram's desktop layout
SCROLL_STEP = 400          # px per scroll step
SCROLL_PAUSE = 1000        # ms between steps: Infogram blocks take ~1-2 s to load once on screen
SETTLE = 2500              # ms after scrolling, for the last charts/animations


# Infogram reports are a stack of "pages" (.InfographicEditor-Contents > div). Infogram only keeps pages near
# the screen loaded: far-away pages are emptied and show a spinner. So we make the window as tall as the whole
# report (every page is "on screen" at once), then check each page has content and no spinner is turning.
PAGE_STATUS_JS = r"""() => {
  const blocks = [...document.querySelectorAll('.InfographicEditor-Contents > div')];
  const loaded = blocks.map(b => ((b.innerText || '').trim().length > 0) || !!b.querySelector('img, canvas, svg'));
  const spinners = document.getAnimations().filter(a => { try { return a.effect.getTiming().iterations === Infinity; }
                                                        catch (e) { return false; } }).length;
  return {n: blocks.length, loaded: loaded, spinners: spinners}; }"""
MAX_VIEWPORT = 32000       # px; Chromium handles windows this tall


def _scroll_through(page):
    last_h = -1
    for _ in range(4):
        h = page.evaluate("document.documentElement.scrollHeight")
        y = 0
        while y < h:
            page.evaluate(f"window.scrollTo(0, {y})")
            page.wait_for_timeout(SCROLL_PAUSE)
            y += SCROLL_STEP
        if h == last_h:
            break
        last_h = h


NEXT_BTN = 'button[aria-label="Next page"]'
PREV_BTN = 'button[aria-label="Previous page"]'


def load_and_render(page, url):
    """Load the page with every part rendered. Returns a mode:
       "ok"       everything loaded in one tall window (or not an Infogram report)
       "fallback" Infogram scrolling report whose pages won't all load at once -> capture page by page
       "slides"   Infogram slideshow report (Previous/Next page arrows) -> capture slide by slide"""
    page.set_viewport_size({"width": VIEW_WIDTH, "height": 900})
    page.goto(url, wait_until="load", timeout=90_000)
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except Exception:
        pass
    page.wait_for_timeout(2000)
    _scroll_through(page)                 # triggers lazy loading on ordinary pages
    page.evaluate("window.scrollTo(0, 0)")

    st = page.evaluate(PAGE_STATUS_JS)
    if not st["n"]:                       # not an Infogram report: done
        page.wait_for_timeout(SETTLE)
        return "ok"
    if page.locator(NEXT_BTN).count() > 0:
        print(f"      Infogram slideshow with {st['n']} page(s): capturing slide by slide")
        return "slides"

    # Make the window as tall as the report so Infogram keeps every page loaded, then wait for them.
    ok = False
    for _ in range(15):
        h = page.evaluate("document.documentElement.scrollHeight")
        if page.viewport_size["height"] < min(h, MAX_VIEWPORT):
            page.set_viewport_size({"width": VIEW_WIDTH, "height": min(h, MAX_VIEWPORT)})
        page.wait_for_timeout(2000)
        st = page.evaluate(PAGE_STATUS_JS)
        if all(st["loaded"]) and st["spinners"] == 0:
            ok = True
            break
    if ok:
        print(f"      all {st['n']} Infogram page(s) loaded")
    else:
        missing = [i + 1 for i, v in enumerate(st["loaded"]) if not v]
        print(f"      Infogram page(s) {missing or '?'} not loaded in a tall window "
              f"({st['spinners']} spinner(s)); will capture page by page instead")
    page.wait_for_timeout(SETTLE)
    return "ok" if ok else "fallback"


def _next_view_number(view_shots):
    return (max(int(l.split(",")[0].split()[1]) for l, _ in view_shots) + 1) if view_shots else 1


# One Infogram page is ready when it has content and no spinner is turning *inside that page*
# (pages that aren't showing keep their own spinners, which don't matter here).
BLOCK_READY_JS = r"""(i) => {
  const b = [...document.querySelectorAll('.InfographicEditor-Contents > div')][i];
  if (!b) return false;
  const loaded = ((b.innerText || '').trim().length > 0) || !!b.querySelector('img, canvas, svg');
  const spinning = document.getAnimations().some(a => { try { return a.effect.getTiming().iterations === Infinity
                     && a.effect.target && b.contains(a.effect.target); } catch (e) { return false; } });
  return loaded && !spinning; }"""


def _wait_page_loaded(page, i, max_s=20):
    for _ in range(max_s):
        page.wait_for_timeout(1000)
        if page.evaluate(BLOCK_READY_JS, i):
            return True
    return False


def _slides(page, views=True, tab_selector=None):
    """Slideshow reports: go to page 1, then click "Next page" through every page, screenshotting each
    (and capturing its tab/dropdown views) while it is showing. Returns (page_pngs, view_shots)."""
    n = page.evaluate("document.querySelectorAll('.InfographicEditor-Contents > div').length")
    prev = page.locator(PREV_BTN).first
    for _ in range(n + 2):                     # back to the first page
        try:
            if prev.is_disabled() or prev.get_attribute("aria-disabled") == "true":
                break
            prev.click(timeout=5000)
            page.wait_for_timeout(800)
        except Exception:
            break
    shots, view_shots = [], []
    for i in range(n):
        if not _wait_page_loaded(page, i):
            print(f"      WARNING: slide {i + 1} still loading; captured anyway")
        page.wait_for_timeout(1000)
        blk = page.locator(".InfographicEditor-Contents > div").nth(i)
        bb = blk.bounding_box()
        if bb and 50 < bb["height"] < 20000:
            shots.append(blk.screenshot(timeout=30000))
        else:
            shots.append(page.screenshot(full_page=True))
        if views:
            view_shots += capture_views(page, tab_selector, root=blk, start=_next_view_number(view_shots))
        if i < n - 1:
            nxt = page.locator(NEXT_BTN).first
            try:
                if nxt.is_disabled() or nxt.get_attribute("aria-disabled") == "true":
                    print(f"      note: 'Next page' stopped at slide {i + 1} of {n}"); break
                nxt.click(timeout=5000)
                page.wait_for_timeout(1500)
            except Exception as e:
                print(f"      could not go to slide {i + 2}: {e}"); break
    return shots, view_shots


# Text of each Infogram page, to check the saved PDF really contains every page.
PAGE_SNIPPETS_JS = r"""() => [...document.querySelectorAll('.InfographicEditor-Contents > div')].map(b =>
  (b.innerText || '').split('\n').map(t => t.trim())
    .filter(t => t.length >= 12 && !/^You have entered page/i.test(t)).slice(0, 3))"""


def _pdf_has_all_pages(page, pdf_path):
    import re
    from pypdf import PdfReader
    norm = lambda t: re.sub(r"[^a-z0-9]", "", t.lower())
    text = norm(" ".join((pg.extract_text() or "") for pg in PdfReader(str(pdf_path)).pages))
    missing = []
    for i, lines in enumerate(page.evaluate(PAGE_SNIPPETS_JS)):
        if lines and not any(norm(l)[:30] in text for l in lines):
            missing.append(i + 1)
    return missing


def _pages_one_by_one(page, views=True, tab_selector=None):
    """Fallback: screenshot each Infogram page while it is on screen, and capture that page's tab views then too.
    Returns (page_pngs, view_shots)."""
    page.set_viewport_size({"width": VIEW_WIDTH, "height": 900})
    shots, view_shots = [], []
    n = page.evaluate("document.querySelectorAll('.InfographicEditor-Contents > div').length")
    for i in range(n):
        blk = page.locator(".InfographicEditor-Contents > div").nth(i)
        blk.scroll_into_view_if_needed(timeout=10000)
        if not _wait_page_loaded(page, i):    # up to ~20 s for this page to fill in
            print(f"      WARNING: page {i + 1} still loading; captured anyway")
        page.wait_for_timeout(1000)
        shots.append(blk.screenshot(timeout=30000))
        if views:
            view_shots += capture_views(page, tab_selector, root=blk, start=_next_view_number(view_shots))
    return shots, view_shots


def images_to_pdf(pngs, out_pdf: Path):
    from PIL import Image
    imgs = [Image.open(io.BytesIO(b)).convert("RGB") for b in pngs]
    imgs[0].save(str(out_pdf), format="PDF", save_all=True, append_images=imgs[1:], resolution=144)


def capture(page, url, out_pdf: Path, letter=False, views=True, tab_selector=None):
    mode = load_and_render(page, resolve_url(page, url))

    # Hide fixed cookie/consent bars and sticky headers that would overlap content.
    page.add_style_tag(content="""
        [class*="cookie" i], [id*="cookie" i], [class*="consent" i], [id*="consent" i] { display: none !important; }
    """)

    page.emulate_media(media="screen")            # use on-screen styles, not print styles
    height = page.evaluate("Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)")
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    fallback_views = None
    if mode == "ok":
        # One tall page = exactly what is on screen, nothing cut at page breaks (text stays selectable).
        page.pdf(path=str(out_pdf), width=f"{VIEW_WIDTH}px", height=f"{height + 40}px",
                 print_background=True, margin={"top": "20px", "bottom": "20px", "left": "0px", "right": "0px"})
        # PNG backup of the full page (handy if a chart did not make it into the PDF).
        page.screenshot(path=str(out_pdf.with_suffix(".png")), full_page=True)
        if letter:
            page.pdf(path=str(out_pdf.with_name(out_pdf.stem + "_letter.pdf")), format="Letter",
                     print_background=True, scale=0.75,
                     margin={"top": "0.4in", "bottom": "0.4in", "left": "0.4in", "right": "0.4in"})
        # Check the PDF really has every Infogram page; if not, redo it page by page.
        missing = _pdf_has_all_pages(page, out_pdf)
        if missing:
            print(f"      PDF is missing Infogram page(s) {missing}; redoing page by page")
            mode = "fallback"
    if mode == "slides":
        pngs, fallback_views = _slides(page, views, tab_selector)
        images_to_pdf(pngs, out_pdf)
        print(f"      saved {len(pngs)} slide(s) (image PDF)")
    elif mode == "fallback":
        # One PDF page per Infogram page, each captured while on screen (images; text not selectable).
        pngs, fallback_views = _pages_one_by_one(page, views, tab_selector)
        images_to_pdf(pngs, out_pdf)
        print(f"      saved {len(pngs)} page(s) page-by-page (image PDF)")

    n_views = 0
    if views:
        shots = fallback_views if fallback_views is not None else capture_views(page, tab_selector)
        if shots:
            append_pdf(out_pdf, views_to_pdf_pages(shots))
            n_views = len(shots)
    return height, n_views


def read_rows(path, sheet="Files"):
    """Rows (as dicts) from the workbook's Files tab, or from a CSV saved from it."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        import csv
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [{k.strip(): (v.strip() if isinstance(v, str) else v) for k, v in r.items() if k}
                    for r in csv.DictReader(f)]
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    head = [c.value for c in ws[1]]
    return [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True)]


def rows_from_workbook(path, include_more, sheet="Files"):
    wanted = {"Infogram web visualization"}
    if include_more:
        wanted |= {"Looker Studio dashboard", "Canva design (view link)"}
    for r in read_rows(path, sheet):
        rid, url, name = r.get("record_id"), r.get("Download from"), r.get("file_name")
        if not rid or not str(rid).startswith("IVSU") or not url or not name:
            continue
        lt = r.get("Link type")
        # A CSV without a 'Link type' column is taken as-is (you already filtered it).
        if "Link type" not in r or lt in wanted:
            yield rid, url, name


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--url"); ap.add_argument("--out")
    ap.add_argument("--workbook"); ap.add_argument("--files-dir", default="files")
    ap.add_argument("--csv", help="CSV saved from the Files tab (needs record_id, file_name, Download from)")
    ap.add_argument("--include-looker-canva", action="store_true")
    ap.add_argument("--letter", action="store_true")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--no-views", action="store_true")
    ap.add_argument("--tab-selector")
    ap.add_argument("--only", nargs="*", help="just these record_ids, e.g. --only IVSU-0001 IVSU-0010")
    a = ap.parse_args()

    if a.url:
        jobs = [("single", a.url, a.out or "capture.pdf")]
        base = Path(".")
    elif a.workbook or a.csv:
        jobs = list(rows_from_workbook(a.csv or a.workbook, a.include_looker_canva))
        if a.only:
            jobs = [j for j in jobs if j[0] in set(a.only)]
        base = Path(a.files_dir)
    else:
        ap.error("give --url (and --out), --workbook, or --csv")

    print(f"{len(jobs)} page(s) to capture")
    ok, failed = 0, []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={"width": VIEW_WIDTH, "height": 900}, device_scale_factor=2)
        page = ctx.new_page()
        for rid, url, name in jobs:
            if a.inspect:
                print(f"  {rid}: {url}")
                mode = load_and_render(page, resolve_url(page, url))
                n_pages = page.evaluate("document.querySelectorAll('.InfographicEditor-Contents > div').length")
                if mode == "ok":
                    views = find_views(page, a.tab_selector)
                else:                                   # look page by page, while each is showing
                    if mode == "fallback":
                        page.set_viewport_size({"width": VIEW_WIDTH, "height": 900})
                    views = []
                    for i in range(n_pages):
                        blk = page.locator(".InfographicEditor-Contents > div").nth(i)
                        if mode == "fallback":
                            blk.scroll_into_view_if_needed(timeout=10000)
                        _wait_page_loaded(page, i)
                        views += find_views(page, a.tab_selector, root=blk)
                        if mode == "slides" and i < n_pages - 1:
                            try:
                                page.locator(NEXT_BTN).first.click(timeout=5000); page.wait_for_timeout(1500)
                            except Exception:
                                break
                if n_pages:
                    print(f"      {n_pages} Infogram page(s), layout: {'slideshow' if mode == 'slides' else 'scrolling'}")
                for k, v in enumerate(views, 1):
                    print(f"      chart {k} ({v['kind']}): " + " | ".join(v["labels"]))
                if not views:
                    print("      no tabs or dropdowns found")
                continue
            out = base / name
            if out.exists() and not a.overwrite:
                print(f"  skip {rid}: {out} exists"); continue
            try:
                h, nv = capture(page, url, out, a.letter, views=not a.no_views, tab_selector=a.tab_selector)
                print(f"  ok   {rid}: {out}  ({h}px tall, {nv} extra view page(s))"); ok += 1
            except Exception as e:
                print(f"  FAIL {rid}: {url}\n        {e}"); failed.append(rid)
        browser.close()
    print(f"done: {ok} saved, {len(failed)} failed {failed if failed else ''}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
