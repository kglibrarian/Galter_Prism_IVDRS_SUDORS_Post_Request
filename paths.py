"""Where everything lives. Every notebook imports this, so a folder only has to be changed here.

    data/      inputs you keep locally (not committed)
      2026_10-22 IVDRS_SUDORS_Reconciled_Deposit_List.xlsx   the deposit workbook
      files/         the PDFs to deposit: ivdrs_reports/, sudors_reports/, articles/, not_open_access/
      canva_pages/   pages downloaded by hand from Canva
    output/    everything the notebooks write (not committed)
      logs/          download logs, files_status.csv
      captures/      PNG backups and test captures from 02
      enrichment/    article_enrichment.xlsx and cached API responses (enrichment_raw/)
      prism/         payloads, registries (KEEP production_registry.json), deposit logs, status, records report
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CONFIG_FILE = ROOT / "config.py"

DATA = ROOT / "data"
WORKBOOK = DATA / "2026_10-22 IVDRS_SUDORS_Reconciled_Deposit_List.xlsx"
FILES_DIR = DATA / "files"
CANVA_DIR = DATA / "canva_pages"

OUTPUT = ROOT / "output"
LOGS = OUTPUT / "logs"
CAPTURES = OUTPUT / "captures"
PNG_BACKUPS = CAPTURES / "png_backups"
ENRICHMENT = OUTPUT / "enrichment"
PRISM_OUT = OUTPUT / "prism"

for _d in (DATA, FILES_DIR, LOGS, CAPTURES, ENRICHMENT, PRISM_OUT):
    _d.mkdir(parents=True, exist_ok=True)
