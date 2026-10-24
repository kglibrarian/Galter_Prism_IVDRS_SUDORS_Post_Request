"""One-time move of an existing project folder into the data/ and output/ layout used by the notebooks.

    python organize_folders.py          # shows what would move (changes nothing)
    python organize_folders.py --move   # moves it

Nothing is deleted or overwritten: if something already exists at the destination it is left where it is and reported.
Run it from the project folder, with Jupyter closed (Windows won't move open files).
"""
import shutil, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORKBOOK_NAME = "2026_10-22 IVDRS_SUDORS_Reconciled_Deposit_List.xlsx"

PLAN = [
    # (what, destination folder)
    (WORKBOOK_NAME, "data"),
    ("files", "data"),
    ("canva_pages", "data"),
    ("download_log.csv", "output/logs"),
    ("download_log_test.csv", "output/logs"),
    ("download_log_not_open_access.csv", "output/logs"),
    ("files_status.csv", "output/logs"),
    ("png_backups", "output/captures"),
    ("test_capture", "output/captures"),
    ("article_enrichment.xlsx", "output/enrichment"),
    ("enrichment_raw", "output/enrichment"),
]
# Prism outputs that used to sit directly in output/
PRISM_PATTERNS = ["staging_*", "production_*", "IVDRS_SUDORS_Prism_Records.xlsx"]   # incl. *_deposit_log.csv


def merge_logs(src, dst, move, rel_s, rel_d):
    """A deposit log exists in both places (a notebook started a new one in output/prism/, or an earlier
    move copied the file but couldn't remove the original). Keep every line once, oldest first."""
    old = src.read_text(encoding="utf-8").splitlines(keepends=True)
    new = dst.read_text(encoding="utf-8").splitlines(keepends=True)
    if new[:len(old)] == old:                       # dst already contains everything in src
        merged = new
    else:
        merged = old + [l for l in new[1:] if l not in set(old)]
    if not move:
        print(f"  would merge {rel_s}  into  {rel_d}  ({len(merged) - 1} lines)")
        return
    dst.write_text("".join(merged), encoding="utf-8")
    src.unlink()
    print(f"  merged {rel_s}  into  {rel_d}  ({len(merged) - 1} lines)")


def main(move):
    todo = []
    for name, dest in PLAN:
        src = ROOT / name
        if src.exists():
            todo.append((src, ROOT / dest / name))
    for pat in PRISM_PATTERNS:
        for src in sorted((ROOT / "output").glob(pat)):
            if src.is_file():
                todo.append((src, ROOT / "output" / "prism" / src.name))
    if not todo:
        print("Nothing to move: the folder already uses the data/ and output/ layout.")
        return
    for src, dst in todo:
        rel_s, rel_d = src.relative_to(ROOT), dst.relative_to(ROOT)
        if dst.is_dir() and src.is_dir() and not any(dst.iterdir()):
            if move: dst.rmdir()          # an empty folder made by paths.py: replace it
        elif src.is_file() and dst.is_file() and src.name.endswith("_deposit_log.csv"):
            merge_logs(src, dst, move, rel_s, rel_d)
            continue
        elif dst.exists():
            print(f"  SKIP  {rel_s}  ->  {rel_d}  (already exists there; left as is)")
            continue
        if move:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            print(f"  moved {rel_s}  ->  {rel_d}")
        else:
            print(f"  would move {rel_s}  ->  {rel_d}")
    leftovers = sorted(p.name for p in (ROOT / "output").glob("*.tmp")) if (ROOT / "output").exists() else []
    if leftovers:
        print("\nLeftover temporary files in output/ (safe to delete):", ", ".join(leftovers))
    if not move:
        print("\nNothing changed. Run again with --move to do it.")


if __name__ == "__main__":
    main("--move" in sys.argv)
