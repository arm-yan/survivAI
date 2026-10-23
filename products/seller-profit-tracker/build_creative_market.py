r"""Package the Creative Market delivery zip.

Usage (from the repo root, after build.py and build_guide.py):
    .venv\Scripts\python.exe products\seller-profit-tracker\build_creative_market.py

Writes dist/creative-market/MarginKite-Seller-Profit-Tracker.zip with exactly the two
tracker workbooks and the Creative Market buyer guide. Entries are sorted and carry fixed
timestamps and permissions, so the same inputs always give the same zip bytes.
"""

import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DIST = HERE / "dist"
CM = DIST / "creative-market"
OUT = CM / "MarginKite-Seller-Profit-Tracker.zip"
CONTENTS = {  # name inside the zip -> source file
    "MarginKite-Buyer-Guide.pdf": CM / "MarginKite-Buyer-Guide.pdf",
    "MarginKite-Seller-Profit-Tracker-BLANK.xlsx": DIST / "MarginKite-Seller-Profit-Tracker-BLANK.xlsx",
    "MarginKite-Seller-Profit-Tracker.xlsx": DIST / "MarginKite-Seller-Profit-Tracker.xlsx",
}
STAMP = (2026, 10, 21, 0, 0, 0)


def main():
    missing = [str(p) for p in CONTENTS.values() if not p.exists()]
    if missing:
        raise SystemExit(f"Missing inputs: {missing}")
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(CONTENTS):
            info = zipfile.ZipInfo(name, date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            info.create_system = 0
            z.writestr(info, CONTENTS[name].read_bytes(), compresslevel=9)
    with zipfile.ZipFile(OUT) as z:
        for i in z.infolist():
            print(f"  {i.filename}  {i.file_size} bytes")
    print(f"Built {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
