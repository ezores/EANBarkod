"""
EAN‑13 Barcode Generator with SVG and PNG Output
───────────────────────────────────────────────
Supports two modes:
1. Excel Mode: Reads product codes and names from an Excel file
2. Manual Mode: Paste GTINs (with optional descriptions) line by line

✅ PNG: for preview or web
✅ SVG: perfect vector quality for InDesign or print
"""

import re, sys
from pathlib import Path
from datetime import datetime
from barcode import get
from barcode.writer import ImageWriter, SVGWriter
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# ---------- USER SETTINGS --------------------------------------------------

EXCEL_PATH  = r"C:\Users\CycleWSWin\Downloads\barcodes.xlsx"
SHEET_NAME  = "Replaced"
START_ROW, END_ROW = 4, 21

COL_E   = "E"                 # where CONCATENATE result lives
COL_BCD = ("B", "C", "D")     # fallback
DESC_COL = "F"                # product label

OUTPUT_DIR = Path(r"C:\Users\CycleWSWin\Downloads\barcodes")

FONT_PATH, FONT_SIZE = r"C:\Windows\Fonts\arial.ttf", 34
BAR_HEIGHT, SIDE_PAD_MM, GAP_MM, DPI = 22, 2.0, 0.7, 600

# ---------------------------------------------------------------------------

# --- Derived values
MM_TO_PX = DPI / 25.4
side_pad_px = int(SIDE_PAD_MM * MM_TO_PX)
gap_px = int(GAP_MM * MM_TO_PX)
font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
ascent = font.getmetrics()[0]

# --- Timestamped output folder
STAMP = datetime.now().strftime("%Y-%m-%d_%H%M")
BATCH_DIR = OUTPUT_DIR / STAMP
BATCH_DIR.mkdir(parents=True, exist_ok=True)

# ---------- Helpers --------------------------------------------------------

def safe_name(s: str) -> str:
    """Make a string safe to use as a filename."""
    return re.sub(r'[\/:*?"<>|\\]', "_", s).strip().replace(" ", "_")[:150]

def check_digit(payload12: str) -> str:
    """Calculate the EAN‑13 check digit for a 12-digit base."""
    s = sum((3 if i % 2 else 1) * int(d) for i, d in enumerate(reversed(payload12), 1))
    return str((-s) % 10)

# ---------- Barcode Writers ------------------------------------------------

def build_bars_png(payload12: str):
    """Render barcode bars as high-resolution PNG (without number)."""
    opts = {
        "module_width": 0.33,
        "module_height": BAR_HEIGHT,
        "quiet_zone": SIDE_PAD_MM,
        "write_text": False,
        "dpi": DPI,
        "mode": "RGBA",
        "background": (255, 255, 255, 0),
    }
    return get("ean13", payload12, writer=ImageWriter()).render(opts)

def write_png(gtin13: str, label: str):
    """Save PNG image with barcode + number underneath."""
    bars = build_bars_png(gtin13[:12])
    bw, bh = bars.size
    txt_w, txt_h = font.getbbox(gtin13)[2:]
    W = max(bw, txt_w) + side_pad_px * 2
    H = bh + 1 + txt_h  # tighter spacing

    img = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    img.paste(bars, ((W - bw) // 2, 0), bars)

    y = bh + 1  # draw text right below bars
    ImageDraw.Draw(img).text(((W - txt_w) // 2, y), gtin13, font=font, fill=(0, 0, 0, 255))

    fn = safe_name(label) if label else gtin13
    path = BATCH_DIR / f"{fn}.png"
    img.save(path, "PNG")
    return path

def write_svg(gtin13: str, label: str):
    """Save SVG vector barcode (best for InDesign and printing)."""
    barcode = get("ean13", gtin13[:12], writer=SVGWriter())
    fn = safe_name(label) if label else gtin13
    path = BATCH_DIR / f"{fn}.svg"
    barcode.save(str(path.with_suffix("")))  # avoid .svg.svg
    return path

def make_valid(gtin: str) -> str | None:
    """Return a valid 13-digit GTIN (fix or append check digit if needed)."""
    if not gtin.isdigit():
        return None
    if len(gtin) == 12:
        return gtin + check_digit(gtin)
    if len(gtin) == 13:
        return gtin[:12] + check_digit(gtin[:12])
    return None

# ---------- Mode 1: Excel Reader -------------------------------------------

def run_auto():
    if not Path(EXCEL_PATH).is_file():
        print("❌ Excel file not found."); return
    try:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        sheet = wb[SHEET_NAME]
    except Exception as e:
        print(f"❌ Error loading Excel sheet: {e}")
        return

    concat = lambda r: "".join(str(sheet[f"{c}{r}"].value or "").strip() for c in COL_BCD)

    ok = skip = 0
    for r in range(START_ROW, END_ROW + 1):
        raw = str(sheet[f"{COL_E}{r}"].value or "").strip() or concat(r)
        desc = str(sheet[f"{DESC_COL}{r}"].value or "").strip()
        if not raw:
            continue
        gtin = make_valid(raw)
        if not gtin:
            print(f"⚠️ Row {r}: “{raw}” is invalid – skipped")
            skip += 1
            continue
        p1 = write_png(gtin, desc)
        p2 = write_svg(gtin, desc)
        print(f"✅ Row {r}: {gtin} → {p1.name}, {p2.name}")
        ok += 1
    print(f"\n✔ Excel Mode done → {ok} barcodes, {skip} skipped → {BATCH_DIR}")

# ---------- Mode 2: Manual Paste -------------------------------------------

def run_manual():
    print("📥 Paste GTINs (one per line). Add a name with TAB or 2+ spaces.")
    print("🧾 Example: 8684771191031    Pistachio Paste 200g")
    print("🔚 Finish with empty line or Ctrl+Z then Enter.\n")

    ok = bad = 0
    try:
        while True:
            line = sys.stdin.readline()
            if not line or line.strip() == "":
                break
            if "\t" in line:
                code, desc = line.strip().split("\t", 1)
            else:
                parts = re.split(r" {2,}", line.strip(), maxsplit=1)
                code, desc = parts[0], parts[1] if len(parts) > 1 else ""
            gtin = make_valid(code.strip())
            if not gtin:
                print(f" ✗  {code} → invalid (ignored)"); bad += 1; continue
            p1 = write_png(gtin, desc.strip())
            p2 = write_svg(gtin, desc.strip())
            print(f" ✓  {gtin} → {p1.name}, {p2.name}")
            ok += 1
    except KeyboardInterrupt:
        pass
    print(f"\n✔ Manual Mode done → {ok} barcodes, {bad} invalid → {BATCH_DIR}")

# ---------- Main Menu ------------------------------------------------------

choice = input("Mode  [1] Excel automatic  /  [2] Manual bulk‑paste  : ").strip()
if choice == "1":
    run_auto()
elif choice == "2":
    print()
    run_manual()
else:
    print("❌ Invalid choice. Exiting.")
