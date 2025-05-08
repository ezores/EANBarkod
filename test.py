"""
EAN‑13 barcode generator with batch sub‑folders
----------------------------------------------

• Column I … 12/13‑digit GTIN/EAN (auto‑computes check digit if 12)
• Column F … product description → becomes file name
• Each run writes PNGs to   OUTPUT_DIR / <YYYY‑MM‑DD_HHMM> /
"""

import os, sys, re, csv
from pathlib import Path
from datetime import datetime
from barcode import get
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# ========= USER SETTINGS ===================================================
EXCEL_PATH  = r"C:\Users\DIR.xlsx"
SHEET_NAME  = "Replaced"
START_ROW   = 4
END_ROW     = 104
CODE_COLUMN = "I"
DESC_COLUMN = "F"
OUTPUT_DIR  = Path(r"C:\Users\OUT")  # base folder
# ---------------------------------------------------------------------------
FONT_PATH   = r"C:\Windows\Fonts\arial.ttf"
FONT_SIZE   = 34
BAR_HEIGHT  = 22        # mm
SIDE_PAD_MM = 2.0
GAP_MM      = 0.7
DPI         = 600
# ===========================================================================

# ---- derived constants ----------------------------------------------------
MM_TO_PX   = DPI / 25.4
side_pad_px = int(SIDE_PAD_MM * MM_TO_PX)
gap_px      = int(GAP_MM      * MM_TO_PX)

font   = ImageFont.truetype(FONT_PATH, FONT_SIZE)
ascent = font.getmetrics()[0]

# ---- batch sub‑folder ------------------------------------------------------
STAMP      = datetime.now().strftime("%Y-%m-%d_%H%M")
BATCH_DIR  = OUTPUT_DIR / STAMP
BATCH_DIR.mkdir(parents=True, exist_ok=True)

# ---- helper functions ------------------------------------------------------
def safe_name(label: str) -> str:
    """Filesystem‑safe: replace forbidden chars, trim/underscore spaces."""
    label = re.sub(r'[\/:*?"<>|\\]', "_", label)
    return label.strip().replace(" ", "_")[:150]

def build_bar_image(base12: str):
    """Return a PIL Image with bars only."""
    opts = {
        "module_width": 0.33,
        "module_height": BAR_HEIGHT,
        "quiet_zone": SIDE_PAD_MM,
        "write_text": False,
        "dpi": DPI,
        "mode": "RGBA",
        "background": (255, 255, 255, 0),
    }
    return get("ean13", base12, writer=ImageWriter()).render(opts)

def save_ean(code: str, label: str) -> str:
    """Generate PNG in BATCH_DIR.  Return full 13‑digit GTIN."""
    if len(code) not in (12, 13) or not code.isdigit():
        raise ValueError("code must be 12 or 13 digits")

    base = code[:12]
    full = get("ean13", base).get_fullcode()

    bars = build_bar_image(base)
    bw, bh = bars.size

    text_w, text_h = font.getbbox(full)[2:]
    canvas_w = max(bw, text_w) + side_pad_px * 2
    canvas_h = bh + gap_px + ascent

    canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    canvas.paste(bars, ((canvas_w - bw)//2, 0), bars)

    draw = ImageDraw.Draw(canvas)
    text_x = (canvas_w - text_w)//2
    text_y = bh + gap_px - (ascent - text_h)
    draw.text((text_x, text_y), full, font=font, fill=(0, 0, 0, 255))

    fname = safe_name(label) if label else full
    out_path = BATCH_DIR / f"{fname}.png"
    canvas.save(out_path, "PNG")
    return full, out_path

# ---- main loop -------------------------------------------------------------
if not Path(EXCEL_PATH).is_file():
    sys.exit(f"❌ Excel file not found: {EXCEL_PATH}")

wb    = load_workbook(EXCEL_PATH, data_only=True)
sheet = wb[SHEET_NAME]

ok, skip = 0, 0
for row in range(START_ROW, END_ROW + 1):
    code  = str(sheet[f"{CODE_COLUMN}{row}"].value or "").strip()
    label = str(sheet[f"{DESC_COLUMN}{row}"].value or "").strip()

    if not code:
        continue
    try:
        gtin, path = save_ean(code, label)
        ok += 1
        print(f"✅ Row {row}: {gtin} → {path.relative_to(OUTPUT_DIR)}")
    except Exception as e:
        skip += 1
        print(f"⚠️ Row {row}: {e}")

print(f"\nFinished: {ok} barcodes written to “{BATCH_DIR.name}”, {skip} skipped.")
