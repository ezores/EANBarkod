"""
Generate transparent EAN‑13 barcodes from column I in an Excel sheet.

• Accepts 12‑ or 13‑digit values (auto‑computes the check‑digit if missing)
• Renders the bars first, then draws the digits underneath on a larger canvas
  so nothing is cut off.
• Writes <full‑13‑digit‑code>.png files into OUTPUT_DIR.
"""

import os, sys
from pathlib import Path
from barcode import get
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# ========= CONFIGURATION ==============================================
EXCEL_PATH = r
SHEET_NAME = "Replaced"
COLUMN     = "I"
START_ROW  = 4                 # inclusive
END_ROW    = 104               # inclusive   (change as needed)
OUTPUT_DIR = r
# ======================================================================

# ---------- rendering parameters --------------------------------------
FONT_PATH   = r"C:\Windows\Fonts\arial.ttf"   # any TrueType font with digits
FONT_SIZE   = 34                              # try 30‑40 first, tweak later
SIDE_PAD_MM = 2.0                             # blank margin each side (mm)
BAR_HEIGHT  = 22                              # bar height (mm) – GS‑1 minimum
DPI         = 600

# pre‑calc millimetres → pixels
MM_TO_PX = DPI / 25.4
side_pad_px = int(SIDE_PAD_MM * MM_TO_PX)
GAP_MM = 0.1
gap_px = int(GAP_MM * MM_TO_PX)     # 0.7 mm gap bars ↔ text            

font = ImageFont.truetype(FONT_PATH, FONT_SIZE)

# ---------- helper: make one PNG --------------------------------------
def make_ean_png(code: str, outdir: str = OUTPUT_DIR) -> str:
    """
    Build a transparent PNG with scan bars + centred digits underneath.
    Returns the full 13‑digit code actually encoded.
    """
    if len(code) not in (12, 13) or not code.isdigit():
        raise ValueError("value must be 12 or 13 digits")

    base = code[:12]
    full = get("ean13", base).get_fullcode()

    # 1) render BAR image directly in memory  ---------------------------
    writer_opts = {
        "module_width" : 0.33,
        "module_height": BAR_HEIGHT,
        "quiet_zone"   : SIDE_PAD_MM,
        "dpi"          : DPI,
        "write_text"   : False,
        "mode"         : "RGBA",
        "background"   : (255, 255, 255, 0),
    }
    bars = get("ean13", base, writer=ImageWriter()).render(writer_opts)
    bw, bh = bars.size

    # 2) prepare final canvas  ------------------------------------------
    text_w, text_h = font.getbbox(full)[2:]
    final_w = max(bw, text_w) + side_pad_px * 2
    final_h = bh + gap_px + text_h
    canvas = Image.new("RGBA", (final_w, final_h), (255, 255, 255, 0))

    # 3) paste bars + text  ---------------------------------------------
    canvas.paste(bars, ((final_w - bw) // 2, 0), bars)
    draw = ImageDraw.Draw(canvas)
    draw.text(((final_w - text_w) // 2, bh + gap_px),
              full, font=font, fill=(0, 0, 0, 255))

    # 4) save final PNG  -------------------------------------------------
    out_path = Path(outdir, f"{full}.png")
    canvas.save(out_path, "PNG")
    return full

# ---------- main loop -------------------------------------------------
if not os.path.isfile(EXCEL_PATH):
    sys.exit(f"❌  Excel file not found: {EXCEL_PATH}")

os.makedirs(OUTPUT_DIR, exist_ok=True)
wb    = load_workbook(EXCEL_PATH, data_only=True)
sheet = wb[SHEET_NAME]

ok, skip = 0, 0
for row in range(START_ROW, END_ROW + 1):
    raw = str(sheet[f"{COLUMN}{row}"].value or "").strip()
    if not raw:
        continue
    try:
        full = make_ean_png(raw)
        ok += 1
        print("✅", full)
    except Exception as e:
        skip += 1
        print(f"⚠️  Row {row}: {e}")

print(f"\nDone: {ok} barcodes created, {skip} skipped.")
