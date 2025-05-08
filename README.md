EAN‑13 Barcode Generator

This Python script generates **EAN‑13** barcodes in both **PNG** and **SVG** formats.

- PNG: High-resolution preview or web use  
- SVG: Scalable vector format, ideal for **InDesign**, **printing**, and **packaging**

---

## 🛠 Requirements

Create a `requirements.txt` file with the following content:

```
python-barcode
openpyxl
Pillow
```

Then install the packages:

```bash
pip install -r requirements.txt
```

---

Configuration

Open the script (`test.py`) and modify these settings at the top:

```python
EXCEL_PATH  = r"C:\Users\YourName\Downloads\barcodes.xlsx"
SHEET_NAME  = "Replaced"
START_ROW, END_ROW = 4, 104

COL_E    = "E"                 # GTIN (or CONCAT from B+C+D)
COL_BCD  = ("B", "C", "D")     # fallback when E is a formula
DESC_COL = "F"                 # Product description

OUTPUT_DIR = Path(r"C:\Users\YourName\Downloads\barcodes")

FONT_PATH, FONT_SIZE = r"C:\Windows\Fonts\arial.ttf", 34
BAR_HEIGHT, SIDE_PAD_MM, GAP_MM, DPI = 22, 2.0, 0.7, 600
```

---

Running the Script

```bash
python test.py
```

You will be asked to choose:

```
Mode  [1] Excel automatic  /  [2] Manual bulk‑paste  :
```

---

### Option 1: Excel Automatic Mode

- Reads GTINs (12 or 13 digits) from column `E` or `B+C+D` if E is empty.
- Product names from column `F` are used as the filenames.
- Saves barcodes into a timestamped folder:

```
barcodes/YYYY-MM-DD_HHMM/
```

Each GTIN will generate:
- `product_name.svg`
- `product_name.png`

---

### Option 2: Manual Paste Mode

Paste GTINs directly into the terminal like this:

```
8684771191017
8684771191018    Pistachio Paste 5kg
8684771191024    Pistachio Paste 200g
```

- You can add a product name using **TAB** or **two+ spaces** after the number.
- End the input by pressing **Enter on a blank line** or **Ctrl+Z + Enter** (on Windows).

---

Output

Each entry generates:

- `product_name.svg` — vector (perfect for InDesign)
- `product_name.png` — high‑DPI PNG preview

Example console output:

```
✅ Row 5: 8684771191024 → Pistachio_Paste_200g.png, Pistachio_Paste_200g.svg
```

---

Notes

- **SVG**: Best for scalable, print‑quality designs.
- **PNG**: 600 DPI, good for previews or testing.
- Barcode text is centered under the bars.
- Check digits are calculated and corrected automatically.

---

Troubleshooting

| Issue                        | Solution                                 |
|-----------------------------|------------------------------------------|
| Excel file not found        | Check the `EXCEL_PATH`                   |
| Wrong sheet name            | Verify `SHEET_NAME` in the Excel file    |
| Description is blank        | Fallbacks to GTIN for filename           |
| GTIN invalid or skipped     | Must be 12 or 13 digits only             |
| SVG looks blurry in preview | It’s not — it’s vector. Use InDesign.    |

---

Example Use Case

- Create product labels for jars, packaging, bottles, boxes.
- Generate **commercial-ready** barcodes from your product inventory.
- Drag SVGs into your Adobe InDesign or Illustrator designs — they scale perfectly.
