````markdown
# EAN‑13 Barcode Generator

This Python script generates **EAN‑13** barcodes in both **PNG** and **SVG** formats.  
Use **PNG** for quick previews or web, and **SVG** for perfect, scalable vector art (ideal for InDesign).

---

## 🚀 Installation

1. **Clone or copy** this repo into a folder, e.g.:  
   `C:\Users\YourName\Downloads\barcode-generator`

2. **Install dependencies** via `requirements.txt`:

   ```bash
   cd C:\Users\YourName\Downloads\barcode-generator
   pip install -r requirements.txt
````

> This will install:
>
> * `python-barcode` (core barcode generator)
> * `openpyxl`     (Excel reader)
> * `Pillow`       (image drawing)

---

## ⚙️ Configuration

At the top of the script (`test.py`), update these to match your setup:

```python
# Path to your Excel file
EXCEL_PATH  = r"C:\Users\YourName\Downloads\barcodes.xlsx"

# Excel sheet name and row range
SHEET_NAME  = "Replaced"
START_ROW, END_ROW = 4, 104

# Columns in Excel:
COL_E    = "E"          # 12‑ or 13‑digit code or CONCAT(B,C,D)
COL_BCD  = ("B","C","D") # fallback if E is blank
DESC_COL = "F"          # product description → filename

# Output folder for barcode images
OUTPUT_DIR = Path(r"C:\Users\YourName\Downloads\barcodes")

# Font for the numbers under the bars
FONT_PATH, FONT_SIZE = r"C:\Windows\Fonts\arial.ttf", 34

# Barcode size (mm) and resolution (DPI)
BAR_HEIGHT, SIDE_PAD_MM, GAP_MM, DPI = 22, 2.0, 0.7, 600
```

---

## ▶️ Usage

Run the script:

```bash
python test.py
```

You’ll be prompted:

```
Mode  [1] Excel automatic  /  [2] Manual bulk‑paste  :
```

1. **Excel automatic**

   * Reads codes from **column E** (or builds from B+C+D if E is empty).
   * Reads descriptions from **column F** for file names.
   * Saves both `.png` and `.svg` in a timestamped folder:
     `…/barcodes/YYYY-MM-DD_HHMM/`

2. **Manual bulk‑paste**

   * **Copy** a column of GTINs (12 or 13 digits) from Excel and **paste** here.
   * Optionally add a product name after a **TAB** or **two spaces**:

     ```
     8684771191017
     8684771191018    Pistachio Paste 5 kg
     ```
   * Press **Enter** on an empty line (or **Ctrl+Z** then **Enter** on Windows) to finish.

All generated files will be logged and saved in the same timestamped folder.

---

## 📂 Output Example

```
✅ Row  5: 8684771191024 → Pistachio_Paste_200g.png, Pistachio_Paste_200g.svg
...
```

---

## 📝 Notes

* **SVG** files are true vectors—import into InDesign at any size with zero blur.
* **PNG** files are high‑DPI (600 dpi) for preview or web.
* You can tweak **module\_width**, **BAR\_HEIGHT**, **quiet\_zone**, **font\_size**, or **DPI** in the script to match your printer specs.

---

## 🛠 Troubleshooting

* **Excel file not found** → check `EXCEL_PATH`.
* **Wrong sheet** → update `SHEET_NAME`.
* **Missing or wrong GTIN** → ensure all codes are 12 or 13 digits; invalid rows will be skipped.

```
```
