# 📄 Certificate Generator

A Python script that automatically generates personalised PDF certificates in bulk from an Excel spreadsheet, using your own PDF template and custom fonts.

---

## ✨ Features

- Generates one PDF certificate per row in your Excel sheet
- Overlays personalised fields onto any PDF template
- All text is **perfectly centred** on the page using exact font metrics
- Supports any **custom TTF/OTF font** or built-in PDF fonts
- Fully customisable **text colour**, **font size**, and **field positions**
- **Blank cells are skipped** — no `nan` printed on certificates
- **Duplicate names** are handled automatically (`John Smith.pdf`, `John Smith_2.pdf`, etc.)
- Output files are **named after the recipient**

---

## 📋 Requirements

- Python 3.8+
- The following Python packages:

```bash
pip install pandas openpyxl pymupdf
```

---

## 📁 File Structure

```
project/
│
├── generate_certificates.py   # Main script
├── certificate_template.pdf   # Your PDF template
├── certificates_data.xlsx     # Excel sheet with recipient data
└── certificates/              # Output folder (auto-created)
```

---

## 📊 Excel Sheet Format

Your Excel file must contain these column headers (case-insensitive):

| name | designation | job | title | title_2 |
|------|-------------|-----|-------|---------|
| Jane Doe | Software Engineer | Backend Development | Best Performer | Q1 2025 |
| John Smith | Designer | UI/UX | Most Creative | |

- Columns can be in any order
- Blank cells are automatically skipped — no placeholder text will appear

---

## ⚙️ Configuration

Open `generate_certificates.py` and edit the **CONFIG** section at the top:

```python
EXCEL_FILE   = r"path\to\your\data.xlsx"
TEMPLATE_PDF = r"path\to\your\template.pdf"
OUTPUT_DIR   = r"path\to\output\folder"
```

### Field Positions

Each field has a vertical position (`y_frac`) as a fraction of the page height — `0.0` is the top, `1.0` is the bottom. Adjust these to match where you want text to appear on your template:

```python
FIELDS = [
    {"column": "name",        "y_frac": 0.46, "font_size": 26},
    {"column": "designation", "y_frac": 0.49, "font_size": 12},
    {"column": "job",         "y_frac": 0.53, "font_size": 12},
    {"column": "title",       "y_frac": 0.57, "font_size": 12},
    {"column": "title_2",     "y_frac": 0.60, "font_size": 12},
]
```

### Font

Set a path to any `.ttf` or `.otf` font file:

```python
FONT = r"C:\Windows\Fonts\Georgia.ttf"
```

Or use a built-in PyMuPDF font name (no file needed):

| Name | Font |
|------|------|
| `"helv"` | Helvetica |
| `"hebo"` | Helvetica Bold |
| `"tiro"` | Times Roman |
| `"tibd"` | Times Bold |
| `"cour"` | Courier |

### Text Colour

Set colour as an RGB tuple, each value between `0.0` and `1.0`:

```python
TEXT_COLOR = (0.067, 0.294, 0.514)   # steel blue
```

Some examples:

| Colour | Value |
|--------|-------|
| Black  | `(0.0, 0.0, 0.0)` |
| White  | `(1.0, 1.0, 1.0)` |
| Red    | `(0.8, 0.1, 0.1)` |
| Gold   | `(0.85, 0.65, 0.1)` |
| Navy   | `(0.1, 0.1, 0.4)` |

---

## ▶️ Usage

1. Place your `certificate_template.pdf` and `certificates_data.xlsx` in the project folder
2. Edit the CONFIG section in `generate_certificates.py`
3. Run the script:

```bash
python generate_certificates.py
```

4. Certificates will be saved in the output folder, each named after the recipient:

```
certificates/
  Jane Doe.pdf
  John Smith.pdf
  John Smith_2.pdf    ← second certificate for the same person
  John Smith_3.pdf    ← third certificate for the same person
```

---

## 🔧 Common Windows Font Paths

| Font | Path |
|------|------|
| Arial | `C:\Windows\Fonts\arial.ttf` |
| Arial Bold | `C:\Windows\Fonts\arialbd.ttf` |
| Times New Roman | `C:\Windows\Fonts\times.ttf` |
| Georgia | `C:\Windows\Fonts\Georgia.ttf` |
| Calibri | `C:\Windows\Fonts\calibri.ttf` |

---

## 🐛 Troubleshooting

**`FileNotFoundError: Template PDF not found`**
→ Check the path in `TEMPLATE_PDF`. If the filename shows as `oseal.pdf.pdf` in File Explorer, Windows is hiding the extension — rename it to remove the duplicate.

**`ValueError: Excel is missing columns`**
→ Make sure your Excel headers are exactly: `name`, `designation`, `job`, `title`, `title_2` (case-insensitive).

**Text appears slightly off-centre**
→ Adjust the `y_frac` values in `FIELDS` to fine-tune vertical positions.

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Read Excel data |
| `openpyxl` | Excel file engine for pandas |
| `pymupdf` | PDF manipulation and text rendering |

---

## 📝 License

MIT License — free to use, modify, and distribute.
