# SPT Extractor

Python tools for extracting data from Indonesian tax PDF documents (SPT Masa) into Excel.

## Extractors

### `ppn_extractor.py` — SPT Masa PPN (Value Added Tax)
Extracts all fields from DJP SPT Masa PPN forms including:
- **Section I** — Penyerahan Barang dan Jasa (items 1–9, all columns)
- **Section II** — Perolehan Barang dan Jasa (items A–J, all columns)
- **Section III** — Perhitungan PPN Kurang/Lebih Bayar

### `pph23_extractor.py` — SPT Masa PPh Unifikasi (Unified Income Tax)
Extracts withholding tax data from PPh 23/26 unified forms.

---

## Usage

### No arguments (auto-detect folder)
```bash
python ppn_extractor.py
python pph23_extractor.py
```
Looks for PDFs in `PPN/` and `PPh 23/` folders respectively. Outputs Excel to the project root.

### Folder
```bash
python ppn_extractor.py <folder> [--output output.xlsx]
python pph23_extractor.py <folder> [--output output.xlsx]
```

### Single file
```bash
python ppn_extractor.py file.pdf
python pph23_extractor.py file.pdf
```

---

## Output

| Extractor | Default output file |
|-----------|-------------------|
| PPN | `hasil_ekstraksi_ppn.xlsx` |
| PPh 23 | `hasil_ekstraksi_pph23.xlsx` |

Excel files contain:
- Styled multi-row headers with section grouping (color-coded)
- One row per PDF file processed
- Alternating row colors
- Summary sheet (total files, success, errors)

---

## Requirements

```bash
pip install pdfplumber openpyxl
```

Python 3.8+

---

## Folder Structure

```
SPT EXTRACTOR/
├── ppn_extractor.py
├── pph23_extractor.py
├── PPN/              # Place SPT Masa PPN PDFs here
│   └── *.pdf
└── PPh 23/           # Place SPT Masa PPh PDFs here
    └── *.pdf
```

---

## Notes

- PDFs must be digital DJP-issued SPT Masa forms (not scanned images)
- Extraction is line-based via regex; tested against 12-month sets of real SPT PDFs
- DPP values are always sourced from Column 1 (Harga Jual/Penggantian/Ekspor/DPP), not DPP Nilai Lain
