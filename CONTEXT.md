# SPT Extractor — Context & Pending Work

## Project Structure
- `ppn_extractor.py` — SPT Masa PPN (VAT), reads `PPN/` folder, outputs `hasil_ekstraksi_ppn.xlsx`
- `pph23_extractor.py` — SPT Masa PPh Unifikasi, reads `PPh 23/` folder, outputs `hasil_ekstraksi_pph23.xlsx`
- Git branches: `master` (baseline) → `improve-extractors` (all fixes)

---

## PPN Extractor — Pending Redesign

### Rule: DPP Always = Column 1
PDF tables have 4 columns:
1. HARGA JUAL/PENGGANTIAN/EKSPOR/DPP ← **this is "DPP"**
2. DPP NILAI LAIN ← **label as "DPP Nilai Lain", NEVER "DPP"**
3. PPN
4. PPnBM

User rule: *"always pull dpp in harga jual, do not ever pull dpp nilai lain to dpp"*

### Section I — All items must be in COLUMNS output
| Item | Description | Columns needed |
|------|-------------|----------------|
| I.A.1 | Ekspor | DPP (col1), PPN, PPnBM |
| I.A.2 | Harga Jual (kode 04/05) | Harga Jual (col1), DPP Nilai Lain (col2), PPN, PPnBM |
| I.A.3 | Turis Pasal 16E | DPP, PPN, PPnBM |
| I.A.4 | Lainnya | DPP, PPN, PPnBM |
| I.A.5 | Digunggung | Harga Jual, PPN, PPnBM |
| I.A.6 | Pemungut PPN | DPP, PPN, PPnBM |
| I.A.7 | Tidak Dipungut ("mendapat fasilitas") | DPP, PPN, PPnBM |
| I.A.8 | Dibebaskan ("mendapat fasilitas") | DPP, PPN, PPnBM |
| I.A.9 | Digunggung fasilitas ("mendapat fasilitas") | DPP, PPN, PPnBM |
| Jumlah I.A | Total | Harga Jual, DPP, PPN, PPnBM |
| I.B | Tidak terutang | single number |
| I.C | Jumlah seluruh penyerahan | single number |

### Section II — All items must be in COLUMNS output
| Item | Description | Columns needed |
|------|-------------|----------------|
| II.A | Impor/Pemanfaatan LN | DPP, PPN, PPnBM |
| II.B | DPP Nilai Lain (kode 04/05) | Harga Jual (col1), DPP Nilai Lain (col2), PPN, PPnBM |
| II.C | Selain DPP Nilai Lain | DPP, PPN, PPnBM |
| II.D | Pemungut PPN | DPP, PPN, PPnBM |
| II.E | Kompensasi kelebihan PM | single number |
| II.F | Hasil penghitungan kembali PM | single number |
| II.G | Jumlah PM diperhitungkan | DPP, PPN |
| II.H | Tidak dikreditkan/fasilitas | DPP, PPN, PPnBM |
| II.I | Digunggung | DPP |
| II.J | Jumlah perolehan | total |

### Section III — DO NOT CHANGE (user satisfied)

### Known Bugs Fixed (already in improve-extractors branch)
1. Items 7/8/9 missed — regex now matches "mendapat fasilitas" not just "Penyerahan yang PPN"
2. `ia_jumlah_dpp` wrong — pdfplumber drops col2 on Jumlah row; derived from items when needed
3. `iif_penghitungan_kembali` = 2.0 always — guard added to exclude numbered list items
4. `ia1_ekspor_dpp` extracted but never in COLUMNS output — needs to be added
5. `ia2_dpp_nilai_lain` mislabeled as "I.A.2 DPP" in COLUMNS — must be "DPP Nilai Lain"

### Technical Note
`pdfplumber.extract_tables()` only returns header rows for Sections I/II — data must stay line-based regex.
