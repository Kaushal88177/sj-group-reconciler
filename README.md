# SJ Group - GST Reconciler Pro

### Automated GSTR-2A vs Tally Reconciliation Tool - Python Desktop Application

> Built by **Kaushal Verma** | Real-world accounting automation | 80% time saved

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows%20EXE-green?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

## 🎯 Problem Solved

Manual GST reconciliation takes **3-4 hours** every month depends on data counts. Accountants match Tally and GSTR-2A invoices manually in Excel.

**This tool does it in 30 seconds.**

## ✨ Key Features (v3.0)

### 1. Optional Column Mapping - Your Choice!
Tick only columns you want to match:
- **2 Columns:** Only Invoice No. + Gross Total
- **5 Columns:** Invoice + Taxable + IGST/CGST/SGST
- **6 Columns:** Full matching

No more "Required columns not found" error!

### 2. Smart Column Mapping UI
- Auto-detects header row (Row 3)
- Auto-maps columns by name (Voucher No., IGST etc.)
- Manual dropdown correction if auto-map fails
- Header row adjustable (1-10)

### 3. Numeric Fix - Formulas Work!
Earlier amounts were saved as TEXT, SUM() failed. Now:
- Amounts saved as **NUMERIC** with `#,##0.00` format
- Excel formulas work perfectly

### 4. Standalone Windows EXE
- Built with PyInstaller
- No Python needed on user PC
- Just double-click and use


## 💻 How to Run

### Option 1: Run Python Code Directly

```bash
git clone https://github.com/Kaushal88177/sj-group-reconciler.git
cd sj-group-reconciler
pip install -r requirements.txt
python src/Main.py
```

### Option 2: Build EXE (Windows)

```bash
# Double-click build_exe.bat
# OR manually:
pyinstaller --onefile --windowed --noupx --name "SJGroup_GST_Reconciler" src/Main.py
# EXE will be in dist/ folder
```

### Option 3: Download Ready EXE

Go to **Releases** → Download `SJGroup_GST_Reconciler.exe` → Double-click to run.

## 📂 Project Structure

```
sj-group-reconciler/
├── src/
│   └── Main.py              # Main App - 700+ lines, Optional Mapping Logic
├── sample_data/
│   ├── GSTR2A_Sample.xlsx   # Sample GSTR file
│   └── Tally_Sample.xlsx    # Sample Tally file
├── .github/workflows/
│   └── build.yml            # Auto-builds EXE on push
├── requirements.txt         # openpyxl, pandas
├── build_exe.bat            # One-click EXE builder (Ordinal 380 fix)
└── README.md
```

## 🔧 How It Works

1. User selects GSTR-2A & Tally Excel files
2. App auto-detects header row & columns
3. User ticks which columns to use for matching (e.g., only 2 columns)
4. Matching logic:
   ```
   Key = VoucherNo | GrossTotal | Taxable | IGST | CGST | SGST (only ticked fields)
   If Key exists in both files → "Match" else "Not Matching"
   ```
5. Fixes numeric formatting & saves to Desktop/SJGroup_Output/

## 📝 Resume Description

```
Project: SJ Group - Automated GST Reconciliation System (Python)
- Developed desktop application using Python (Tkinter, OpenPyXL, Pandas) automating GSTR-2A vs Tally reconciliation
- Implemented Optional Column Mapping: users can match on 2-6 columns (Invoice No., Gross, Taxable, IGST, CGST, SGST) via checkbox UI
- Fixed numeric formatting bug (TEXT→NUMBER) enabling Excel SUM formulas, reduced manual time from 4 hrs to 30 sec (80% saved)
- Built smart header detection & manual column mapping to handle different Tally exports, eliminating "Required columns" errors
- Packaged as standalone Windows EXE using PyInstaller (Ordinal 380 fix with --noupx) + Auto EXE build via GitHub Actions
- Real-world usage: STARAX Mineral Pvt Ltd, 4000+ invoices/month

Tech: Python, Tkinter, OpenPyXL, Pandas, PyInstaller, GitHub Actions
GitHub: github.com/YOUR_USERNAME/sj-group-reconciler
```

## 👨‍💻 Author

**Kaushal Verma** - Bhilai, Chhattisgarh
- Excel Automation & GST Tools
- This tool saves 3-4 hours every month (Depends on data Counts/Size) in real accounting work

## 📄 License

MIT License - Feel free to use for your business

---
⭐ **Star this repo** if it saved your time!
