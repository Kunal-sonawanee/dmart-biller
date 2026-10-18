# D-Mart Bill Generator & Analytics Hub (DMart-Biller)

Automated billing invoice generator, Excel data backup, and analytics system tailored for contractor billing (Pradip Patil • Code: 417744) servicing Avenue Supermarts Ltd (D-Mart) branches in Vadodara.

---

## 🚀 Key Features

- **Automated Tax Invoice Generation**:
  - Live calculations for **Ankhi DC Unloading (@ ₹95/Ton)** and **Por DC Unloading (@ ₹1.75/Box)**.
  - Automatic invoice number incrementing.
  - Indian numbering currency in words auto-conversion (e.g. *Seven Thousand Seven Hundred Thirty-Eight only*).
  - One-click print-ready PDF invoice generation (`html2pdf.js`).
- **Data Backup & Persistence**:
  - Every bill generated is automatically logged to local storage backup history.
  - Option to manually save bills to backup without generating a PDF.
- **Excel Spreadsheet Export with Dynamic Formulas**:
  - Exports a dual-sheet Microsoft Excel workbook (`DMart_Billing_Backup.xlsx`) via SheetJS (`xlsx`).
  - **Sheet 1 (`Invoice_Backup`)**: Detailed transactional log with `=ROUND(...)` and `=SUM(...)` formulas.
  - **Sheet 2 (`Analytics_Dashboard`)**: Executive KPI cards, store-wise performance breakdown with `=COUNTIF(...)` & `=SUMIFS(...)`, and monthly trend calculations.
- **In-App Analytics Hub**:
  - Real-time KPI summaries: Total Revenue, Total Invoices, Tons & Boxes unloaded, and earnings.
  - Visual store distribution with color-coded progress bars for Bhayali, Makarpura, Manjalpur, Waghodia, etc.
  - Searchable invoice history table with quick reload into form or deletion.
  - JSON backup export and import for seamless sync across devices.
- **Python CLI Management (`backup_manager.py`)**:
  - Summary report in terminal (`python backup_manager.py summary`).
  - Add invoices directly to Excel via command line (`python backup_manager.py add ...`).

---

## 📁 Repository Structure

```
billing-dmart/
├── index.html                 # Main billing app & Analytics Hub
├── manifest.json              # PWA manifest
├── DMart_Billing_Backup.xlsx  # Master formatted Excel backup workbook with formulas
├── create_excel_backup.py     # Python script generating the Excel template via openpyxl
├── backup_manager.py          # Python CLI utility for analytics, adding bills, & export
├── backup_records.json        # JSON data export of baseline invoices
├── .gitignore                 # Standard project ignore rules
└── README.md                  # Project documentation
```

---

## 🛠️ Usage

### 1. Web Application
Simply open `index.html` in any modern web browser or serve it locally:
```bash
python -m http.server 3000
```
Open [http://localhost:3000](http://localhost:3000) to create bills and view analytics.

### 2. Python CLI
View live billing analytics:
```bash
python backup_manager.py summary
```

Add an invoice programmatically:
```bash
python backup_manager.py add --inv 49 --store "Avenue Super marts Ltd Makarpura - vadodara" --ankhi 15.200 --por 3200
```

Re-generate or rebuild the formatted Excel backup sheet:
```bash
python create_excel_backup.py
```

---

## 📄 License
Private & Proprietary. All rights reserved.
