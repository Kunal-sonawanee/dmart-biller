# D-Mart Bill Generator (DMart-Biller)

Automated billing invoice generator, Excel data backup, and analytics system tailored for contractor billing (Pradip Patil • Code: 417744) servicing Avenue Supermarts Ltd (D-Mart) branches in Vadodara.

---

## 🚀 Key Features

- **Single-screen mobile bill form** built for non-technical use: D-Mart store (tap to choose), billing half (1–15 / 16–end, current month selected automatically, with month arrows), Ankhi tons and Por boxes.
- **Automatic Bill No. and date**: bill numbers continue from **50** and go up by one after each download; the date is always today.
- **All fields required**: nothing is marked as required, but tapping Download with a field empty highlights it in red.
- **Only two actions**: **Preview** (fit-to-screen view of the exact bill) and **Download** (PDF).
- **Stores under contract**: Bhayali, Makarpura, Manjalpur, Sayajipura, Ankhol (Vadodara).
- Live totals at **₹95/Ton** (Ankhi DC) and **₹1.75/Box** (Por DC), with Indian-numbering amount in words.
- After each download the form clears, ready for the next bill.
- Owner shortcut: **press and hold "Bill No."** (about 1 second) to correct the bill number.
- **Installable app with offline support**: add it to the phone's home screen; it keeps working without internet once opened.
- **Background record keeping** (not shown to the user): every downloaded bill is saved silently in the browser's local storage.
  - Owner shortcut: press **Ctrl + Shift + E** on the bill screen to download `DMart_Billing_Backup.xlsx` (invoice log + analytics sheet with formulas).
- **Python CLI (`backup_manager.py`)** for summaries and adding bills to the Excel workbook from the terminal.

---

## 📁 Repository Structure

```
billing-dmart/
├── index.html                 # Bill maker app (Preview + Download)
├── styles.css                 # Compiled Tailwind CSS (generated — run npm run build:css)
├── src/input.css              # Tailwind source
├── tailwind.config.js         # Tailwind config (scans index.html)
├── sw.js                      # Service worker for offline use
├── manifest.json              # PWA manifest
├── icons/                     # App icons
├── vendor/                    # Pinned html2pdf.js 0.10.1 and SheetJS 0.18.5
├── DMart_Billing_Backup.xlsx  # Master formatted Excel backup workbook with formulas
├── create_excel_backup.py     # Python script generating the Excel template via openpyxl
├── backup_manager.py          # Python CLI utility for analytics, adding bills, & export
├── backup_records.json        # JSON data export of baseline invoices
└── README.md                  # Project documentation
```

---

## 🛠️ Usage

### 1. Web Application
Host the folder on any static HTTPS host (GitHub Pages, Netlify, etc.) — HTTPS is needed for offline mode and home-screen install. To run locally:
```bash
python -m http.server 3000
```
Open [http://localhost:3000](http://localhost:3000) to create bills.

After changing classes in `index.html`, rebuild the CSS and bump `CACHE_VERSION` in `sw.js`:
```bash
npm install
npm run build:css
```

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
