"""
backup_manager.py
Management utility for DMart Billing Backup Excel workbook.
Supports:
- Adding new invoices directly into DMart_Billing_Backup.xlsx
- Recalculating/verifying analytics
- Viewing recent records and analytics summary
- Exporting to and importing from JSON (for syncing with the Web App)
"""

import os
import sys
import json
import argparse
from datetime import datetime

# Configure UTF-8 for console output on Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

EXCEL_FILE = "DMart_Billing_Backup.xlsx"
ANKHI_RATE = 95.0
POR_RATE = 1.75

def load_workbook_safe(filename=EXCEL_FILE):
    if not os.path.exists(filename):
        from create_excel_backup import create_billing_backup_excel
        create_billing_backup_excel(filename)
    return openpyxl.load_workbook(filename)

def add_invoice(inv_no, inv_date_str, store_name, billing_period, ankhi_qty, por_qty, filename=EXCEL_FILE):
    """
    Appends an invoice to Invoice_Backup sheet in DMart_Billing_Backup.xlsx
    """
    wb = load_workbook_safe(filename)
    ws = wb["Invoice_Backup"]

    # Parse Month-Year from date (expects DD-MM-YYYY or YYYY-MM-DD)
    try:
        if "-" in inv_date_str:
            parts = inv_date_str.split("-")
            if len(parts[0]) == 4: # YYYY-MM-DD
                dt = datetime.strptime(inv_date_str, "%Y-%m-%d")
                inv_date_display = dt.strftime("%d-%m-%Y")
            else: # DD-MM-YYYY
                dt = datetime.strptime(inv_date_str, "%d-%m-%Y")
                inv_date_display = inv_date_str
            month_year = dt.strftime("%b-%Y")
        else:
            inv_date_display = inv_date_str
            month_year = "Oct-2026"
    except Exception:
        inv_date_display = inv_date_str
        month_year = "Oct-2026"

    # Find the next empty row before the TOTAL row (or find first row where column A is None)
    # Search between row 5 and row 500
    target_row = None
    total_row = None

    for r in range(5, 500):
        val = ws.cell(row=r, column=1).value
        if str(val).strip().upper() == "TOTAL":
            total_row = r
            if target_row is None:
                # Need to insert a row before TOTAL
                ws.insert_rows(r)
                target_row = r
                total_row = r + 1
            break
        elif val is None or str(val).strip() == "":
            target_row = r
            break

    if target_row is None:
        target_row = 5

    # If target_row is before total_row, we use it directly
    font_data = Font(name="Segoe UI", size=10)
    font_data_bold = Font(name="Segoe UI", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    ws.row_dimensions[target_row].height = 20

    # Inv #
    c1 = ws.cell(row=target_row, column=1, value=int(inv_no))
    c1.alignment = align_center
    c1.font = font_data_bold

    # Date
    c2 = ws.cell(row=target_row, column=2, value=inv_date_display)
    c2.alignment = align_center
    c2.font = font_data

    # Month-Year
    c3 = ws.cell(row=target_row, column=3, value=month_year)
    c3.alignment = align_center
    c3.font = font_data

    # Store
    c4 = ws.cell(row=target_row, column=4, value=store_name)
    c4.alignment = align_left
    c4.font = font_data

    # Period
    c5 = ws.cell(row=target_row, column=5, value=billing_period)
    c5.alignment = align_center
    c5.font = font_data

    # Ankhi Qty
    c6 = ws.cell(row=target_row, column=6, value=float(ankhi_qty))
    c6.alignment = align_right
    c6.number_format = "0.000"
    c6.font = font_data

    # Ankhi Rate
    c7 = ws.cell(row=target_row, column=7, value=ANKHI_RATE)
    c7.alignment = align_right
    c7.number_format = "₹#,##0.00"
    c7.font = font_data

    # Ankhi Total formula: =ROUND(F{r}*G{r}, 2)
    c8 = ws.cell(row=target_row, column=8, value=f"=ROUND(F{target_row}*G{target_row}, 2)")
    c8.alignment = align_right
    c8.number_format = "₹#,##0.00"
    c8.font = font_data

    # Por Qty
    c9 = ws.cell(row=target_row, column=9, value=int(por_qty))
    c9.alignment = align_right
    c9.number_format = "#,##0"
    c9.font = font_data

    # Por Rate
    c10 = ws.cell(row=target_row, column=10, value=POR_RATE)
    c10.alignment = align_right
    c10.number_format = "₹#,##0.00"
    c10.font = font_data

    # Por Total formula: =ROUND(I{r}*J{r}, 2)
    c11 = ws.cell(row=target_row, column=11, value=f"=ROUND(I{target_row}*J{target_row}, 2)")
    c11.alignment = align_right
    c11.number_format = "₹#,##0.00"
    c11.font = font_data

    # Grand Total formula: =ROUND(H{r}+K{r}, 0)
    c12 = ws.cell(row=target_row, column=12, value=f"=ROUND(H{target_row}+K{target_row}, 0)")
    c12.alignment = align_right
    c12.number_format = "₹#,##0"
    c12.font = font_data_bold

    # Recorded At
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c13 = ws.cell(row=target_row, column=13, value=now_str)
    c13.alignment = align_center
    c13.font = Font(name="Segoe UI", size=9, color="64748B")

    # Apply borders & alternate zebra fill
    for c_idx in range(1, 14):
        cell = ws.cell(row=target_row, column=c_idx)
        cell.border = thin_border
        if target_row % 2 == 1:
            cell.fill = fill_zebra

    # If there is a total row, update formulas if we inserted rows
    wb.save(filename)
    grand_tot_calc = round(float(ankhi_qty) * ANKHI_RATE + int(por_qty) * POR_RATE)
    print(f"✓ Recorded Invoice #{inv_no} into {filename} (Row {target_row}) | Grand Total: ₹{grand_tot_calc:,}")
    return True

def get_analytics_summary(filename=EXCEL_FILE):
    """
    Extracts high level metrics from Excel
    """
    wb = load_workbook_safe(filename)
    ws = wb["Invoice_Backup"]

    records = []
    for r in range(5, 500):
        inv_no = ws.cell(row=r, column=1).value
        if inv_no is None:
            continue
        if str(inv_no).strip().upper() == "TOTAL":
            break
        
        date_val = ws.cell(row=r, column=2).value
        month_year = ws.cell(row=r, column=3).value
        store = ws.cell(row=r, column=4).value
        period = ws.cell(row=r, column=5).value
        ankhi_q = float(ws.cell(row=r, column=6).value or 0)
        por_q = int(ws.cell(row=r, column=9).value or 0)
        
        ankhi_tot = round(ankhi_q * ANKHI_RATE, 2)
        por_tot = round(por_q * POR_RATE, 2)
        grand_tot = round(ankhi_tot + por_tot)

        records.append({
            "invNo": inv_no,
            "date": str(date_val),
            "monthYear": str(month_year),
            "store": str(store),
            "period": str(period),
            "ankhiQty": ankhi_q,
            "ankhiTotal": ankhi_tot,
            "porQty": por_q,
            "porTotal": por_tot,
            "grandTotal": grand_tot
        })

    total_rev = sum(r["grandTotal"] for r in records)
    total_ankhi_q = sum(r["ankhiQty"] for r in records)
    total_ankhi_rev = sum(r["ankhiTotal"] for r in records)
    total_por_q = sum(r["porQty"] for r in records)
    total_por_rev = sum(r["porTotal"] for r in records)

    return {
        "invoiceCount": len(records),
        "totalRevenue": total_rev,
        "totalAnkhiTons": total_ankhi_q,
        "totalAnkhiRevenue": total_ankhi_rev,
        "totalPorBoxes": total_por_q,
        "totalPorRevenue": total_por_rev,
        "records": records
    }

def print_analytics_cli():
    summary = get_analytics_summary()
    print("=" * 65)
    print("      PRADIP PATIL • D-MART BILLING ANALYTICS SUMMARY")
    print("=" * 65)
    print(f" Total Invoices Generated : {summary['invoiceCount']}")
    print(f" Total Revenue Billed     : ₹{summary['totalRevenue']:,.2f}")
    print(f" Total Ankhi DC (Tons)    : {summary['totalAnkhiTons']:,.3f} Tons  (₹{summary['totalAnkhiRevenue']:,.2f})")
    print(f" Total Por DC (Boxes)     : {summary['totalPorBoxes']:,} Boxes  (₹{summary['totalPorRevenue']:,.2f})")
    print("-" * 65)
    print(" Recent Logged Invoices:")
    for r in summary["records"][-5:]:
        print(f"  #{r['invNo']} | {r['date']} | {r['store'][:28]:<28} | ₹{r['grandTotal']:,}")
    print("=" * 65)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DMart Billing Excel Backup Manager")
    parser.add_argument("action", choices=["summary", "add", "export-json"], nargs="?", default="summary")
    parser.add_argument("--inv", type=int, help="Invoice Number")
    parser.add_argument("--date", type=str, default=datetime.now().strftime("%d-%m-%Y"), help="Date (DD-MM-YYYY)")
    parser.add_argument("--store", type=str, default="Avenue Super marts Ltd Makarpura - vadodara", help="Store Name")
    parser.add_argument("--period", type=str, default="1/10/2026 to 15/10/2026", help="Billing Period")
    parser.add_argument("--ankhi", type=float, default=0.0, help="Ankhi DC (Tons)")
    parser.add_argument("--por", type=int, default=0, help="Por DC (Boxes)")

    args = parser.parse_args()

    if args.action == "summary":
        print_analytics_cli()
    elif args.action == "add":
        if not args.inv:
            print("Error: --inv is required to add an invoice.")
            sys.exit(1)
        add_invoice(args.inv, args.date, args.store, args.period, args.ankhi, args.por)
    elif args.action == "export-json":
        summary = get_analytics_summary()
        with open("backup_records.json", "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2)
        print("✓ Exported backup_records.json successfully!")
