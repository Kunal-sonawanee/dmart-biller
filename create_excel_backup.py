"""
Script to create DMart_Billing_Backup.xlsx with:
1. 'Invoice_Backup' sheet: detailed transactional records with Excel formulas
2. 'Analytics_Dashboard' sheet: executive KPI cards, store-wise breakdown, and monthly trends
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def create_billing_backup_excel(filename="DMart_Billing_Backup.xlsx"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # -------------------------------------------------------------
    # 1. STYLES DEFINITION
    # -------------------------------------------------------------
    font_family = "Segoe UI"
    
    # Palette
    c_gold_dark = "B45309"      # Amber 700
    c_gold_light = "FEF3C7"     # Amber 100
    c_slate_dark = "0F172A"     # Slate 900
    c_slate_card = "1E293B"     # Slate 800
    c_slate_sub = "475569"      # Slate 600
    c_header_fill = "1E3A8A"    # Deep Blue / Navy
    c_header_text = "FFFFFF"
    c_zebra = "F8FAFC"
    c_emerald = "059669"
    c_emerald_light = "ECFDF5"

    font_title = Font(name=font_family, size=16, bold=True, color="FFFFFF")
    font_subtitle = Font(name=font_family, size=10, bold=False, color="FEF3C7")
    font_card_num = Font(name=font_family, size=18, bold=True, color="FFFFFF")
    font_card_lbl = Font(name=font_family, size=9, bold=True, color="94A3B8")
    
    font_tbl_header = Font(name=font_family, size=10, bold=True, color=c_header_text)
    fill_tbl_header = PatternFill(start_color=c_header_fill, end_color=c_header_fill, fill_type="solid")
    
    font_sub_header = Font(name=font_family, size=11, bold=True, color="0F172A")
    fill_sub_header = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    font_data = Font(name=font_family, size=10)
    font_data_bold = Font(name=font_family, size=10, bold=True)
    
    fill_zebra = PatternFill(start_color=c_zebra, end_color=c_zebra, fill_type="solid")
    fill_total = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    double_bottom_side = Side(border_style="double", color="0F172A")
    top_thin_side = Side(border_style="thin", color="0F172A")
    border_total = Border(top=top_thin_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)

    # -------------------------------------------------------------
    # 2. SHEET 1: INVOICE_BACKUP
    # -------------------------------------------------------------
    ws_inv = wb.create_sheet(title="Invoice_Backup")
    ws_inv.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_inv.merge_cells("A1:M1")
    ws_inv["A1"] = "PRADIP PATIL • D-MART BILLING RECORDS & DATA BACKUP"
    ws_inv["A1"].font = font_title
    ws_inv["A1"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ws_inv["A1"].alignment = align_left
    ws_inv.row_dimensions[1].height = 36

    ws_inv.merge_cells("A2:M2")
    ws_inv["A2"] = "Contractor Code: 417744 | Loading & Unloading Services | Rates: Ankhi DC = ₹95/Ton, Por DC = ₹1.75/Box"
    ws_inv["A2"].font = Font(name=font_family, size=9.5, italic=True, color="FEF3C7")
    ws_inv["A2"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws_inv["A2"].alignment = align_left
    ws_inv.row_dimensions[2].height = 20

    # Headers in Row 4
    headers = [
        "Inv #",
        "Invoice Date",
        "Month-Year",
        "D-Mart Store / Branch",
        "Billing Period",
        "Ankhi DC (Tons)",
        "Rate (₹/Ton)",
        "Ankhi Total (₹)",
        "Por DC (Boxes)",
        "Rate (₹/Box)",
        "Por Total (₹)",
        "Grand Total (₹)",
        "Recorded At"
    ]
    
    ws_inv.row_dimensions[4].height = 26
    for col_idx, h in enumerate(headers, 1):
        cell = ws_inv.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center if col_idx in [1, 2, 3, 5, 7, 10, 13] else (align_right if col_idx in [6, 8, 9, 11, 12] else align_left)
        cell.border = border_cell

    # Baseline Sample Data (Invoice #47 and #48 as initial verified records)
    sample_records = [
        (47, "15-10-2026", "Oct-2026", "Avenue Super marts Ltd Makarpura - vadodara", "1/10/2026 to 15/10/2026", 18.450, 95, 3420, 1.75, "2026-10-15 17:30:00"),
        (48, "18-10-2026", "Oct-2026", "Avenue Super marts Ltd Bhayali - vadodara", "1/10/2026 to 15/10/2026", 24.120, 95, 4150, 1.75, "2026-10-18 11:20:00")
    ]

    start_row = 5
    for i, r in enumerate(sample_records):
        curr_row = start_row + i
        ws_inv.row_dimensions[curr_row].height = 20
        inv_no, inv_date, my, store, period, ankhi_q, ankhi_r, por_q, por_r, recorded = r

        ws_inv.cell(row=curr_row, column=1, value=inv_no).alignment = align_center
        ws_inv.cell(row=curr_row, column=2, value=inv_date).alignment = align_center
        ws_inv.cell(row=curr_row, column=3, value=my).alignment = align_center
        ws_inv.cell(row=curr_row, column=4, value=store).alignment = align_left
        ws_inv.cell(row=curr_row, column=5, value=period).alignment = align_center

        # Ankhi Qty & Rate
        c_aq = ws_inv.cell(row=curr_row, column=6, value=ankhi_q)
        c_aq.alignment = align_right
        c_aq.number_format = "0.000"

        c_ar = ws_inv.cell(row=curr_row, column=7, value=ankhi_r)
        c_ar.alignment = align_right
        c_ar.number_format = "₹#,##0.00"

        # Ankhi Total Formula: =F{row}*G{row}
        c_at = ws_inv.cell(row=curr_row, column=8, value=f"=ROUND(F{curr_row}*G{curr_row}, 2)")
        c_at.alignment = align_right
        c_at.number_format = "₹#,##0.00"

        # Por Qty & Rate
        c_pq = ws_inv.cell(row=curr_row, column=9, value=por_q)
        c_pq.alignment = align_right
        c_pq.number_format = "#,##0"

        c_pr = ws_inv.cell(row=curr_row, column=10, value=por_r)
        c_pr.alignment = align_right
        c_pr.number_format = "₹#,##0.00"

        # Por Total Formula: =I{row}*J{row}
        c_pt = ws_inv.cell(row=curr_row, column=11, value=f"=ROUND(I{curr_row}*J{curr_row}, 2)")
        c_pt.alignment = align_right
        c_pt.number_format = "₹#,##0.00"

        # Grand Total Formula: =ROUND(H{row}+K{row}, 0)
        c_gt = ws_inv.cell(row=curr_row, column=12, value=f"=ROUND(H{curr_row}+K{curr_row}, 0)")
        c_gt.alignment = align_right
        c_gt.font = font_data_bold
        c_gt.number_format = "₹#,##0"

        c_rec = ws_inv.cell(row=curr_row, column=13, value=recorded)
        c_rec.alignment = align_center
        c_rec.font = Font(name=font_family, size=9, color="64748B")

        for c_idx in range(1, 14):
            cell = ws_inv.cell(row=curr_row, column=c_idx)
            cell.font = font_data_bold if c_idx in [1, 12] else font_data
            cell.border = border_cell
            if i % 2 == 1:
                cell.fill = fill_zebra

    # Blank rows for template readiness (rows 7 to 15)
    for blank_row in range(7, 16):
        ws_inv.row_dimensions[blank_row].height = 19
        for c_idx in range(1, 14):
            cell = ws_inv.cell(row=blank_row, column=c_idx)
            cell.border = border_cell
            if blank_row % 2 == 1:
                cell.fill = fill_zebra

    # Formula row for TOTALS (Row 16)
    row_tot = 16
    ws_inv.row_dimensions[row_tot].height = 24
    ws_inv.cell(row=row_tot, column=1, value="TOTAL").alignment = align_center
    ws_inv.cell(row=row_tot, column=1).font = font_data_bold
    ws_inv.cell(row=row_tot, column=1).fill = fill_total
    ws_inv.cell(row=row_tot, column=1).border = border_total

    for c in range(2, 6):
        ws_inv.cell(row=row_tot, column=c).fill = fill_total
        ws_inv.cell(row=row_tot, column=c).border = border_total

    # Sum of Ankhi Tons
    c_tot_aq = ws_inv.cell(row=row_tot, column=6, value=f"=SUM(F5:F{row_tot-1})")
    c_tot_aq.alignment = align_right
    c_tot_aq.font = font_data_bold
    c_tot_aq.number_format = "0.000"
    c_tot_aq.fill = fill_total
    c_tot_aq.border = border_total

    ws_inv.cell(row=row_tot, column=7).fill = fill_total
    ws_inv.cell(row=row_tot, column=7).border = border_total

    # Sum of Ankhi Total
    c_tot_at = ws_inv.cell(row=row_tot, column=8, value=f"=SUM(H5:H{row_tot-1})")
    c_tot_at.alignment = align_right
    c_tot_at.font = font_data_bold
    c_tot_at.number_format = "₹#,##0.00"
    c_tot_at.fill = fill_total
    c_tot_at.border = border_total

    # Sum of Por Boxes
    c_tot_pq = ws_inv.cell(row=row_tot, column=9, value=f"=SUM(I5:I{row_tot-1})")
    c_tot_pq.alignment = align_right
    c_tot_pq.font = font_data_bold
    c_tot_pq.number_format = "#,##0"
    c_tot_pq.fill = fill_total
    c_tot_pq.border = border_total

    ws_inv.cell(row=row_tot, column=10).fill = fill_total
    ws_inv.cell(row=row_tot, column=10).border = border_total

    # Sum of Por Total
    c_tot_pt = ws_inv.cell(row=row_tot, column=11, value=f"=SUM(K5:K{row_tot-1})")
    c_tot_pt.alignment = align_right
    c_tot_pt.font = font_data_bold
    c_tot_pt.number_format = "₹#,##0.00"
    c_tot_pt.fill = fill_total
    c_tot_pt.border = border_total

    # Grand Total Overall
    c_tot_gt = ws_inv.cell(row=row_tot, column=12, value=f"=SUM(L5:L{row_tot-1})")
    c_tot_gt.alignment = align_right
    c_tot_gt.font = Font(name=font_family, size=11, bold=True, color="9A3412")
    c_tot_gt.number_format = "₹#,##0.00"
    c_tot_gt.fill = fill_total
    c_tot_gt.border = border_total

    ws_inv.cell(row=row_tot, column=13).fill = fill_total
    ws_inv.cell(row=row_tot, column=13).border = border_total

    # Column widths
    col_widths = {
        "A": 9,   # Inv #
        "B": 14,  # Date
        "C": 13,  # Month-Year
        "D": 38,  # Store / Branch
        "E": 25,  # Billing Period
        "F": 17,  # Ankhi DC (Tons)
        "G": 14,  # Rate
        "H": 17,  # Ankhi Amount
        "I": 15,  # Por DC (Boxes)
        "J": 14,  # Rate
        "K": 16,  # Por Amount
        "L": 18,  # Grand Total
        "M": 21   # Recorded At
    }
    for col_l, width in col_widths.items():
        ws_inv.column_dimensions[col_l].width = width

    # -------------------------------------------------------------
    # 3. SHEET 2: ANALYTICS_DASHBOARD
    # -------------------------------------------------------------
    ws_dash = wb.create_sheet(title="Analytics_Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("A1:G1")
    ws_dash["A1"] = "D-MART BILLING ANALYTICS & EXECUTIVE DASHBOARD"
    ws_dash["A1"].font = font_title
    ws_dash["A1"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ws_dash["A1"].alignment = align_left
    ws_dash.row_dimensions[1].height = 36

    ws_dash.merge_cells("A2:G2")
    ws_dash["A2"] = "Live Key Performance Indicators, Store Distribution, and Volume Analytics (Formulas Linked to Invoice_Backup)"
    ws_dash["A2"].font = Font(name=font_family, size=9.5, italic=True, color="CBD5E1")
    ws_dash["A2"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws_dash["A2"].alignment = align_left
    ws_dash.row_dimensions[2].height = 20

    # 4 Executive KPI Cards in Rows 4-6
    # Card 1: Total Revenue (Cols A-B)
    # Card 2: Total Invoices (Col C)
    # Card 3: Ankhi Tons & Revenue (Cols D-E)
    # Card 4: Por Boxes & Revenue (Cols F-G)

    fill_card_dark = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_card_amber = PatternFill(start_color="78350F", end_color="78350F", fill_type="solid")
    fill_card_blue = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_card_emerald = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")

    # Card 1: TOTAL REVENUE
    ws_dash.merge_cells("A4:B4")
    ws_dash["A4"] = "TOTAL BILLING REVENUE"
    ws_dash["A4"].font = font_card_lbl
    ws_dash["A4"].fill = fill_card_amber
    ws_dash["A4"].alignment = align_center

    ws_dash.merge_cells("A5:B6")
    ws_dash["A5"] = "=SUM(Invoice_Backup!L5:L1000)"
    ws_dash["A5"].font = Font(name=font_family, size=20, bold=True, color="FDE68A")
    ws_dash["A5"].fill = fill_card_amber
    ws_dash["A5"].alignment = align_center
    ws_dash["A5"].number_format = "₹#,##0.00"

    # Card 2: TOTAL INVOICES
    ws_dash["C4"] = "TOTAL INVOICES"
    ws_dash["C4"].font = font_card_lbl
    ws_dash["C4"].fill = fill_card_dark
    ws_dash["C4"].alignment = align_center

    ws_dash.merge_cells("C5:C6")
    ws_dash["C5"] = "=COUNT(Invoice_Backup!A5:A1000)"
    ws_dash["C5"].font = font_card_num
    ws_dash["C5"].fill = fill_card_dark
    ws_dash["C5"].alignment = align_center
    ws_dash["C5"].number_format = "#,##0"

    # Card 3: ANKHI DC (TONS & REVENUE)
    ws_dash.merge_cells("D4:E4")
    ws_dash["D4"] = "ANKHI DC (TONS / REVENUE @ ₹95)"
    ws_dash["D4"].font = font_card_lbl
    ws_dash["D4"].fill = fill_card_blue
    ws_dash["D4"].alignment = align_center

    ws_dash["D5"] = "=SUM(Invoice_Backup!F5:F1000)"
    ws_dash["D5"].font = Font(name=font_family, size=15, bold=True, color="93C5FD")
    ws_dash["D5"].fill = fill_card_blue
    ws_dash["D5"].alignment = align_center
    ws_dash["D5"].number_format = '0.000 "Tons"'

    ws_dash["E5"] = "=SUM(Invoice_Backup!H5:H1000)"
    ws_dash["E5"].font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    ws_dash["E5"].fill = fill_card_blue
    ws_dash["E5"].alignment = align_center
    ws_dash["E5"].number_format = "₹#,##0.00"

    ws_dash.merge_cells("D6:E6")
    ws_dash["D6"] = "Unloading Rate: ₹95.00 / Metric Ton"
    ws_dash["D6"].font = Font(name=font_family, size=8.5, italic=True, color="BFDBFE")
    ws_dash["D6"].fill = fill_card_blue
    ws_dash["D6"].alignment = align_center

    # Card 4: POR DC (BOXES & REVENUE)
    ws_dash.merge_cells("F4:G4")
    ws_dash["F4"] = "POR DC (BOXES / REVENUE @ ₹1.75)"
    ws_dash["F4"].font = font_card_lbl
    ws_dash["F4"].fill = fill_card_emerald
    ws_dash["F4"].alignment = align_center

    ws_dash["F5"] = "=SUM(Invoice_Backup!I5:I1000)"
    ws_dash["F5"].font = Font(name=font_family, size=15, bold=True, color="A7F3D0")
    ws_dash["F5"].fill = fill_card_emerald
    ws_dash["F5"].alignment = align_center
    ws_dash["F5"].number_format = '#,##0 "Boxes"'

    ws_dash["G5"] = "=SUM(Invoice_Backup!K5:K1000)"
    ws_dash["G5"].font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    ws_dash["G5"].fill = fill_card_emerald
    ws_dash["G5"].alignment = align_center
    ws_dash["G5"].number_format = "₹#,##0.00"

    ws_dash.merge_cells("F6:G6")
    ws_dash["F6"] = "Unloading Rate: ₹1.75 / Box"
    ws_dash["F6"].font = Font(name=font_family, size=8.5, italic=True, color="D1FAE5")
    ws_dash["F6"].fill = fill_card_emerald
    ws_dash["F6"].alignment = align_center

    for r_idx in range(4, 7):
        for col_l in ["A", "B", "C", "D", "E", "F", "G"]:
            ws_dash[f"{col_l}{r_idx}"].border = border_cell

    # -------------------------------------------------------------
    # STORE-WISE PERFORMANCE BREAKDOWN TABLE
    # -------------------------------------------------------------
    row_st_hdr = 8
    ws_dash.merge_cells(f"A{row_st_hdr}:G{row_st_hdr}")
    ws_dash[f"A{row_st_hdr}"] = "STORE-WISE PERFORMANCE & REVENUE BREAKDOWN"
    ws_dash[f"A{row_st_hdr}"].font = font_sub_header
    ws_dash[f"A{row_st_hdr}"].fill = fill_sub_header
    ws_dash[f"A{row_st_hdr}"].alignment = align_left
    ws_dash.row_dimensions[row_st_hdr].height = 24

    st_headers = [
        "Store / Branch Name",
        "Keyword Match",
        "Invoice Count",
        "Total Ankhi (Tons)",
        "Total Por (Boxes)",
        "Total Amount (₹)",
        "% Share of Revenue"
    ]
    row_th = 9
    ws_dash.row_dimensions[row_th].height = 24
    for col_idx, h in enumerate(st_headers, 1):
        c = ws_dash.cell(row=row_th, column=col_idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center if col_idx in [2, 3] else (align_right if col_idx >= 4 else align_left)
        c.border = border_cell

    stores = [
        ("Avenue Super marts Ltd Bhayali - vadodara", "Bhayali"),
        ("Avenue Super marts Ltd Makarpura - vadodara", "Makarpura"),
        ("Avenue Super marts Ltd Manjalpur - vadodara", "Manjalpur"),
        ("Avenue Super marts Ltd Sayajipura - vadodara", "Sayajipura"),
        ("Avenue Super marts Ltd Ankhol - vadodara", "Ankhol")
    ]

    for idx, (st_name, kw) in enumerate(stores):
        r_num = 10 + idx
        ws_dash.row_dimensions[r_num].height = 21

        # Store Name
        c1 = ws_dash.cell(row=r_num, column=1, value=st_name)
        c1.alignment = align_left
        c1.font = font_data_bold if kw else font_data

        # Keyword
        c2 = ws_dash.cell(row=r_num, column=2, value=kw)
        c2.alignment = align_center
        c2.font = font_data

        if kw != "":
            # Invoice count formula: =COUNTIF(Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")
            c3 = ws_dash.cell(row=r_num, column=3, value=f'=COUNTIF(Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")')
            # Ankhi tons: =SUMIFS(Invoice_Backup!$F$5:$F$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")
            c4 = ws_dash.cell(row=r_num, column=4, value=f'=SUMIFS(Invoice_Backup!$F$5:$F$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")')
            # Por boxes: =SUMIFS(Invoice_Backup!$I$5:$I$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")
            c5 = ws_dash.cell(row=r_num, column=5, value=f'=SUMIFS(Invoice_Backup!$I$5:$I$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")')
            # Total amount: =SUMIFS(Invoice_Backup!$L$5:$L$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")
            c6 = ws_dash.cell(row=r_num, column=6, value=f'=SUMIFS(Invoice_Backup!$L$5:$L$1000, Invoice_Backup!$D$5:$D$1000, "*"&B{r_num}&"*")')
        else:
            # Other stores (Total minus known 4 branches)
            c3 = ws_dash.cell(row=r_num, column=3, value=f'=MAX(0, COUNT(Invoice_Backup!$A$5:$A$1000)-SUM(C10:C13))')
            c4 = ws_dash.cell(row=r_num, column=4, value=f'=MAX(0, SUM(Invoice_Backup!$F$5:$F$1000)-SUM(D10:D13))')
            c5 = ws_dash.cell(row=r_num, column=5, value=f'=MAX(0, SUM(Invoice_Backup!$I$5:$I$1000)-SUM(E10:E13))')
            c6 = ws_dash.cell(row=r_num, column=6, value=f'=MAX(0, SUM(Invoice_Backup!$L$5:$L$1000)-SUM(F10:F13))')

        c3.alignment = align_center
        c3.number_format = "#,##0"

        c4.alignment = align_right
        c4.number_format = "0.000"

        c5.alignment = align_right
        c5.number_format = "#,##0"

        c6.alignment = align_right
        c6.number_format = "₹#,##0.00"
        c6.font = font_data_bold

        # % Share formula: =IFERROR(F{r_num}/$A$5, 0)
        c7 = ws_dash.cell(row=r_num, column=7, value=f'=IFERROR(F{r_num}/$A$5, 0)')
        c7.alignment = align_right
        c7.number_format = "0.0%"

        for col_idx in range(1, 8):
            cell = ws_dash.cell(row=r_num, column=col_idx)
            cell.border = border_cell
            if idx % 2 == 1:
                cell.fill = fill_zebra

    # Store Table Total Row (Row 15)
    row_st_tot = 15
    ws_dash.row_dimensions[row_st_tot].height = 24
    c_st_tot_lbl = ws_dash.cell(row=row_st_tot, column=1, value="STORE SUMMARY TOTAL")
    c_st_tot_lbl.font = font_data_bold
    c_st_tot_lbl.fill = fill_total
    c_st_tot_lbl.border = border_total

    ws_dash.cell(row=row_st_tot, column=2).fill = fill_total
    ws_dash.cell(row=row_st_tot, column=2).border = border_total

    c_st_tot_c = ws_dash.cell(row=row_st_tot, column=3, value=f"=SUM(C10:C14)")
    c_st_tot_c.alignment = align_center
    c_st_tot_c.font = font_data_bold
    c_st_tot_c.number_format = "#,##0"
    c_st_tot_c.fill = fill_total
    c_st_tot_c.border = border_total

    c_st_tot_aq = ws_dash.cell(row=row_st_tot, column=4, value=f"=SUM(D10:D14)")
    c_st_tot_aq.alignment = align_right
    c_st_tot_aq.font = font_data_bold
    c_st_tot_aq.number_format = "0.000"
    c_st_tot_aq.fill = fill_total
    c_st_tot_aq.border = border_total

    c_st_tot_pq = ws_dash.cell(row=row_st_tot, column=5, value=f"=SUM(E10:E14)")
    c_st_tot_pq.alignment = align_right
    c_st_tot_pq.font = font_data_bold
    c_st_tot_pq.number_format = "#,##0"
    c_st_tot_pq.fill = fill_total
    c_st_tot_pq.border = border_total

    c_st_tot_rev = ws_dash.cell(row=row_st_tot, column=6, value=f"=SUM(F10:F14)")
    c_st_tot_rev.alignment = align_right
    c_st_tot_rev.font = Font(name=font_family, size=10.5, bold=True, color="9A3412")
    c_st_tot_rev.number_format = "₹#,##0.00"
    c_st_tot_rev.fill = fill_total
    c_st_tot_rev.border = border_total

    c_st_tot_pct = ws_dash.cell(row=row_st_tot, column=7, value=f"=SUM(G10:G14)")
    c_st_tot_pct.alignment = align_right
    c_st_tot_pct.font = font_data_bold
    c_st_tot_pct.number_format = "0.0%"
    c_st_tot_pct.fill = fill_total
    c_st_tot_pct.border = border_total

    # -------------------------------------------------------------
    # MONTH-WISE BREAKDOWN TABLE
    # -------------------------------------------------------------
    row_mo_hdr = 17
    ws_dash.merge_cells(f"A{row_mo_hdr}:G{row_mo_hdr}")
    ws_dash[f"A{row_mo_hdr}"] = "MONTHLY BILLING TRENDS & QUANTITIES"
    ws_dash[f"A{row_mo_hdr}"].font = font_sub_header
    ws_dash[f"A{row_mo_hdr}"].fill = fill_sub_header
    ws_dash[f"A{row_mo_hdr}"].alignment = align_left
    ws_dash.row_dimensions[row_mo_hdr].height = 24

    mo_headers = [
        "Month - Year",
        "Period Filter",
        "Invoices Billed",
        "Ankhi DC (Tons)",
        "Por DC (Boxes)",
        "Monthly Amount (₹)",
        "Average Bill Value (₹)"
    ]
    row_mo_th = 18
    ws_dash.row_dimensions[row_mo_th].height = 24
    for col_idx, h in enumerate(mo_headers, 1):
        c = ws_dash.cell(row=row_mo_th, column=col_idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center if col_idx in [1, 2, 3] else align_right
        c.border = border_cell

    months = [
        ("Oct-2026", "October 2026"),
        ("Nov-2026", "November 2026"),
        ("Dec-2026", "December 2026"),
        ("Jan-2027", "January 2027")
    ]

    for m_idx, (m_code, m_full) in enumerate(months):
        r_num = 19 + m_idx
        ws_dash.row_dimensions[r_num].height = 21

        c1 = ws_dash.cell(row=r_num, column=1, value=m_code)
        c1.alignment = align_center
        c1.font = font_data_bold

        c2 = ws_dash.cell(row=r_num, column=2, value=m_full)
        c2.alignment = align_left
        c2.font = font_data

        # COUNTIFS
        c3 = ws_dash.cell(row=r_num, column=3, value=f'=COUNTIF(Invoice_Backup!$C$5:$C$1000, A{r_num})')
        c3.alignment = align_center
        c3.number_format = "#,##0"

        # Ankhi tons
        c4 = ws_dash.cell(row=r_num, column=4, value=f'=SUMIF(Invoice_Backup!$C$5:$C$1000, A{r_num}, Invoice_Backup!$F$5:$F$1000)')
        c4.alignment = align_right
        c4.number_format = "0.000"

        # Por boxes
        c5 = ws_dash.cell(row=r_num, column=5, value=f'=SUMIF(Invoice_Backup!$C$5:$C$1000, A{r_num}, Invoice_Backup!$I$5:$I$1000)')
        c5.alignment = align_right
        c5.number_format = "#,##0"

        # Total amount
        c6 = ws_dash.cell(row=r_num, column=6, value=f'=SUMIF(Invoice_Backup!$C$5:$C$1000, A{r_num}, Invoice_Backup!$L$5:$L$1000)')
        c6.alignment = align_right
        c6.number_format = "₹#,##0.00"
        c6.font = font_data_bold

        # Average per bill
        c7 = ws_dash.cell(row=r_num, column=7, value=f'=IFERROR(F{r_num}/C{r_num}, 0)')
        c7.alignment = align_right
        c7.number_format = "₹#,##0.00"

        for col_idx in range(1, 8):
            cell = ws_dash.cell(row=r_num, column=col_idx)
            cell.border = border_cell
            if m_idx % 2 == 1:
                cell.fill = fill_zebra

    # Column widths for Analytics sheet
    dash_col_widths = {
        "A": 36,  # Store Name / Month
        "B": 18,  # Keyword / Period
        "C": 16,  # Count
        "D": 22,  # Ankhi Qty
        "E": 20,  # Por Qty
        "F": 22,  # Total
        "G": 22   # % / Avg
    }
    for col_l, width in dash_col_widths.items():
        ws_dash.column_dimensions[col_l].width = width

    # Save workbook
    wb.save(filename)
    print(f"Created {filename} successfully with Invoice_Backup and Analytics_Dashboard!")

if __name__ == "__main__":
    create_billing_backup_excel()
