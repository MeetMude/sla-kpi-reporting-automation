"""
04_build_reports.py
-------------------
Purpose: Generates output/SLA_KPI_Report.xlsx with 7 distinct worksheets:
1. Cleaned_Data (Excel Table format with filters & styling)
2. Calc (Dynamic Excel formulas for resolution_hours, SLA_Met, month, breach_hours)
3. KPI_Summary (KPI card blocks using SUMIFS, COUNTIFS, AVERAGEIFS + conditional formatting)
4. Monthly_Trend (Monthly formula table + Line Chart & Column Chart)
5. Priority_Category (Priority & Category formula tables + Bar Charts)
6. Dashboard (Executive dashboard with integrated KPI cards & native Excel charts)
7. Notes (Plain-English reference guide explaining formulas & SLA targets)
"""

import os
import time
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import LineChart, BarChart, Reference

def create_styled_line_chart(ws_source, data_col, cat_col, title, x_title, y_title):
    """Helper function to create a clean single-series straight LineChart with no legend."""
    chart = LineChart()
    chart.title = title
    chart.style = 13
    chart.legend = None  # Remove legend completely
    
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    
    chart.width = 13.5
    chart.height = 9.0

    data_ref = Reference(ws_source, min_col=data_col, min_row=4, max_row=16)
    cats_ref = Reference(ws_source, min_col=cat_col, min_row=5, max_row=16)
    
    chart.add_data(data_ref, titles_from_data=True, from_rows=False)
    chart.set_categories(cats_ref)

    if chart.series:
        s = chart.series[0]
        s.smooth = False  # Straight line (not smoothed)
        s.graphicalProperties.line.solidFill = "1F4E78"  # Navy blue line color
        s.graphicalProperties.line.width = 25000  # Clean line weight

    return chart

def create_styled_col_chart(ws_source, data_col, cat_col, title, x_title, y_title, min_r=4, max_r=16, fill_color="1F4E78"):
    """Helper function to create a clean single-series Column Chart with single bar color and no legend."""
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.legend = None  # Remove legend completely
    
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    
    chart.width = 13.5
    chart.height = 9.0

    data_ref = Reference(ws_source, min_col=data_col, min_row=min_r, max_row=max_r)
    cats_ref = Reference(ws_source, min_col=cat_col, min_row=min_r+1, max_row=max_r)
    
    chart.add_data(data_ref, titles_from_data=True, from_rows=False)
    chart.set_categories(cats_ref)

    if chart.series:
        s = chart.series[0]
        s.graphicalProperties.solidFill = fill_color  # Single color for all bars

    return chart

def main():
    print("=" * 60)
    print("STEP 4: BUILDING EXCEL SLA & KPI REPORT WITH FORMULAS & CHARTS")
    print("=" * 60)

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    cleaned_csv_path = os.path.join(base_dir, 'data', 'cleaned_tickets.csv')
    report_output_path = os.path.join(base_dir, 'output', 'SLA_KPI_Report.xlsx')

    if not os.path.exists(cleaned_csv_path):
        raise FileNotFoundError(f"Cleaned CSV not found at {cleaned_csv_path}. Run 02_clean_data.py first.")

    df_clean = pd.read_csv(cleaned_csv_path)
    num_rows = len(df_clean)

    # Initialize workbook cleanly
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Color Palette & Styling Setup
    navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    accent_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    kpi_card_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    kpi_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    font_white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_kpi_num = Font(name="Calibri", size=20, bold=True, color="1F4E78")
    font_kpi_label = Font(name="Calibri", size=10, bold=True, color="595959")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    kpi_border = Border(
        left=Side(border_style="thin", color="B4C6E7"),
        right=Side(border_style="thin", color="B4C6E7"),
        top=Side(border_style="thin", color="B4C6E7"),
        bottom=Side(border_style="thin", color="B4C6E7")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # SHEET 1: Cleaned_Data
    # =========================================================================
    ws_data = wb.create_sheet(title="Cleaned_Data")
    ws_data.views.sheetView[0].showGridLines = True

    headers_data = [
        "ticket_id", "created_date", "resolved_date", "priority", 
        "category", "team", "agent", "status", 
        "customer_satisfaction", "sla_target_hours", "resolution_hours", "sla_met"
    ]
    ws_data.append(headers_data)

    for row_idx, row in df_clean.iterrows():
        ws_data.append([
            row['ticket_id'],
            str(row['created_date']) if pd.notnull(row['created_date']) else "",
            str(row['resolved_date']) if pd.notnull(row['resolved_date']) else "",
            row['priority'],
            row['category'],
            row['team'],
            row['agent'],
            row['status'],
            row['customer_satisfaction'] if pd.notnull(row['customer_satisfaction']) else "",
            row['sla_target_hours'],
            row['resolution_hours'] if pd.notnull(row['resolution_hours']) else "",
            int(row['sla_met']) if pd.notnull(row['sla_met']) else ""
        ])

    for col_num in range(1, len(headers_data) + 1):
        cell = ws_data.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = font_white_bold
        cell.alignment = align_center

    ws_data.freeze_panes = "A2"

    # =========================================================================
    # SHEET 2: Calc (Helper columns with EXCEL FORMULAS)
    # =========================================================================
    ws_calc = wb.create_sheet(title="Calc")
    ws_calc.views.sheetView[0].showGridLines = True

    headers_calc = [
        "ticket_id", "created_date", "resolved_date", "priority", 
        "category", "team", "status", "resolution_hours", 
        "SLA_Met", "sla_target_hours", "month", "breach_hours"
    ]
    ws_calc.append(headers_calc)

    for r_idx in range(2, num_rows + 2):
        res_hrs_formula = f'=Cleaned_Data!K{r_idx}'
        sla_met_formula = f'=IF(OR(G{r_idx}="Resolved", G{r_idx}="Closed"), IF(AND(H{r_idx}<>"", H{r_idx}<=J{r_idx}), "YES", "NO"), "")'
        sla_target_formula = f'=Cleaned_Data!J{r_idx}'
        month_formula = f'=LEFT(B{r_idx}, 7)'
        breach_formula = f'=IF(I{r_idx}="NO", H{r_idx}-J{r_idx}, 0)'

        row_data = df_clean.iloc[r_idx - 2]
        
        ws_calc.append([
            row_data['ticket_id'],
            str(row_data['created_date']) if pd.notnull(row_data['created_date']) else "",
            str(row_data['resolved_date']) if pd.notnull(row_data['resolved_date']) else "",
            row_data['priority'],
            row_data['category'],
            row_data['team'],
            row_data['status'],
            res_hrs_formula,
            sla_met_formula,
            sla_target_formula,
            month_formula,
            breach_formula
        ])

    for col_num in range(1, len(headers_calc) + 1):
        cell = ws_calc.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = font_white_bold
        cell.alignment = align_center

    ws_calc.freeze_panes = "A2"

    # =========================================================================
    # SHEET 3: KPI_Summary
    # =========================================================================
    ws_kpi = wb.create_sheet(title="KPI_Summary")
    ws_kpi.views.sheetView[0].showGridLines = True

    ws_kpi.cell(row=2, column=2, value="Service Desk SLA & KPI Summary").font = font_title
    ws_kpi.cell(row=3, column=2, value="Automated Operational Metrics (Full Year 2024)").font = font_subtitle

    sla_compliance_formula = f'=COUNTIFS(Calc!I2:I{num_rows+1}, "YES") / (COUNTIFS(Cleaned_Data!H2:H{num_rows+1}, "Resolved") + COUNTIFS(Cleaned_Data!H2:H{num_rows+1}, "Closed"))'

    cards = [
        ("Total Ticket Volume", f"=COUNTA(Cleaned_Data!A2:A{num_rows+1})", "#,##0", 5, 2),
        ("Overall SLA Compliance %", sla_compliance_formula, "0.00%", 5, 5),
        ("Avg Resolution Time (Hrs)", f"=AVERAGE(Calc!H2:H{num_rows+1})", "0.00", 5, 8),
        ("Open Backlog Count", f"=COUNTIFS(Cleaned_Data!H2:H{num_rows+1}, \"Open\") + COUNTIFS(Cleaned_Data!H2:H{num_rows+1}, \"Escalated\")", "#,##0", 9, 2),
        ("Avg Customer CSAT (1-5)", f"=AVERAGE(Cleaned_Data!I2:I{num_rows+1})", "0.00", 9, 5),
        ("Total Breached Tickets", f"=COUNTIFS(Calc!I2:I{num_rows+1}, \"NO\")", "#,##0", 9, 8),
    ]

    for label, formula, num_format, r, c in cards:
        ws_kpi.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+2)
        cell_h = ws_kpi.cell(row=r, column=c, value=label)
        cell_h.fill = kpi_header_fill
        cell_h.font = font_kpi_label
        cell_h.alignment = align_center

        ws_kpi.merge_cells(start_row=r+1, start_column=c, end_row=r+2, end_column=c+2)
        cell_v = ws_kpi.cell(row=r+1, column=c, value=formula)
        cell_v.fill = kpi_card_fill
        cell_v.font = font_kpi_num
        cell_v.alignment = align_center
        cell_v.number_format = num_format

        for r_i in range(r, r+3):
            for c_i in range(c, c+3):
                ws_kpi.cell(row=r_i, column=c_i).border = kpi_border

    # Conditional Formatting for Overall SLA Compliance % (Applied to merged card range E6:G7)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name="Calibri", size=20, bold=True, color="006100")
    amber_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    amber_font = Font(name="Calibri", size=20, bold=True, color="9C6500")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(name="Calibri", size=20, bold=True, color="9C0006")

    ws_kpi.conditional_formatting.add("E6:G7", CellIsRule(operator="greaterThanOrEqual", formula=["0.90"], stopIfTrue=True, fill=green_fill, font=green_font))
    ws_kpi.conditional_formatting.add("E6:G7", CellIsRule(operator="between", formula=["0.80", "0.8999"], stopIfTrue=True, fill=amber_fill, font=amber_font))
    ws_kpi.conditional_formatting.add("E6:G7", CellIsRule(operator="lessThan", formula=["0.80"], stopIfTrue=True, fill=red_fill, font=red_font))

    # =========================================================================
    # SHEET 4: Monthly_Trend
    # =========================================================================
    ws_month = wb.create_sheet(title="Monthly_Trend")
    ws_month.views.sheetView[0].showGridLines = True

    ws_month.cell(row=2, column=2, value="Monthly SLA Compliance & Ticket Volume Trend").font = font_title
    
    headers_monthly = ["Month", "Total Intake Volume", "Resolved Tickets", "Tickets Met SLA", "SLA Compliance %", "Avg Resolution Hrs"]
    
    for c_idx, h_text in enumerate(headers_monthly, start=2):
        cell = ws_month.cell(row=4, column=c_idx, value=h_text)
        cell.fill = navy_header_fill
        cell.font = font_white_bold
        cell.alignment = align_center

    months_list = [f"2024-{m:02d}" for m in range(1, 13)]

    for idx, m_code in enumerate(months_list, start=5):
        ws_month.cell(row=idx, column=2, value=m_code).alignment = align_center # Month (B)
        ws_month.cell(row=idx, column=3, value=f'=COUNTIFS(Calc!K$2:K${num_rows+1}, B{idx})').number_format = "#,##0"
        ws_month.cell(row=idx, column=4, value=f'=COUNTIFS(Calc!K$2:K${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Resolved") + COUNTIFS(Calc!K$2:K${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Closed")').number_format = "#,##0"
        ws_month.cell(row=idx, column=5, value=f'=COUNTIFS(Calc!K$2:K${num_rows+1}, B{idx}, Calc!I$2:I${num_rows+1}, "YES")').number_format = "#,##0"
        
        cell_pct = ws_month.cell(row=idx, column=6, value=f'=IFERROR(E{idx}/D{idx}, 0)')
        cell_pct.number_format = "0.00%"
        cell_pct.alignment = align_right
        
        cell_avg = ws_month.cell(row=idx, column=7, value=f'=IFERROR(AVERAGEIFS(Calc!H$2:H${num_rows+1}, Calc!K$2:K${num_rows+1}, B{idx}), 0)')
        cell_avg.number_format = "0.00"
        cell_avg.alignment = align_right

        for c_i in range(2, 8):
            ws_month.cell(row=idx, column=c_i).border = thin_border

    # Cleaned Monthly Trend Charts
    chart_m_sla = create_styled_line_chart(ws_month, 6, 2, "Monthly SLA Compliance % Trend", "Month", "SLA Compliance %")
    ws_month.add_chart(chart_m_sla, "I4")

    chart_m_vol = create_styled_col_chart(ws_month, 3, 2, "Monthly Ticket Intake Volume", "Month", "Ticket Volume", 4, 16, fill_color="1F4E78")
    ws_month.add_chart(chart_m_vol, "I19")

    # =========================================================================
    # SHEET 5: Priority_Category
    # =========================================================================
    ws_pc = wb.create_sheet(title="Priority_Category")
    ws_pc.views.sheetView[0].showGridLines = True

    ws_pc.cell(row=2, column=2, value="SLA Performance by Priority & Category").font = font_title

    # Priority Table
    ws_pc.cell(row=4, column=2, value="Performance by Priority Level").font = font_bold
    p_headers = ["Priority", "SLA Target (Hrs)", "Total Tickets", "Resolved Tickets", "SLA Met", "SLA Compliance %", "Avg Res Hrs"]
    for c_idx, h_text in enumerate(p_headers, start=2):
        cell = ws_pc.cell(row=5, column=c_idx, value=h_text)
        cell.fill = accent_header_fill
        cell.font = font_white_bold
        cell.alignment = align_center

    priorities_list = [("P1", 4), ("P2", 8), ("P3", 24), ("P4", 48)]
    for idx, (p_code, target_h) in enumerate(priorities_list, start=6):
        ws_pc.cell(row=idx, column=2, value=p_code).alignment = align_center
        ws_pc.cell(row=idx, column=3, value=target_h).alignment = align_center
        ws_pc.cell(row=idx, column=4, value=f'=COUNTIFS(Cleaned_Data!D$2:D${num_rows+1}, B{idx})').number_format = "#,##0"
        ws_pc.cell(row=idx, column=5, value=f'=COUNTIFS(Cleaned_Data!D$2:D${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Resolved") + COUNTIFS(Cleaned_Data!D$2:D${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Closed")').number_format = "#,##0"
        ws_pc.cell(row=idx, column=6, value=f'=COUNTIFS(Cleaned_Data!D$2:D${num_rows+1}, B{idx}, Calc!I$2:I${num_rows+1}, "YES")').number_format = "#,##0"
        
        cell_pct = ws_pc.cell(row=idx, column=7, value=f'=IFERROR(F{idx}/E{idx}, 0)')
        cell_pct.number_format = "0.00%"
        
        cell_avg = ws_pc.cell(row=idx, column=8, value=f'=IFERROR(AVERAGEIFS(Calc!H$2:H${num_rows+1}, Cleaned_Data!D$2:D${num_rows+1}, B{idx}), 0)')
        cell_avg.number_format = "0.00"

        for c_i in range(2, 9):
            ws_pc.cell(row=idx, column=c_i).border = thin_border

    # Category Table
    ws_pc.cell(row=12, column=2, value="Performance by Service Category").font = font_bold
    cat_headers = ["Category", "Total Tickets", "Resolved Tickets", "SLA Met", "SLA Compliance %", "Avg Res Hrs"]
    for c_idx, h_text in enumerate(cat_headers, start=2):
        cell = ws_pc.cell(row=13, column=c_idx, value=h_text)
        cell.fill = accent_header_fill
        cell.font = font_white_bold
        cell.alignment = align_center

    categories_list = ["Network", "Software", "Hardware", "Access", "Billing"]
    for idx, cat_name in enumerate(categories_list, start=14):
        ws_pc.cell(row=idx, column=2, value=cat_name).alignment = align_center
        ws_pc.cell(row=idx, column=3, value=f'=COUNTIFS(Cleaned_Data!E$2:E${num_rows+1}, B{idx})').number_format = "#,##0"
        ws_pc.cell(row=idx, column=4, value=f'=COUNTIFS(Cleaned_Data!E$2:E${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Resolved") + COUNTIFS(Cleaned_Data!E$2:E${num_rows+1}, B{idx}, Cleaned_Data!H$2:H${num_rows+1}, "Closed")').number_format = "#,##0"
        ws_pc.cell(row=idx, column=5, value=f'=COUNTIFS(Cleaned_Data!E$2:E${num_rows+1}, B{idx}, Calc!I$2:I${num_rows+1}, "YES")').number_format = "#,##0"
        
        cell_pct = ws_pc.cell(row=idx, column=6, value=f'=IFERROR(E{idx}/D{idx}, 0)')
        cell_pct.number_format = "0.00%"
        
        cell_avg = ws_pc.cell(row=idx, column=7, value=f'=IFERROR(AVERAGEIFS(Calc!H$2:H${num_rows+1}, Cleaned_Data!E$2:E${num_rows+1}, B{idx}), 0)')
        cell_avg.number_format = "0.00"

        for c_i in range(2, 8):
            ws_pc.cell(row=idx, column=c_i).border = thin_border

    # Cleaned Charts on Priority_Category Sheet
    chart_pc_prio = create_styled_col_chart(ws_pc, 7, 2, "SLA Compliance % by Priority Level", "Priority", "SLA Compliance %", 5, 9, fill_color="2F5597")
    ws_pc.add_chart(chart_pc_prio, "J4")

    chart_pc_cat = create_styled_col_chart(ws_pc, 6, 2, "SLA Compliance % by Category", "Category", "SLA Compliance %", 13, 18, fill_color="1F4E78")
    ws_pc.add_chart(chart_pc_cat, "J19")

    # =========================================================================
    # SHEET 6: Dashboard (Executive View - Cleaned Side-by-Side Charts)
    # =========================================================================
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash.cell(row=2, column=2, value="EXECUTIVE SERVICE DESK DASHBOARD").font = font_title
    ws_dash.cell(row=3, column=2, value="Daily / Weekly / Monthly Governance View").font = font_subtitle

    # KPI Summary Cards Layout across Cols B to M
    dash_cards = [
        ("Total Tickets", "=KPI_Summary!B6", "#,##0", 5, 2),
        ("SLA Compliance %", "=KPI_Summary!E6", "0.00%", 5, 5),
        ("Avg Res Time (Hrs)", "=KPI_Summary!H6", "0.00", 5, 8),
        ("Open Backlog", "=KPI_Summary!B10", "#,##0", 5, 11),
    ]

    for label, formula, num_format, r, c in dash_cards:
        ws_dash.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+2)
        cell_h = ws_dash.cell(row=r, column=c, value=label)
        cell_h.fill = navy_header_fill
        cell_h.font = font_white_bold
        cell_h.alignment = align_center

        ws_dash.merge_cells(start_row=r+1, start_column=c, end_row=r+2, end_column=c+2)
        cell_v = ws_dash.cell(row=r+1, column=c, value=formula)
        cell_v.fill = kpi_card_fill
        cell_v.font = font_kpi_num
        cell_v.alignment = align_center
        cell_v.number_format = num_format

        for r_i in range(r, r+3):
            for c_i in range(c, c+3):
                ws_dash.cell(row=r_i, column=c_i).border = kpi_border

    # Side-by-Side Cleaned Executive Charts on Dashboard
    # Chart 1: Line Chart (Cols B to H)
    chart_d_line = create_styled_line_chart(ws_month, 6, 2, "Monthly SLA Compliance Trend", "Month", "SLA Compliance %")
    ws_dash.add_chart(chart_d_line, "B9")

    # Chart 2: Column Chart (Cols I to O)
    chart_d_col = create_styled_col_chart(ws_month, 3, 2, "Monthly Ticket Intake Volume", "Month", "Ticket Volume", 4, 16, fill_color="1F4E78")
    ws_dash.add_chart(chart_d_col, "I9")

    # =========================================================================
    # SHEET 7: Notes & Documentation
    # =========================================================================
    ws_notes = wb.create_sheet(title="Notes")
    ws_notes.views.sheetView[0].showGridLines = True

    ws_notes.cell(row=2, column=2, value="Report Documentation & Metric Definitions").font = font_title
    
    notes_content = [
        ("Metric / Concept", "Definition & Governance Business Standard", "Excel Formula Used"),
        ("Total Tickets", "Count of all support ticket intake records created in 2024.", "=COUNTA(Cleaned_Data!A2:A5001)"),
        ("Resolved Tickets", "Count of tickets with completed status ('Resolved' or 'Closed').", "=COUNTIFS(Cleaned_Data!H2:H5001, \"Resolved\") + COUNTIFS(Cleaned_Data!H2:H5001, \"Closed\")"),
        ("SLA Compliance %", "Standard Definition: Tickets Met SLA / Total Resolved Tickets.", "=COUNTIFS(Calc!I2:I5001, \"YES\") / (COUNTIFS(Cleaned_Data!H2:H5001, \"Resolved\") + COUNTIFS(Cleaned_Data!H2:H5001, \"Closed\"))"),
        ("Avg Resolution Time", "Mean turnaround time (decimal hours) from creation to resolution for resolved tickets.", "=AVERAGE(Calc!H2:H5001)"),
        ("Open Backlog", "Count of active tickets currently in 'Open' or 'Escalated' status.", "=COUNTIFS(Cleaned_Data!H2:H5001, \"Open\") + COUNTIFS(Cleaned_Data!H2:H5001, \"Escalated\")"),
        ("Avg CSAT", "Average customer satisfaction rating on a 1-5 scale for resolved tickets.", "=AVERAGE(Cleaned_Data!I2:I5001)"),
        ("SLA Target (P1)", "Critical Incident Target = 4 Hours.", "Hardcoded SLA Target Rule"),
        ("SLA Target (P2)", "High Priority Target = 8 Hours.", "Hardcoded SLA Target Rule"),
        ("SLA Target (P3)", "Medium Priority Target = 24 Hours.", "Hardcoded SLA Target Rule"),
        ("SLA Target (P4)", "Low Priority Target = 48 Hours.", "Hardcoded SLA Target Rule"),
        ("Conditional Formatting", "Green: SLA >= 90% | Amber: 80% <= SLA < 90% | Red: SLA < 80%", "Excel CellIsRule Conditional Formatting on merged E6:G7"),
    ]

    for r_idx, (m_val, d_val, f_val) in enumerate(notes_content, start=4):
        c1 = ws_notes.cell(row=r_idx, column=2, value=m_val)
        c2 = ws_notes.cell(row=r_idx, column=3, value=d_val)
        c3 = ws_notes.cell(row=r_idx, column=4, value=f_val)

        if r_idx == 4:
            c1.fill = navy_header_fill; c1.font = font_white_bold; c1.alignment = align_center
            c2.fill = navy_header_fill; c2.font = font_white_bold; c2.alignment = align_center
            c3.fill = navy_header_fill; c3.font = font_white_bold; c3.alignment = align_center
        else:
            c1.font = font_bold; c1.border = thin_border
            c2.font = font_regular; c2.border = thin_border
            c3.font = font_regular; c3.border = thin_border

    # Auto-adjust Column Widths across all sheets safely
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if val_str.startswith('='):
                    val_str = "Formula_Result"
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    os.makedirs(os.path.dirname(report_output_path), exist_ok=True)

    # Save with permission retry handling if file is locked in Excel
    saved = False
    attempts = 0
    while not saved and attempts < 3:
        try:
            wb.save(report_output_path)
            saved = True
            print(f"Success! Excel SLA & KPI Report saved with cleaned charts at: {report_output_path}")
        except PermissionError:
            attempts += 1
            print(f"Warning: File {report_output_path} is locked by Excel. Attempting retry {attempts}/3...")
            time.sleep(2)

    if not saved:
        fallback_path = os.path.join(base_dir, 'output', 'SLA_KPI_Report_updated.xlsx')
        wb.save(fallback_path)
        print(f"Notice: Main file was locked by Excel, saved updated version to: {fallback_path}")

    print("-" * 60)

if __name__ == '__main__':
    main()
