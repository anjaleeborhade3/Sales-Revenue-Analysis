import os
import random
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter

random.seed(42)

# -----------------------------
# 1. Create realistic dataset
# -----------------------------
products = {
    "Laptop": (55000, 85000),
    "Smartphone": (18000, 55000),
    "Tablet": (15000, 40000),
    "Headphones": (1500, 12000),
    "Monitor": (9000, 30000),
    "Keyboard": (800, 5000),
    "Mouse": (500, 3000),
}

regions = ["North", "South", "East", "West"]
categories = {
    "Laptop": "Electronics",
    "Smartphone": "Electronics",
    "Tablet": "Electronics",
    "Headphones": "Accessories",
    "Monitor": "Accessories",
    "Keyboard": "Accessories",
    "Mouse": "Accessories",
}

rows = []

for order_id in range(1001, 1251):
    product = random.choice(list(products.keys()))
    low, high = products[product]
    price = random.randint(low, high)
    quantity = random.randint(1, 5)
    discount = random.choice([0, 0.05, 0.10, 0.15])
    revenue = price * quantity * (1 - discount)
    cost = price * quantity * random.uniform(0.65, 0.82)
    profit = revenue - cost

    date = pd.Timestamp("2026-01-01") + pd.Timedelta(
        days=random.randint(0, 299)
    )

    rows.append([
        order_id,
        date,
        product,
        categories[product],
        random.choice(regions),
        quantity,
        price,
        discount,
        round(revenue, 2),
        round(profit, 2)
    ])

df = pd.DataFrame(rows, columns=[
    "Order_ID",
    "Order_Date",
    "Product",
    "Category",
    "Region",
    "Quantity",
    "Unit_Price",
    "Discount",
    "Revenue",
    "Profit"
])

# -----------------------------
# 2. Save raw CSV
# -----------------------------
df.to_csv("sales_data.csv", index=False)

# -----------------------------
# 3. Create Excel workbook
# -----------------------------
wb = Workbook()

ws = wb.active
ws.title = "Sales Data"

# Header
headers = list(df.columns)

for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center")

# Data
for row in df.itertuples(index=False):
    ws.append(list(row))

# Formatting
for row in ws.iter_rows(min_row=2):
    row[1].number_format = "dd-mm-yyyy"
    row[6].number_format = '₹#,##0'
    row[7].number_format = '0%'
    row[8].number_format = '₹#,##0.00'
    row[9].number_format = '₹#,##0.00'

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

widths = {
    "A": 12, "B": 14, "C": 16, "D": 16, "E": 12,
    "F": 12, "G": 14, "H": 12, "I": 16, "J": 16
}

for col, width in widths.items():
    ws.column_dimensions[col].width = width

# -----------------------------
# 4. Summary sheet
# -----------------------------
summary = wb.create_sheet("Analysis")

summary["A1"] = "SALES & REVENUE ANALYSIS"
summary["A1"].font = Font(size=20, bold=True)
summary["A3"] = "Total Revenue"
summary["B3"] = "=SUM('Sales Data'!I2:I251)"

summary["A4"] = "Total Profit"
summary["B4"] = "=SUM('Sales Data'!J2:J251)"

summary["A5"] = "Total Orders"
summary["B5"] = "=COUNTA('Sales Data'!A2:A251)"

summary["A6"] = "Total Units Sold"
summary["B6"] = "=SUM('Sales Data'!F2:F251)"

summary["A7"] = "Average Order Revenue"
summary["B7"] = "=AVERAGE('Sales Data'!I2:I251)"

for cell in ["A3", "A4", "A5", "A6", "A7"]:
    summary[cell].font = Font(bold=True)

for cell in ["B3", "B4", "B7"]:
    summary[cell].number_format = '₹#,##0.00'

# -----------------------------
# 5. Region Analysis
# -----------------------------
summary["A10"] = "Region"
summary["B10"] = "Revenue"
summary["C10"] = "Profit"

for c in ["A10", "B10", "C10"]:
    summary[c].font = Font(bold=True, color="FFFFFF")
    summary[c].fill = PatternFill("solid", fgColor="1F4E78")

for i, region in enumerate(regions, 11):
    summary[f"A{i}"] = region
    summary[f"B{i}"] = f'=SUMIF(\'Sales Data\'!E:E,A{i},\'Sales Data\'!I:I)'
    summary[f"C{i}"] = f'=SUMIF(\'Sales Data\'!E:E,A{i},\'Sales Data\'!J:J)'
    summary[f"B{i}"].number_format = '₹#,##0.00'
    summary[f"C{i}"].number_format = '₹#,##0.00'

# -----------------------------
# 6. Category Analysis
# -----------------------------
summary["E10"] = "Category"
summary["F10"] = "Revenue"
summary["G10"] = "Profit"

for c in ["E10", "F10", "G10"]:
    summary[c].font = Font(bold=True, color="FFFFFF")
    summary[c].fill = PatternFill("solid", fgColor="1F4E78")

categories_list = sorted(df["Category"].unique())

for i, category in enumerate(categories_list, 11):
    summary[f"E{i}"] = category
    summary[f"F{i}"] = f'=SUMIF(\'Sales Data\'!D:D,E{i},\'Sales Data\'!I:I)'
    summary[f"G{i}"] = f'=SUMIF(\'Sales Data\'!D:D,E{i},\'Sales Data\'!J:J)'
    summary[f"F{i}"].number_format = '₹#,##0.00'
    summary[f"G{i}"].number_format = '₹#,##0.00'

# -----------------------------
# 7. Monthly Analysis
# -----------------------------
summary["A18"] = "Month"
summary["B18"] = "Revenue"

for c in ["A18", "B18"]:
    summary[c].font = Font(bold=True, color="FFFFFF")
    summary[c].fill = PatternFill("solid", fgColor="1F4E78")

months = pd.date_range("2026-01-01", "2026-10-01", freq="MS")

for i, month in enumerate(months, 19):
    summary[f"A{i}"] = month
    summary[f"A{i}"].number_format = "mmm-yyyy"
    summary[f"B{i}"] = (
        f'=SUMIFS(\'Sales Data\'!I:I,'
        f'\'Sales Data\'!B:B,">="&A{i},'
        f'\'Sales Data\'!B:B,"<"&EDATE(A{i},1))'
    )
    summary[f"B{i}"].number_format = '₹#,##0.00'

# -----------------------------
# 8. Charts
# -----------------------------
bar = BarChart()
bar.title = "Revenue by Region"
bar.y_axis.title = "Revenue"
bar.x_axis.title = "Region"

data = Reference(summary, min_col=2, min_row=10, max_row=14)
cats = Reference(summary, min_col=1, min_row=11, max_row=14)

bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.height = 8
bar.width = 16
bar.style = 10
bar.legend.position = "b"

summary.add_chart(bar, "I3")

line = LineChart()
line.title = "Monthly Revenue Trend"
line.y_axis.title = "Revenue"
line.x_axis.title = "Month"

data = Reference(summary, min_col=2, min_row=18, max_row=28)
cats = Reference(summary, min_col=1, min_row=19, max_row=28)

line.add_data(data, titles_from_data=True)
line.set_categories(cats)
line.height = 7
line.width = 12

line.style = 10
line.legend.position = "b"
summary.add_chart(line, "I19")

pie = PieChart()
pie.title = "Revenue by Category"

data = Reference(summary, min_col=6, min_row=10, max_row=12)
labels = Reference(summary, min_col=5, min_row=11, max_row=12)

pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.height = 8
pie.width = 16
pie.style = 10
summary.add_chart(pie, "I35")

# -----------------------------
# 9. Professional workbook formatting
# -----------------------------
navy, blue, teal = "17365D", "1F4E78", "0F6B78"
pale, white, ink, gray = "EAF2F8", "FFFFFF", "243447", "D6E0E8"
thin = Side(style="thin", color=gray)

# Transaction table
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:J{ws.max_row}"
ws.row_dimensions[1].height = 30
for cell in ws[1]:
    cell.fill = PatternFill("solid", fgColor=navy)
    cell.font = Font(name="Aptos", bold=True, color=white, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(bottom=Side(style="medium", color=teal))
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = Font(name="Aptos", size=10, color=ink)
        cell.alignment = Alignment(vertical="center", horizontal="right" if cell.column in (1, 6, 7, 8, 9, 10) else "left")
        cell.border = Border(bottom=Side(style="hair", color="E7EDF2"))
    row[1].number_format = "dd-mmm-yyyy"
    row[6].number_format = "₹#,##0.00"
    row[7].number_format = "0%"
    row[8].number_format = "₹#,##0.00"
    row[9].number_format = "₹#,##0.00"
for col, width in {"A": 12, "B": 15, "C": 18, "D": 17, "E": 13, "F": 12, "G": 16, "H": 12, "I": 18, "J": 18}.items():
    ws.column_dimensions[col].width = width
transactions = Table(displayName="SalesTransactions", ref=f"A1:J{ws.max_row}")
transactions.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
ws.add_table(transactions)
ws.sheet_properties.tabColor = blue

# Dashboard title and KPI cards
summary.sheet_view.showGridLines = False
summary.sheet_view.zoomScale = 85
summary.sheet_properties.tabColor = teal
summary.merge_cells("A1:G1")
summary["A1"].fill = PatternFill("solid", fgColor=navy)
summary["A1"].font = Font(name="Aptos Display", size=20, bold=True, color=white)
summary["A1"].alignment = Alignment(vertical="center", horizontal="left", indent=1)
summary.row_dimensions[1].height = 42
for row_num in range(3, 8):
    summary.row_dimensions[row_num].height = 30
    for col_num in (1, 2):
        cell = summary.cell(row_num, col_num)
        cell.fill = PatternFill("solid", fgColor=pale if row_num % 2 else "F4F8FB")
        cell.border = Border(bottom=thin)
    summary.cell(row_num, 1).font = Font(name="Aptos", bold=True, color=blue, size=10)
    summary.cell(row_num, 1).alignment = Alignment(vertical="center", indent=1)
    summary.cell(row_num, 2).font = Font(name="Aptos Display", bold=True, color=teal, size=13)
    summary.cell(row_num, 2).alignment = Alignment(horizontal="right", vertical="center", indent=1)
for addr in ("B3", "B4", "B7"):
    summary[addr].number_format = "₹#,##0.00"
for addr in ("B5", "B6"):
    summary[addr].number_format = "#,##0"
for row_num in (10, 18):
    summary.row_dimensions[row_num].height = 25
    for cell in summary[row_num]:
        if cell.column <= 7 and cell.value is not None:
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.font = Font(name="Aptos", bold=True, color=white, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=Side(style="medium", color=teal))
for row_num in list(range(11, 15)) + [11, 12] + list(range(19, 29)):
    for cell in summary[row_num]:
        if cell.value is not None:
            cell.font = Font(name="Aptos", size=10, color=ink)
            cell.border = Border(bottom=Side(style="hair", color="E7EDF2"))
            if cell.column in (2, 3, 6, 7):
                cell.number_format = "₹#,##0.00"
                cell.alignment = Alignment(horizontal="right")
for row_num in range(19, 29):
    summary.cell(row_num, 1).number_format = "mmm yyyy"
summary.conditional_formatting.add("B11:C14", ColorScaleRule(start_type="min", start_color="FCE8E6", mid_type="percentile", mid_value=50, mid_color="FFF4D6", end_type="max", end_color="D9EAD3"))
summary.conditional_formatting.add("F11:G12", ColorScaleRule(start_type="min", start_color="FCE8E6", mid_type="percentile", mid_value=50, mid_color="FFF4D6", end_type="max", end_color="D9EAD3"))
for col, width in {"A": 22, "B": 18, "C": 18, "D": 4, "E": 18, "F": 18, "G": 18, "H": 3, "I": 3}.items():
    summary.column_dimensions[col].width = width
summary.freeze_panes = "A3"
summary.sheet_properties.pageSetUpPr.fitToPage = True
summary.page_setup.fitToWidth = 1
summary.page_setup.fitToHeight = 2
summary.page_setup.orientation = "landscape"
summary.print_area = "A1:X50"

# -----------------------------
# 10. Save workbook
# -----------------------------
output = "Sales_Revenue_Analysis.xlsx"
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = "auto"
wb.save(output)

print()
print("=" * 55)
print("PROJECT CREATED SUCCESSFULLY")
print("=" * 55)
print(f"Folder : {os.getcwd()}")
print(f"Excel  : {output}")
print("CSV    : sales_data.csv")
print("README : README.md")
print("=" * 55)
