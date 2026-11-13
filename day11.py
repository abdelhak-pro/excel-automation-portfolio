import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

df = pd.read_excel("final_sales.xlsx")
totals = df.groupby("city")["amount"].sum().reset_index()
totals.to_excel("report_formatted.xlsx", index=False)

wb = load_workbook("report_formatted.xlsx")
ws = wb.active
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 15
for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="14509E")
wb.save("report_formatted.xlsx")
print("Rapport créé")
