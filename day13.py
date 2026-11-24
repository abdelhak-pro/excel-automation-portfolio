import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

df = pd.read_excel("test_final.xlsx")

df["product"] = df["product"].str.strip()
df["product"] = df["product"].str.title()
df["city"] = df["city"].str.strip()
df["city"] = df["city"].str.title()
df = df.drop_duplicates()
df["product"] = df["product"].fillna("Vide")
df["amount"] = df["amount"].fillna(0)
print(df)

by_product = df.groupby("product")["amount"].sum().reset_index()
by_city = df.groupby("city")["amount"].sum().reset_index()

with pd.ExcelWriter("rapport_test.xlsx") as writer:
    by_product.to_excel(writer, sheet_name="Par product", index=False)
    by_city.to_excel(writer, sheet_name="Par ville", index=False)
    df.to_excel(writer, sheet_name="Données nettoyées", index=False)

wb = load_workbook("rapport_test.xlsx")

thin = Side(style="thin")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for ws in wb:
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["c"].width = 15
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="14509E")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

    if ws.title == "Données nettoyées":
        ws.auto_filter.ref = ws.dimensions      

wb.save("rapport_test.xlsx")
print(by_product)
print(by_city)
