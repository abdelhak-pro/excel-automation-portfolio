import pandas as pd
import glob
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# 1. Fusionner
files = sorted(glob.glob("data/*.xlsx"))
frames = []
for f in files:
    frames.append(pd.read_excel(f))
df = pd.concat(frames, ignore_index=True)

# 2. Nettoyer
df["product"] = df["product"].str.strip().str.title()
df = df.drop_duplicates()
df = df.dropna()

# 3. Rapports
by_product = df.groupby("product")["amount"].sum().reset_index()
by_city = df.groupby("city")["amount"].sum().reset_index()

# 4. Ecrire
with pd.ExcelWriter("rapport_final.xlsx") as writer:
    by_product.to_excel(writer, sheet_name="Par produit", index=False)
    by_city.to_excel(writer, sheet_name="Par ville", index=False)

# 5. Formater
wb = load_workbook("rapport_final.xlsx")
for ws in wb:
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="14509E")
wb.save("rapport_final.xlsx")
print("Rapport cree")
