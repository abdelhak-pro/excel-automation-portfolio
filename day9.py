import pandas as pd

df = pd.read_excel("final_sales.xlsx")

totals = df.groupby("product")["amount"].sum().reset_index()
print(totals)

totals.to_excel("rapport_par_prodit.xlsx", index=False)
