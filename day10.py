import pandas as pd

df = pd.read_excel("final_sales.xlsx")

totals = df.groupby("city")["amount"].sum().reset_index()
totals = totals.sort_values("amount", ascending=False)
print(totals)

totals.to_excel("rapport_par_ville.xlsx", index=False)
