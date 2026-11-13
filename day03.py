import pandas as pd
df = pd.read_excel("sales.xlsx")
print(df)
print(df["amount"].sum())
print(df["city"].unique())
