import pandas as pd
df = pd.read_excel("sales.xlsx")
print(df.shape)
print(df[df["amount"] > 100])
print(df.sort_values("amount", ascending=False))
