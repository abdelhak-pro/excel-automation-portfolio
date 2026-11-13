import pandas as pd
import glob
files = sorted(glob.glob("data/*.xlsx"))
frames = []
for f in files:
    frames.append(pd.read_excel(f))
df = pd.concat(frames, ignore_index=True)
df["product"] = df["product"].str.strip().str.title()
df = df.drop_duplicates()
df = df.dropna()
print(df)
print(df.shape)
print(df["product"].unique())
df.to_excel("final_sales.xlsx", index=False)
