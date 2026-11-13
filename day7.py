import pandas as pd
import glob
files = sorted(glob.glob("data/*.xlsx"))
print(files)
frames = []
for f in files:
    frames.append(pd.read_excel(f))
df = pd.concat(frames, ignore_index=True)
print(df)
print(df.shape)
print(df["amount"].sum())
df.to_excel("merged_sales.xlsx", index=False)
