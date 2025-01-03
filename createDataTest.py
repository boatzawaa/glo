import pandas as pd
import datetime

df = pd.read_excel('dataTest/patternTest.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5000)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/patternTest100000row.xlsx",index=False)