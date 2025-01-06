import pandas as pd
import datetime

'''
df = pd.read_excel('dataTest/patternTest.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5000)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/patternTest100000row.xlsx",index=False)
'''
df = pd.read_excel('dataTest/dataTest48000.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/dataTestDup240000row.xlsx",index=False)

df = pd.read_excel('dataTest/dataTest46000.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/dataTestDup230000row.xlsx",index=False)

df = pd.read_excel('dataTest/dataTest44000.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/dataTestDup220000row.xlsx",index=False)

df = pd.read_excel('dataTest/dataTest42000.xlsx',dtype=str)
df_dup = df.loc[df.index.repeat(5)].reset_index(drop=True)
print(df_dup)
df_dup.to_excel("dataTest/dataTestDup210000row.xlsx",index=False)