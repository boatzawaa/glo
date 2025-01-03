import pandas as pd
from pathlib import Path

#read data test
df = pd.read_excel('dataTest/dataTestDup.xlsx',dtype=str)
#print("The dataframe is:")
#print(df)

#duplicate data
row_duplicate = 5 #input num row
df_dup = df.loc[df.index.repeat(row_duplicate)].reset_index(drop=True)
print(df_dup)
#write dup data to new file
df_dup.to_excel("fileDuplicateData/newDuplicate2.xlsx",index=False)

"""
#example data of Glo
glo_df = pd.DataFrame({'B':[7777]}) #input num of Glo

#duplicate data of Glo and input num row for dup
glo_df_dup = glo_df.loc[glo_df.index.repeat(row_duplicate)].reset_index(drop=True)
#print(glo_df_dup)

#manage data
row_start_stop = 2000 #input num row
result_df1 = df_dup.iloc[row_start_stop:] #input num row for start data
result_df2 = df_dup.iloc[0:row_start_stop] #end data
#print(result_df1)
#print(result_df2)

#concat data result_df1,glo_df_dup,result_df2
result = pd.concat([result_df1,glo_df_dup,result_df2], ignore_index=True)
#print("concat: ",result)

#write dup data to new file
result.to_excel("fileDuplicateData/newDuplicate.xlsx",index=False)

#only column B
bCol = result.loc[:,['B']]
print("bCol: ",bCol)

#read pattern
pattern_df = pd.read_excel('dataTest/patternTest100000row.xlsx',dtype=str)
print("pattern_df: ",pattern_df)

#add column B to pattern DF
pattern_df['newB'] = bCol
#print("pattern_df: ",pattern_df)

#write new pattern file
pattern_df.to_excel("final/finalFile.xlsx",index=False)

#splite finalFile to mutiple csv 25 file
row_per_file = 10000 #input num row
for i in range(0, len(pattern_df), row_per_file):
    #print("len: ",len(pattern_df))
    #print("i: ",i," i+row_per_file: ",i+row_per_file)
    df_subset = pattern_df.iloc[i:i+row_per_file]
    #print("splite :",df_subset)
    df_subset.to_csv(f'csvFile/output_part_{i//row_per_file + 1}.csv', index=False)
    df_subset.to_csv(f'txtFile/output_part_{i//row_per_file + 1}.txt', index=False , header=False)
    #print(f"Saved csvFile/output_part_{i//row_per_file + 1}.csv")
"""
