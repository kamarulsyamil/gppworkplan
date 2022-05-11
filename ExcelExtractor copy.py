from itertools import dropwhile
from operator import index
from turtle import right
from cv2 import threshold
import pandas as pd
import numpy as np
import glob

# to do list
# 1. convert the daily excel to dataframe.
# 2. process the dataframe.
# 3. put the dataframe into the consolidated excel.

fact_data = pd.DataFrame()

f = r"sources\Production Line Arrangement of 2022.xlsx"

# process excel shifts of CCC4


xl = pd.ExcelFile(f)

# day shift C2 side
df = xl.parse(sheet_name=0, usecols="B:I")

# print(df)

# day shift C4 side
df = xl.parse(sheet_name=0, usecols="K:R")
df1 = df.dropna(how='all')
df2 = df1.dropna(how='all', axis=1)
df3 = df1.reset_index(drop=True)

max_col = df3.shape[1]

col_list = [str(x) for x in range(0, df3.shape[1])]


print(col_list)
df3.columns = col_list

# print(df2[df2['1']:])

if not df3[df3['0'].str.contains("CCC4", na=False)].empty:

    date = df3[df3['0'] == "CCC4 Next Day Shift (May-09)"].index.values

    dfList = df3[df3['0'] == "Total HC:"].index.values

    print(dfList[dfList > 2284][0])

    #print(df3.iloc[date[0]:2285, :])
    #print(dfList, date)
