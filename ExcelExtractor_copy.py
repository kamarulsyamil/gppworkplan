from calendar import month
from itertools import dropwhile
from operator import index
from turtle import right
#from cv2 import threshold
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


def CCC2Day():
    shift = 'day'

    # day shift C4 side
    df = xl.parse(sheet_name=0, usecols="B:I")
    df1 = df.dropna(how='all')
    df2 = df1.dropna(how='all', axis=1)
    df3 = df2.reset_index(drop=True)

    max_col = df3.shape[1]

    col_list = [str(x) for x in range(0, df3.shape[1])]

    # print(col_list)
    df3.columns = col_list

    # print(df2[df2['1']:])

    # if not df3[df3['0'].str.contains("CCC4", na=False)].empty:

    num = "May.10"

    date = df3[df3['0'].str.contains(
        "(?=.*%s)(?=.*Day Shift).*" % num, na=False)].index.values

    dfList = df3[df3['0'] == "Total HC:"].index.values

    for i in range(0, 2):
        C4Df = df3.iloc[date[i]+1:dfList[dfList > date[i]][0], :]
        header = C4Df.iloc[0]
        C4Df = C4Df[1:]
        C4Df.columns = header

        if i == 0:
            on_duty = C4Df.dropna(how="all", axis=1)
        elif i == 1:
            off_duty = C4Df.dropna(how="all", axis=1)

    return on_duty, off_duty, shift, num


def CCC4Day():
    shift = 'day'

    # day shift C4 side
    df = xl.parse(sheet_name=0, usecols="K:R")
    df1 = df.dropna(how='all')
    df2 = df1.dropna(how='all', axis=1)
    df3 = df2.reset_index(drop=True)

    max_col = df3.shape[1]

    col_list = [str(x) for x in range(0, df3.shape[1])]

    # print(col_list)
    df3.columns = col_list

    # print(df2[df2['1']:])

    # if not df3[df3['0'].str.contains("CCC4", na=False)].empty:
    num = "Apr.20"
    date = df3[df3['0'].str.contains(
        "(?=.*%s)(?=.*Day Shift).*" % num, na=False)].index.values

    dfList = df3[df3['0'] == "Total HC:"].index.values

    for i in range(0, 2):
        C4Df = df3.iloc[date[i]+1:dfList[dfList > date[i]][0], :]
        header = C4Df.iloc[0]
        C4Df = C4Df[1:]
        C4Df.columns = header

        if i == 0:
            on_duty = C4Df.dropna(how="all", axis=1)
        elif i == 1:
            off_duty = C4Df.dropna(how="all", axis=1)

    return on_duty, off_duty, shift, num


def CCC2Night():
    shift = 'night'

    # day shift C4 side
    df = xl.parse(sheet_name=2, usecols="C:K")
    df1 = df.dropna(how='all')
    df2 = df1.dropna(how='all', axis=1)
    df3 = df2.reset_index(drop=True)

    max_col = df3.shape[1]

    col_list = [str(x) for x in range(0, df3.shape[1])]

    # print(col_list)
    df3.columns = col_list

    # print(df2[df2['1']:])

    # if not df3[df3['0'].str.contains("CCC4", na=False)].empty:

    num = 'May-06'
    date = df3[df3['0'].str.contains(
        "(?=.*%s)(?=.*Night-Shift).*" % num, na=False)].index.values

    dfList = df3[df3['0'] == "Total HC:"].index.values

    for i in range(0, 2):
        C4Df = df3.iloc[date[i]+1:dfList[dfList > date[i]][0], :]
        header = C4Df.iloc[0]
        C4Df = C4Df[1:]
        C4Df.columns = header

        if i == 0:
            on_duty = C4Df.dropna(how="all", axis=1)
        elif i == 1:
            off_duty = C4Df.dropna(how="all", axis=1)

    return on_duty, off_duty, shift, num


def CCC4Night():
    shift = 'night'

    # day shift C4 side
    df = xl.parse(sheet_name=2, usecols="L:S")
    df1 = df.dropna(how='all')
    df2 = df1.dropna(how='all', axis=1)
    df3 = df2.reset_index(drop=True)

    max_col = df3.shape[1]

    col_list = [str(x) for x in range(0, df3.shape[1])]

    # print(col_list)
    df3.columns = col_list

    # print(df2[df2['1']:])

    # if not df3[df3['0'].str.contains("CCC4", na=False)].empty:

    mnth = 'Apr'
    day = '20'

    num = 'Apr-20'

    date = df3[df3['0'].str.contains(
        "(?=.*%s)(?=.*%s)(?i:.*Night-shift).*" % (mnth, day), na=False)].index.values

    dfList = df3[df3['0'] == "Total HC:"].index.values

    for i in range(0, 2):
        C4Df = df3.iloc[date[i]+1:dfList[dfList > date[i]][0], :]
        header = C4Df.iloc[0]
        C4Df = C4Df[1:]
        C4Df.columns = header

        if i == 0:
            on_duty = C4Df.dropna(how="all", axis=1)
        elif i == 1:
            off_duty = C4Df.dropna(how="all", axis=1)

    return on_duty, off_duty, shift, num


CCC4Night()
