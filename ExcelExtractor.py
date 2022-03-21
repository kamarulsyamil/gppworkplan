from itertools import dropwhile
from operator import index
from turtle import right
import pandas as pd
import numpy as np
import glob

# to do list
# 1. convert the daily excel to dataframe.
# 2. process the dataframe.
# 3. put the dataframe into the consolidated excel.

fact_data = pd.DataFrame()

f = "C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\Production Line Arrangement of 2022.xlsx"


def tableA(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)

    # night shift
    df = xl.parse(2)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-10:max_row, max_col-8:max_col]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    return df5


def tableB(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)
    df = xl.parse(2)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-10:max_row, max_col-18:max_col-9]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    return df5


def CCC4Df():
    # On duty time
    global fName

    rightDf = tableA(f)

    #This is CCC2
    #leftDf = tableB(f)

    rightDf.columns = ['Line', 'Start Time', 'End Time', '4', 'UPH']
    #leftDf.columns = ['Line', 'Time', '3', '4', 'UPH']

    fNameDf = rightDf['Line'].str.extract(r'(CCC[2-4])')
    dateDf = rightDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')

    fName = fNameDf.loc[0].iat[0]
    date = dateDf.loc[0].iat[0]

    print(date)

    print("The name of the factory is: ", fName)

    if not rightDf[rightDf['Line'].str.contains("Next Night-Shift")].empty:
        print("Start shift.")

        df1 = rightDf.drop(columns=['4'])

        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    elif not rightDf[rightDf['Line'].str.contains("Today")].empty:
        print("Means that the DF is on off duty time time or end shift.")

    return df2.reset_index(drop=True)


def CCC2Df():
    # On duty time
    fName = ''

    #This is CCC2
    leftDf = tableB(f)

    leftDf.columns = ['Line', 'Time', 'End Time', '4', 'UPH']

    fNameDf = leftDf['Line'].str.extract(r'(CCC[2-4])')
    dateDf = leftDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')

    fName = fNameDf.loc[0].iat[0]
    date = dateDf.loc[0].iat[0]

    print(date)

    print("The name of the factory is: ", fName)

    if not leftDf[leftDf['Line'].str.contains("Next Night-Shift")].empty:
        print("Start shift.")
        df1 = leftDf.drop(columns=['4'])
        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    elif not leftDf[leftDf['Line'].str.contains("Today")].empty:
        print("Means that the DF is on off duty time time or end shift.")

    return df2.reset_index(drop=True)


# removed duplicate
# df6 = CCC4Df()['Line'].drop_duplicates()
# df7 = df6.dropna(how='all')

# s = df7.sort_index()
# df8 = s.to_frame().T

# print(filterDf())
# print(tableB(f))

print(CCC4Df())
print(CCC2Df())
