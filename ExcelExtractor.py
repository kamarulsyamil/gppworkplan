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

# process excel shifts of CCC4


def day_CCC4(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)

    # night shift
    df = xl.parse(0)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-12:max_row, max_col-8:max_col]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    return df5


def night_CCC4(filepath):
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

# ----------------------------------------------
# process excel shifts of CCC2


def day_CCC2(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)

    # night shift
    df = xl.parse(0)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-10:max_row, max_col-18:max_col-8]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    return df5


def night_CCC2(filepath):
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

# ----------------------------------------------
# process dataframe shifts of CCC4


def day_CCC4Df():
    # On duty time
    global fName

    rightDf = day_CCC4(f)

    # print(night_CCC4(f))

    #leftDf = night_CCC2(f)

    if not rightDf[rightDf['11'].str.contains("Next Day Shift")].empty:
        rightDf.columns = ['Line', 'Start Time', 'End Time', 'UPH']

        fNameDf = rightDf['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDf['Line'].str.extract(r'([A-Z][a-z][a-z].[0-3][0-9])')
        isNight = rightDf[rightDf['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'
        print("Start shift.")

        #df1 = rightDf.drop(columns=['4'])

        df2 = rightDf.drop([0, 1, (rightDf.shape[0])-1])

    # Means that the DF is on off duty time time or end shift.
    # process df

    elif not rightDf[rightDf['11'].str.contains("Today")].empty:

        rightDf.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = rightDf['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')
        isNight = rightDf[rightDf['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'
        print("End shift.")

        #df1 = rightDf.drop(columns=['Line', 'OT', 'HC'])

        df2 = rightDf.drop([0, 1, (rightDf.shape[0])-1])

        # print(df2)

    return df2.reset_index(drop=True), fName, date, shift, isNight


def night_CCC4Df():
    # On duty time
    global fName

    rightDf = night_CCC4(f)

    # print(night_CCC4(f))

    #leftDf = night_CCC2(f)

    if not rightDf[rightDf['12'].str.contains("Next Night-Shift")].empty:
        rightDf.columns = ['Line', 'Start Time', 'End Time', '4', 'UPH']

        fNameDf = rightDf['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')
        isNight = rightDf[rightDf['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'
        print("Start shift.")

        df1 = rightDf.drop(columns=['4'])

        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    # Means that the DF is on off duty time time or end shift.
    # process df

    elif not rightDf[rightDf['12'].str.contains("Today")].empty:

        rightDf.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = rightDf['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')
        isNight = rightDf[rightDf['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'
        print("End shift.")

        #df1 = rightDf.drop(columns=['Line', 'OT', 'HC'])

        df2 = rightDf.drop([0, 1, (rightDf.shape[0])-1])

        print(df2)

    return df2.reset_index(drop=True), fName, date, shift, isNight

# ----------------------------------------------
# process dataframe shifts of CCC2


def night_CCC2Df():
    # On duty time
    fName = ''

    #This is CCC2
    leftDf = night_CCC2(f)

    leftDf.columns = ['Line', 'Time', 'End Time', '4', 'UPH']

    fNameDf = leftDf['Line'].str.extract(r'(CCC[2-4])')
    dateDf = leftDf['Line'].str.extract(r'([A-Z][a-z][a-z]-[0-3][0-9])')

    fName = fNameDf.loc[0].iat[0]
    date = dateDf.loc[0].iat[0]

    # print(date)

    # print("The name of the factory is: ", fName)

    if not leftDf[leftDf['Line'].str.contains("Next Night-Shift")].empty:
        print("Start shift.")
        df1 = leftDf.drop(columns=['4'])
        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    elif not leftDf[leftDf['Line'].str.contains("Today")].empty:
        print("Means that the DF is on off duty time time or end shift.")

    return df2.reset_index(drop=True), fName, date


# removed duplicate
# df6 = CCC4Df()['Line'].drop_duplicates()
# df7 = df6.dropna(how='all')

# s = df7.sort_index()
# df8 = s.to_frame().T

# print(filterDf())
# print(night_CCC2(f))

day_CCC4(f)
# print(CCC2Df())
