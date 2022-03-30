from cmath import nan
import os
from numpy import NaN
import pandas as pd
import win32com.client
import datetime

date = datetime.datetime.today()


def getTableEmail():
    factName = 'ICC'

    email_dir = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\ICC Shift timings_.msg"
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")

    msg = outlook.OpenSharedItem(email_dir)

    # dataframe of APCC table from email
    data = pd.read_html(msg.HTMLBody)

    #data2 = data1.drop_duplicates(subset='0')

    if factName == 'APCC':
        # drop duplicates and NA
        data1 = data[0].dropna(axis=1, how='all', thresh=3)

        data1.columns = ['Date', 'Line', 'Frontend', 'Backend']

        data2 = data1  # .drop_duplicates()
        # print(data2)

        if not data2[data2['Date'].astype(str).str.contains("Date")].empty:
            data3 = data2.drop(data2.index[range(5)])
            # print(data3.reset_index(drop=True))
            df = data3.reset_index(drop=True)

    elif factName == 'ICC':
        # change header of datarframe
        new_header = data[3].iloc[0]
        df = data[3][1:]
        df.columns = new_header
        print(df)

    # separate tables by date

    return df, factName
    # date_format = date.strftime('%d-%b')

# APCC


def APCClogic():
    date = '22-Feb'

    first_shift = []
    second_shift = []

    df, factName = getTableEmail()

    df['group_no'] = df.isnull().all(axis=1).cumsum()

    # dictionary
    d = {i: df.loc[df.group_no == i, ['Date', 'Line', 'Frontend', 'Backend']]
         for i in range(0, df.group_no.iat[-1])}

    # k v = key value
    new_d = {k: v for (k, v) in d.items() if not v.empty}

    result = []
    for k, v in new_d.items():

        # select dataframe by
        if date in v.to_string():
            df = v.dropna(how='all')
            result.append(df)

    first_shift = result[0]['Frontend']
    second_shift = result[0]['Backend']

    return first_shift, second_shift


def ICClogic():

    # print(df)

    # BRH dont have shift times so maybe just put the hours?

    # ICC


getTableEmail()
# APCClogic()
