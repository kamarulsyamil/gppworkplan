from cmath import nan
import os
from numpy import NaN
import pandas as pd
from scipy.fftpack import shift
import win32com.client
import datetime
import textwrap

date = datetime.datetime.today()


def getTableEmail():
    factName = 'APCC'

    #email_dir = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\ICC Shift timings_.msg"
    #email_dir = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\ICC Shift timings_.msg"
    email_dir = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\APCC Work Plan.msg"

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

        # print(df1)

        # end_time = shift_time[:len(shift_time)//2] [shift_time[i:i+chunk_size] for i in range(0, chunk, chunk_size)]
        # print([shift_time[i:i+chunk_size]
        #       for i in range(0, chunk, chunk_size)])
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

    # print(df)

    # BRH dont have shift times so maybe just put the hours?

    # ICC


def ICClogic():

    df1 = []

    df2 = []

    df, factoryname = getTableEmail()

    #shift_time = df.loc[1]['BACK END'].replace(" ", "")

    for i, row in df.iterrows():
        data = str(row['FRONT END']).replace(" ", "")

        if len(data) > 12:
            chunk, chunk_size = len(data), len(data)//2

            split_data = [data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df1.append(split_data)

            front_df = pd.DataFrame(
                df1, columns=["first_shift", "second_shift"])
        else:
            chunk, chunk_size = len(data), len(data)//1

            split_data = [data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df1.append(split_data)

            front_df = pd.DataFrame(df1, columns=["first_shift"])

    # Backend array
    for i, row in df.iterrows():
        back_data = str(row['BACK END']).replace(" ", "")

        if len(back_data) > 12:
            chunk, chunk_size = len(back_data), len(back_data)//2

            split_data = [back_data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df2.append(split_data)

            back_df = pd.DataFrame(
                df2, columns=["first_shift", "second_shift"])
        else:
            chunk, chunk_size = len(back_data), len(back_data)//1

            split_data = [back_data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df2.append(split_data)

            back_df = pd.DataFrame(
                df2, columns=["first_shift"])

    return front_df, back_df


# ICClogic()
# APCClogic()
