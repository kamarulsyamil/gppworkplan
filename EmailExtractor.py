from cmath import nan
import os
from numpy import NaN
import pandas as pd
import win32com.client
import datetime

date = datetime.datetime.today()
date = '23-Feb'

def getTableEmail():
    factName = 'APCC'

    email_dir = r"C:\Users\Yusuf_Budiawan\Documents\Factory Work Plan\APCC Work Plan.msg"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    msg = outlook.OpenSharedItem(email_dir)

    #dataframe of APCC table from email
    data = pd.read_html(msg.HTMLBody)

    #drop duplicates and NA
    data1 = data[0].dropna(axis=1, how='all', thresh=3)
    #data2 = data1.drop_duplicates(subset='0')

    if factName == 'APCC':
        data1.columns = ['Date','Line','Frontend','Backend']
        
        data2 = data1 #.drop_duplicates()
        # print(data2)

        if not data2[data2['Date'].astype(str).str.contains("Date")].empty:
            data3 = data2.drop(data2.index[range(5)])
            #print(data3.reset_index(drop=True))
            df = data3.reset_index(drop=True)

    #separate tables by date

    return df, factName
    # date_format = date.strftime('%d-%b')

#APCC
def APCClogic():
    df, factName = getTableEmail()

    df['group_no'] = df.isnull().all(axis=1).cumsum()

    #dictionary
    d = {i: df.loc[df.group_no == i, ['Date', 'Line', 'Frontend', 'Backend']]
        for i in range(0, df.group_no.iat[-1])}

    # k v = key value
    new_d = {k:v for (k,v) in d.items() if not v.empty}

    result = []
    for k, v in new_d.items():

        #select dataframe by date
        if date in v.to_string() :
            print('ass')
            result.append(v)

    print(result)

    #print(df)

#BRH dont have shift times so maybe just put the hours?

#ICC

APCClogic()
