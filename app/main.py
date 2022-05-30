from extractor.EmailExtractor import APCClogic, BRHlogic, EMFPlogic, ICClogic, getTableEmail
from extractor.ExcelExtractor_copy import CCC2Night, CCC4Day, CCC2Day, CCC4Night
import creator.ExcelCreator as ExcelCreator
import os.path
import re
import pandas as pd
import os
import win32com.client as win32
import json
from openpyxl import load_workbook
import datetime
from openpyxl.styles import Alignment
from time import gmtime, strftime


def main():

    config_path = r"app\configuration\tool_config.json"

    # read config file
    with open(config_path) as config_file:
        config = json.load(config_file)
        file_dir = config['file_dir']
        sharepoint = config["sharepoint"]

    #f = "C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\Production Line Arrangement of 2022.xlsx"
    #f = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\Production Line Arrangement of 2022.xlsx"

    # r'\\w1039fnf93.dhcp.apac.dell.com\PlannerDoc\SHIFT ARRANGEMENT\Production Line Arrangement of 2022.xlsx' #file_dir['CCC2/4']

    f = r'sources\Production Line Arrangement of 2022.xlsx'

    # create workbook named Consolidated Factory Workplan
    if not os.path.exists(file_dir['main_excel']):
        print("Creating workbook...")
        ExcelCreator.createWorkbook(file_dir['main_excel'])

    print("Data processing started...")

    try:
        print("inserting data for CCC4...")

        ExcelCreator.CCC4DataInsert(CCC4Day(), file_dir['main_excel'])
        ExcelCreator.CCC2DataInsert(CCC2Day(), file_dir['main_excel'])

        ExcelCreator.CCC2DataInsert(CCC2Night(), file_dir['main_excel'])
        ExcelCreator.CCC4DataInsert(CCC4Night(), file_dir['main_excel'])

    except PermissionError as e:
        print("Theres a problem while saving the excel file. Close file if its active.")
        print(str(e))

    except Exception as e:
        print("Theres a problem while inserting CCC2 data")
        print("Error: ", str(e))

    # CCC6 data insertion
    try:
        ExcelCreator.CCC6DataInsert(file_dir['main_excel'])
    except:
        print("Error while reading config file")

    # FROM EMAIL
    # ---------------------------------------------------

    #try:
    print("Inserting data for ICC and APCC workplans")

    for i in getTableEmail():

        factName = re.search('(Tom Shift|Work Plan|Commit Produção)', i.Body) #perlu diubah Tom shift = ICC, Work Plan = APCC, 

        # EMFP OT
        Ot_EMFP = re.search('(EMFP Overtime)', i.Subject)

        # dataframe of table from email
        #data = pd.read_html(i.HTMLBody)

        # print(i.Body)

        #data1 = data1.drop_duplicates(subset='0')

        if factName != None:
            if factName.group(0) == 'Work Plan':
                print("Found APCC workplan")

                # drop duplicates and NA
                data1 = pd.read_html(i.HTMLBody)[0].dropna(
                    axis=1, how='all', thresh=3)


                data1.columns = ['Date', 'Line', 'Frontend', 'Backend']

                if not data1[data1['Date'].astype(str).str.contains("Date")].empty:
                    data3 = data1.drop(data1.index[range(5)])
                    #print(data3.reset_index(drop=True))
                    df = data3.reset_index(drop=True)

                    # insert data for APCC
                    ExcelCreator.APCCDataInsert(
                        APCClogic(df, factName), file_dir['main_excel'])
                    print("Successfully inserted data for APCC.")

            elif factName.group(0) == 'Tom Shift':
                print("Found ICC workplan")

                # change header of datarframe
                new_header = pd.read_html(i.HTMLBody)[4].iloc[0] #either 3 or 4
                df = pd.read_html(i.HTMLBody)[4][1:] #either 3 or 4
                df.columns = new_header

                # insert data for ICC
                ExcelCreator.ICCDataInsert(
                    ICClogic(df, factName), file_dir['main_excel'])

            elif factName.group(0) == 'Commit Produção':
                print("Found BRH workplan")
                new_header = pd.read_html(i.HTMLBody)[0].iloc[0]
                df = pd.read_html(i.HTMLBody)[0][1:]
                df.columns = new_header

                df1 = df.drop(['LINE', 'CAP', 'Config'], axis=1)

                new_header = ['LOB', 'HRS1', 'UPH1', 'HRS2', 'UPH2']
                df1.columns = new_header

                df1 = df1.fillna(0)

                ExcelCreator.BRH1DataInsert(
                    BRHlogic(df1, factName), file_dir['main_excel'])

        elif Ot_EMFP != None:
            print("EMFP OT email found")
            ExcelCreator.OTDataInsert(
                'EMFP', i.Body, file_dir['main_excel'])

    print("Done!")

    # except PermissionError as e:
    #     print("Theres a problem while saving the excel file. Close file if its active.")
    #     print(str(e))

    # except Exception as e:
    #     print("Error while processing data from e-mail")
    #     print(str(e))

    # EMFP insertion
    try:
        print("Inserting data for EMFP...")
        ExcelCreator.EMFPDataInsert(
            EMFPlogic(file_dir["EMFP"]), file_dir['main_excel'])
    except Exception as e:
        print("failed to insert EMFP data")
        print(str(e))

    # Updated time
    wb = load_workbook(file_dir['main_excel'])
    ws = wb.active

    ws['F5'] = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    ws['F5'].alignment = Alignment(horizontal='left')
    ws['I5'] = strftime("%z", gmtime())

    wb.save(file_dir['main_excel'])


if __name__ == "__main__":
    main()
