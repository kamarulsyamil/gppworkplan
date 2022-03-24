import imp
from statistics import mode
from tkinter.font import BOLD
from sqlalchemy import column
import xlsxwriter as xlwrite
from ExcelExtractor import *
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import WriteOnlyCell

workbook = xlwrite.Workbook('test.xlsx')


def createWorkbook():

    double = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

    header = ['Factory/Site', '', '', '', 'Date', '', '']

    fName = ['CCC4', 'CCC2', 'CCC6', 'APCC', 'ICC', 'EMFP', 'BRH1']

    CCC4List = ['DT Kitting&Cell', 'DT Backend', 'SV Kitting&Cell K6',
                'SV Kitting&Cell K7', 'SV Backend', 'Storage line', 'CFS']

    CCC2List = ['DT Kitting&Cell', 'DT Backend', 'SV Kitting&Cell',
                'SV Backend', 'K8 line', 'ARB']

    APCCList = ['Desktop', 'HYBRID 1', 'HYBRID 2', 'Server']

    ICCList = ['Line 1', 'Line 2', 'Line 3']

    subheader = ['Line', 'Start Time', 'End Time',
                 'UPH', 'Start Time', 'End Time', 'UPH']

    wb = Workbook()
    ws = wb.active

    # title of worksheet
    ws.title = "Workplans"

    # create legend
    ws['J7'] = 'Legend'
    ws['J7'].font = Font(bold=True)
    ws.merge_cells('J7:K7')
    ws['J7'].border = double
    ws['K7'].border = double

    ws['J8'] = 'First shift'
    ws['J9'] = 'Second shift'
    ws['J8'].border = double
    ws['J9'].border = double

    ws['K8'].fill = PatternFill("solid", fgColor="00FFFFCC")
    ws['K9'].fill = PatternFill("solid", fgColor="00FFFF00")
    ws['K8'].border = double
    ws['K9'].border = double

    # subheader and cell color
    for col in ws.iter_cols(min_col=2, min_row=8, max_col=2, max_row=63):
        for cell in col:
            cell.fill = PatternFill("solid", fgColor="DDDDDD")

    for col in ws.iter_cols(min_col=3, min_row=8, max_col=5, max_row=63):
        for cell in col:
            cell.fill = PatternFill("solid", fgColor="00FFFFCC")

    for col in ws.iter_cols(min_col=6, min_row=8, max_col=8, max_row=63):
        for cell in col:
            cell.fill = PatternFill("solid", fgColor="00FFFF00")

    # add factory lines
        #CCC4 and CCC2
    # merge for night UPH
    ws.merge_cells(start_column=8, start_row=11, end_column=8, end_row=13)

    # merge for day UPH
    ws.merge_cells(start_column=5, start_row=9, end_column=5, end_row=10)
    ws.merge_cells(start_column=5, start_row=11, end_column=5, end_row=13)

    for col in ws.iter_cols(min_col=2, min_row=9, max_col=2, max_row=8 + len(CCC4List)):
        i = 0
        for cell in col:
            cell.value = CCC4List[i]
            i += 1

    for col in ws.iter_cols(min_col=2, min_row=18, max_col=2, max_row=17 + len(CCC2List)):
        i = 0
        for cell in col:
            cell.value = CCC2List[i]
            i += 1

        # APCC
    for col in ws.iter_cols(min_col=2, min_row=34, max_col=2, max_row=33 + len(APCCList)):
        i = 0
        for cell in col:
            cell.value = APCCList[i]
            i += 1

        # ICC
    for col in ws.iter_cols(min_col=2, min_row=42, max_col=2, max_row=41 + len(ICCList)):
        i = 0
        for cell in col:
            cell.value = ICCList[i]
            i += 1

    # create tables
    rows = 7
    findex = 0

    for row in ws.iter_rows(min_row=7, min_col=2, max_row=63, max_col=8):
        for cell in row:
            cell.border = double

        if rows == 7 or rows == 16 or rows == 24 or rows == 32 or rows == 40 or rows == 49 or rows == 57:
            i = 0

            for cell in row:
                cell.value = header[i]
                #cell.border = double
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="00FFCC99")
                i += 1

            # factory name
            ws.cell(row=rows, column=4, value=fName[findex])
            findex += 1

        elif rows == 8 or rows == 17 or rows == 25 or rows == 33 or rows == 41 or rows == 50 or rows == 58:
            i = 0
            for cell in row:
                cell.value = subheader[i]
                cell.font = Font(bold=True)
                i += 1

        rows += 1

    # merge header
    header = 7
    max_row = 63

    for x in range(max_row):
        if header == 7 or header == 16 or header == 24 or header == 32 or header == 40 or header == 49 or header == 57:
            ws.merge_cells(start_row=header, start_column=2,
                           end_row=header, end_column=3)
            ws.merge_cells(start_row=header, start_column=4,
                           end_row=header, end_column=5)
            ws.merge_cells(start_row=header, start_column=6,
                           end_row=header, end_column=7)

        header += 1

    # save xl to explorer
    wb.save('Consolidated Factory Workplan.xlsx')


def insertData(factDf):
    wb = load_workbook('Consolidated Factory Workplan.xlsx')
    ws = wb.active

    # testing
    #df, fname, fdate = CCC4Df()

    df, fname, fdate, fshift = factDf

    rows = dataframe_to_rows(df, index=False, header=False)

    # Night start shift
    if fname == 'CCC4' and fshift == 'start':
        for r_idx, row in enumerate(rows, 9):
            for c_idx, value in enumerate(row, 6):
                ws.cell(row=r_idx, column=c_idx, value=value)

        ws['H7'] = fdate
        print(df)

    # Night end shift
    elif fname == 'CCC4' and fshift == 'end':
        for r_idx, row in enumerate(rows, 9):
            for c_idx, value in enumerate(row, 4):
                ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save('Consolidated Factory Workplan.xlsx')


def insertData2(factDf):
    wb = load_workbook('Consolidated Factory Workplan.xlsx')
    ws = wb.active

    df, fname, fdate, fshift, isNight = factDf

    df_list = df.values.tolist()

    # print(df_list)

    # start shift
    if fname == 'CCC4' and fshift == 'start':

        if isNight:
            # gather data
            ws['H7'] = fdate
            K6_df = df[df['Line'].str.contains("Kitting&Cell K6")]
            K6_start = K6_df.loc[K6_df.first_valid_index(), 'Start Time']

            K7_df = df[df['Line'].str.contains("Kitting&Cell K7")]
            K7_start = K7_df.loc[K7_df.first_valid_index(), 'Start Time']

            SVbackend_df = df[df['Line'].str.contains("SV Backend")]
            SVbackend_start = SVbackend_df.loc[SVbackend_df.first_valid_index(
            ), 'Start Time']

            uph = df.loc[0, 'UPH']

            start_time = [K6_start, K7_start, SVbackend_start, uph]

            # insert data
            ws['H11'] = start_time[3]
            ws['H11'].alignment = Alignment(
                horizontal='center', vertical='center')

            for col in ws.iter_cols(min_col=6, min_row=11, max_col=6, max_row=13):
                i = 0
                for cell in col:
                    cell.value = start_time[i]
                    cell.alignment = Alignment(horizontal='center')

                    i += 1

        else:
            # else is day
            ws['H7'] = fdate

            DTFront_df = df[df['Line'].str.contains("DT Kitting&Cell")]
            DTFront_start = DTFront_df.loc[DTFront_df.first_valid_index(
            ), 'Start Time']

            DTBack_df = df[df['Line'].str.contains("DT Backend")]
            DTBack_start = DTBack_df.loc[DTBack_df.first_valid_index(
            ), 'Start Time']

            K6_df = df[df['Line'].str.contains("Kitting&Cell K6")]
            K6_start = K6_df.loc[K6_df.first_valid_index(), 'Start Time']

            K7_df = df[df['Line'].str.contains("Kitting&Cell K7")]
            K7_start = K7_df.loc[K7_df.first_valid_index(), 'Start Time']

            SVbackend_df = df[df['Line'].str.contains("SV Backend")]
            SVbackend_start = SVbackend_df.loc[SVbackend_df.first_valid_index(
            ), 'Start Time']

            storage_df = df[df['Line'].str.contains("Storage line")]
            storage_start = storage_df.loc[storage_df.first_valid_index(
            ), 'Start Time']

            CFS_df = df[df['Line'].str.contains("CFS")]
            CFS_start = CFS_df.loc[CFS_df.first_valid_index(), 'Start Time']

            uph = [df.loc[0, 'UPH'], df.loc[2, 'UPH']]

            # print(df)

            # store data in array

            start_time = [DTFront_start, DTBack_start, K6_start,
                          K7_start, SVbackend_start, storage_start, CFS_start]

            # insert UPH
            ws['E9'] = uph[0]
            ws['E11'] = uph[1]

            ws['E9'].alignment = Alignment(
                horizontal='center', vertical='center')

            ws['E11'].alignment = Alignment(
                horizontal='center', vertical='center')

            # insert shift time

            for col in ws.iter_cols(min_col=3, min_row=9, max_col=3, max_row=15):
                i = 0
                for cell in col:
                    cell.value = start_time[i]
                    cell.alignment = Alignment(horizontal='center')

                    i += 1

    # Night end shift
    elif fname == 'CCC4' and fshift == 'end':
        if isNight:
            # gather data
            K6_df = df[df['Line'].str.contains("Kitting&Cell K6")]
            K6_end = K6_df.loc[K6_df.first_valid_index(), 'End shift']

            K7_df = df[df['Line'].str.contains("Kitting&Cell K7")]
            K7_end = K7_df.loc[K7_df.first_valid_index(), 'End shift']

            SVbackend_df = df[df['Line'].str.contains("SV Backend")]
            SVbackend_end = SVbackend_df.loc[SVbackend_df.first_valid_index(
            ), 'End shift']

            end_time = [K6_end, K7_end, SVbackend_end]

            # insert data
            for col in ws.iter_cols(min_col=7, min_row=11, max_col=7, max_row=13):
                i = 0
                for cell in col:
                    cell.value = end_time[i]
                    cell.alignment = Alignment(horizontal='center')

                    i += 1

        # else is day
        else:
            # else is day
            ws['H7'] = fdate

            DTFront_df = df[df['Line'].str.contains("DT Kitting&Cell")]
            DTFront_end = DTFront_df.loc[DTFront_df.first_valid_index(
            ), 'End shift']

            DTBack_df = df[df['Line'].str.contains("DT Backend")]
            DTBack_end = DTBack_df.loc[DTBack_df.first_valid_index(
            ), 'End shift']

            K6_df = df[df['Line'].str.contains("Kitting&Cell K6")]
            K6_end = K6_df.loc[K6_df.first_valid_index(), 'End shift']

            K7_df = df[df['Line'].str.contains("Kitting&Cell K7")]
            K7_end = K7_df.loc[K7_df.first_valid_index(), 'End shift']

            SVbackend_df = df[df['Line'].str.contains("SV Backend")]
            SVbackend_end = SVbackend_df.loc[SVbackend_df.first_valid_index(
            ), 'End shift']

            storage_df = df[df['Line'].str.contains("Storage line")]
            storage_end = storage_df.loc[storage_df.first_valid_index(
            ), 'End shift']

            CFS_df = df[df['Line'].str.contains("CFS")]
            CFS_end = CFS_df.loc[CFS_df.first_valid_index(), 'End shift']

            print('im here')

            # print(df)

            # store data in array

            end_time = [DTFront_end, DTBack_end, K6_end,
                        K7_end, SVbackend_end, storage_end, CFS_end]

            # insert shift time

            for col in ws.iter_cols(min_col=4, min_row=9, max_col=4, max_row=15):
                i = 0
                for cell in col:
                    cell.value = end_time[i]
                    cell.alignment = Alignment(horizontal='center')

                    i += 1

    wb.save('Consolidated Factory Workplan.xlsx')


# createWorkbook()
# openpxWorkbook()
# insertData(CCC4Df())
insertData2(day_CCC4Df())
