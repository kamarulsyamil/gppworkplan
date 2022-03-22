import imp
from statistics import mode
import xlsxwriter as xlwrite
from ExcelExtractor import *
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Color, PatternFill, Font, GradientFill, Alignment

workbook = xlwrite.Workbook('test.xlsx')


def createWorkbook():
    worksheet = workbook.add_worksheet()

    # worksheet.set_column()

    # Increase the cell size of the merged cells to highlight the formatting.
    # worksheet.set_column('B:D', 12)
    # worksheet.set_row(3, 30)
    # worksheet.set_row(6, 30)
    # worksheet.set_row(7, 30)

    # Create a format to use in the merged range.

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#f4B084'})

    bold = workbook.add_format({'bold': 1,
                                'border': 1})

    global data_theme

    data_theme = workbook.add_format({'border': 1})

    # df = CCC4Df()

    # fname = 'DCCC4'  # df.fName
    # date = 'Mar-11'

    # Merge 3 cells.
    # headers
    worksheet.merge_range('B4:C4', 'Factory/Site', merge_format)
    worksheet.merge_range('D4:E4', fname, merge_format)
    worksheet.merge_range('F4:G4', 'Date', merge_format)
    worksheet.write('H4', '', merge_format)

    worksheet.write('B5', 'Line', bold)
    worksheet.write('C5', 'Start Time', bold)
    worksheet.write('D5', 'End Time', bold)
    worksheet.write('E5', 'UPH', bold)
    worksheet.write('F5', 'Start Time', bold)
    worksheet.write('G5', 'End Time', bold)
    worksheet.write('H5', 'UPH', bold)

    workbook.close()


def insertData():
    wb = Workbook()
    ws = workbook.active

    df, fname, fdate = CCC4Df()

    writer = pd.ExcelWriter('test.xlsx', engine='openpyxl',
                            mode='a', if_sheet_exists='overlay')

    worksheet = writer.sheets['Sheet1']

    CCC4Df().to_excel(writer, sheet_name='Sheet1',
                      startrow=5, startcol=1, header=False, index=False)

    #worksheet.write('H4', fdate)

    worksheet.set_column(2, 2, None, data_theme)

    writer.save()


def openpxWorkbook():

    double = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

    header = ['Factory/Site', '', '', '', 'Date', '', '']

    fName = ['CCC4', 'CCC2', 'CCC6', 'APCC', 'ICC', 'EMFP', 'BRH1']

    subheader = ['Line', 'Start Time', 'End Time',
                 'UPH', 'Start Time', 'End Time', 'UPH']

    wb = Workbook()
    ws = wb.active

    # title of worksheet
    ws.title = "Workplans"

    # create tables

    rows = 7

    for row in ws.iter_rows(min_row=7, min_col=2, max_row=63, max_col=8):
        for cell in row:
            cell.border = double

        if rows == 7 or rows == 15 or rows == 23 or rows == 31 or rows == 39 or rows == 48 or rows == 56:
            i = 0
            for cell in row:
                cell.value = header[i]
                cell.border = double
                cell.fill = PatternFill("solid", fgColor="DDDDDD")
                i += 1

        elif rows == 8 or rows == 16 or rows == 24 or rows == 32 or rows == 40 or rows == 49 or rows == 57:
            i = 0
            for cell in row:
                cell.value = subheader[i]
                i += 1

        rows += 1

    # merge header
    header = 7
    max_row = 63

    for x in range(max_row):
        if header == 7 or header == 15 or header == 23 or header == 31 or header == 39 or header == 48 or header == 56:
            ws.merge_cells(start_row=header, start_column=2,
                           end_row=header, end_column=3)
            ws.merge_cells(start_row=header, start_column=4,
                           end_row=header, end_column=5)
            ws.merge_cells(start_row=header, start_column=6,
                           end_row=header, end_column=7)

        header += 1

    # subheader

    # save xl to explorer
    wb.save('Consolidated Factory Workplan.xlsx')


# createWorkbook()
# insertData()
openpxWorkbook()
