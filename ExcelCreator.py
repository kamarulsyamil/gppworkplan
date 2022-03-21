import imp
from statistics import mode
import xlsxwriter as xlwrite
from ExcelExtractor import *
import pandas as pd

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

    fname = 'DCCC4'  # df.fName
    date = 'Mar-11'

    # Merge 3 cells.
    # headers
    worksheet.merge_range('B4:C4', 'Factory/Site', merge_format)
    worksheet.merge_range('D4:E4', fname, merge_format)
    worksheet.merge_range('F4:G4', 'Date', merge_format)
    worksheet.write('H4', date, merge_format)

    worksheet.write('B5', 'Line', bold)
    worksheet.write('C5', 'Start Time', bold)
    worksheet.write('D5', 'End Time', bold)
    worksheet.write('E5', 'UPH', bold)
    worksheet.write('F5', 'Start Time', bold)
    worksheet.write('G5', 'End Time', bold)
    worksheet.write('H5', 'UPH', bold)

    workbook.close()


def insertData():

    writer = pd.ExcelWriter('test.xlsx', engine='openpyxl',
                            mode='a', if_sheet_exists='overlay')

    worksheet = writer.sheets['Sheet1']

    CCC4Df().to_excel(writer, sheet_name='Sheet1',
                      startrow=5, startcol=1, header=False, index=False)

    worksheet.set_column(2, 2, None, data_theme)

    writer.save()


# createWorkbook()
insertData()
