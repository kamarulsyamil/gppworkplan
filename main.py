from webbrowser import get
from EmailExtractor import APCClogic, BRHlogic, ICClogic, getTableEmail
import ExcelCreator
import ExcelExtractor
import os.path
import re
import pandas as pd


def main():
    #f = "C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\Production Line Arrangement of 2022.xlsx"
    f = r"sources\Production Line Arrangement of 2022.xlsx"

    # create workbook named Consolidated Factory Workplan
    if not os.path.exists(r"Consolidated Factory Workplan.xlsx"):
        print("Creating workbook...")
        ExcelCreator.createWorkbook()

    print("Data processing started...")

    # gather dataframes

    try:
        print("Gathering data...")

        CCC4_day_df = ExcelExtractor.day_CCC4(f)[0]
        CCC4_night_df = ExcelExtractor.night_CCC4(f)[0]

        CCC4_day_df2 = ExcelExtractor.day_CCC4(f)[1]
        CCC4_night_df2 = ExcelExtractor.night_CCC4(f)[1]

        CCC2_day_df = ExcelExtractor.day_CCC2(f)[0]
        CCC2_night_df = ExcelExtractor.night_CCC2(f)[0]

        CCC2_day_df2 = ExcelExtractor.day_CCC2(f)[1]
        CCC2_night_df2 = ExcelExtractor.night_CCC2(f)[1]

        print("Succesfully gathered data")

    except:
        print("Theres a problem while gathering dataframes")

    # process dataframes

    try:
        print("Processing data...")

        CCC4_day_df_clean = ExcelExtractor.day_CCC4Df(CCC4_day_df)
        CCC4_night_df_clean = ExcelExtractor.night_CCC4Df(CCC4_night_df)
        CCC4_day_df_clean2 = ExcelExtractor.day_CCC4Df(CCC4_day_df2)
        CCC4_night_df_clean2 = ExcelExtractor.night_CCC4Df(CCC4_night_df2)

        CCC2_day_df_clean = ExcelExtractor.day_CCC2Df(CCC2_day_df)
        CCC2_night_df_clean = ExcelExtractor.night_CCC2Df(CCC2_night_df)
        CCC2_day_df_clean2 = ExcelExtractor.day_CCC2Df(CCC2_day_df2)
        CCC2_night_df_clean2 = ExcelExtractor.night_CCC2Df(CCC2_night_df2)

        print("Succesfully processed data")

    except:
        print("Theres a problem while processing dataframes")

    # insert data for CCC4
    try:
        print("inserting data for CCC4...")

        ExcelCreator.CCC4DataInsert(CCC4_day_df_clean)
        ExcelCreator.CCC4DataInsert(CCC4_night_df_clean)
        ExcelCreator.CCC4DataInsert(CCC4_day_df_clean2)
        ExcelCreator.CCC4DataInsert(CCC4_night_df_clean2)

        print("Succesfully inserted CCC4 data")
    except PermissionError as e:
        print("Theres a problem while saving the excel file. Close file if its active.")
        print(str(e))

    except:
        print("Theres a problem while inserting CCC4 data")

    # insert data for CCC2
    try:
        print("inserting data for CCC2...")

        ExcelCreator.CCC2DataInsert(CCC2_day_df_clean)
        ExcelCreator.CCC2DataInsert(CCC2_night_df_clean)
        ExcelCreator.CCC2DataInsert(CCC2_day_df_clean2)
        ExcelCreator.CCC2DataInsert(CCC2_night_df_clean2)

        print("Succesfully inserted CCC2 data")

    except PermissionError as e:
        print("Theres a problem while saving the excel file. Close file if its active.")
        print(str(e))

    except:
        print("Theres a problem while inserting CCC2 data")
    
    #CCC6 data insertion
    try:
        ExcelCreator.CCC6DataInsert()
    except:
        print("Error while reading config file")

    # FROM EMAIL
    # ---------------------------------------------------

    try:
        print("Inserting data for ICC and APCC workplans")

        for i in getTableEmail():

            factName = re.search('(ICC|APCC|BRH)', i.Body).group(0)

            # dataframe of APCC table from email
            data = pd.read_html(i.HTMLBody)
            # print(msg.Body)

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

                    # insert data for APCC
                    ExcelCreator.APCCDataInsert(APCClogic(df, factName))

            elif factName == 'ICC':
                # change header of datarframe
                new_header = data[3].iloc[0]
                df = data[3][1:]
                df.columns = new_header

                # insert data for ICC
                ExcelCreator.ICCDataInsert(ICClogic(df, factName))

            elif factName == 'BRH':
                new_header = data[0].iloc[0]
                df = data[0][1:]
                df.columns = new_header

                df1 = df.drop(['LINE','CAP','Config'], axis = 1)

                new_header = ['LOB', 'HRS1', 'UPH1', 'HRS2', 'UPH2']                
                df1.columns = new_header

                df1 = df1.fillna(0)
                
                ExcelCreator.BRH1DataInsert(BRHlogic(df1, factName))

        print("Done!")

    except PermissionError as e:
        print("Theres a problem while saving the excel file. Close file if its active.")
        print(str(e))

    except:
        print("Error while processing data from e-mail")


if __name__ == "__main__":
    main()
