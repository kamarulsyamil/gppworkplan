from EmailExtractor import APCClogic, ICClogic
import ExcelCreator
import ExcelExtractor
import os.path


def main():
    #f = "C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\Production Line Arrangement of 2022.xlsx"
    f = r"sources\Production Line Arrangement of 2022.xlsx"

    # create workbook named Consolidated Factory Workplan
    if not os.path.exists("C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\ExcelExtractor\\Consolidated Factory Workplan.xlsx"):
        print("Creating workbook...")
        ExcelCreator.createWorkbook()

    print("Processing data...")
    # gather dataframe
    CCC4_day_df = ExcelExtractor.day_CCC4(f)
    CCC4_night_df = ExcelExtractor.night_CCC4(f)

    CCC2_day_df = ExcelExtractor.day_CCC2(f)
    CCC2_night_df = ExcelExtractor.night_CCC2(f)

    # process dataframe
    CCC4_day_df_clean = ExcelExtractor.day_CCC4Df(CCC4_day_df)
    CCC4_night_df_clean = ExcelExtractor.night_CCC4Df(CCC4_night_df)

    CCC2_day_df_clean = ExcelExtractor.day_CCC2Df(CCC2_day_df)
    CCC2_night_df_clean = ExcelExtractor.night_CCC2Df(CCC2_night_df)

    # insert data for CCC4
    ExcelCreator.CCC4DataInsert(CCC4_day_df_clean)
    ExcelCreator.CCC4DataInsert(CCC4_night_df_clean)

    # insert data for CCC2
    ExcelCreator.CCC2DataInsert(CCC2_day_df_clean)
    ExcelCreator.CCC2DataInsert(CCC2_night_df_clean)

    print("Data insertion to workbook successful!")

    # insert data for APCC
    ExcelCreator.APCCDataInsert(APCClogic())

    # insert data for ICCC
    # ExcelCreator.ICCDataInsert(ICClogic())


if __name__ == "__main__":
    main()
