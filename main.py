from EmailExtractor import APCClogic, ICClogic
import ExcelCreator
import ExcelExtractor
import os.path


def main():
    #f = "C:\\Users\\Yusuf\\Documents\\My Project\\Factory Work Plan\\Production Line Arrangement of 2022.xlsx"
    f = r"sources\Production Line Arrangement of 2022.xlsx"

    # create workbook named Consolidated Factory Workplan
    if not os.path.exists(r"C:\Users\Yusuf_Budiawan\Documents\Factory-Work-Plan-Consolidate\Factory-Work-Plan-Consolidate\Consolidated Factory Workplan.xlsx"):
        print("Creating workbook...")
        ExcelCreator.createWorkbook()

    print("Processing data...")
    # gather dataframes

    CCC4_day_df = ExcelExtractor.day_CCC4(f)[0]
    CCC4_night_df = ExcelExtractor.night_CCC4(f)[0]

    CCC4_day_df2 = ExcelExtractor.day_CCC4(f)[1]
    CCC4_night_df2 = ExcelExtractor.night_CCC4(f)[1]

    CCC2_day_df = ExcelExtractor.day_CCC2(f)[0]
    CCC2_night_df = ExcelExtractor.night_CCC2(f)[0]

    CCC2_day_df2 = ExcelExtractor.day_CCC2(f)[1]
    CCC2_night_df2 = ExcelExtractor.night_CCC2(f)[1]

    # process dataframes
    CCC4_day_df_clean = ExcelExtractor.day_CCC4Df(CCC4_day_df)
    CCC4_night_df_clean = ExcelExtractor.night_CCC4Df(CCC4_night_df)
    CCC4_day_df_clean2 = ExcelExtractor.day_CCC4Df(CCC4_day_df2)
    CCC4_night_df_clean2 = ExcelExtractor.night_CCC4Df(CCC4_night_df2)

    CCC2_day_df_clean = ExcelExtractor.day_CCC2Df(CCC2_day_df)
    CCC2_night_df_clean = ExcelExtractor.night_CCC2Df(CCC2_night_df)
    CCC2_day_df_clean2 = ExcelExtractor.day_CCC2Df(CCC2_day_df2)
    CCC2_night_df_clean2 = ExcelExtractor.night_CCC2Df(CCC2_night_df2)

    # insert data for CCC4
    ExcelCreator.CCC4DataInsert(CCC4_day_df_clean)
    ExcelCreator.CCC4DataInsert(CCC4_night_df_clean)
    ExcelCreator.CCC4DataInsert(CCC4_day_df_clean2)
    ExcelCreator.CCC4DataInsert(CCC4_night_df_clean2)

    # insert data for CCC2
    ExcelCreator.CCC2DataInsert(CCC2_day_df_clean)
    ExcelCreator.CCC2DataInsert(CCC2_night_df_clean)
    ExcelCreator.CCC2DataInsert(CCC2_day_df_clean2)
    ExcelCreator.CCC2DataInsert(CCC2_night_df_clean2)

    # insert data for ICC
    ExcelCreator.ICCDataInsert(ICClogic())

    # insert data for APCC
    ExcelCreator.APCCDataInsert(APCClogic())

    print("Data insertion to workbook successful!")


if __name__ == "__main__":
    main()
