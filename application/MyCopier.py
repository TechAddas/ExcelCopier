import pandas as pd
import mywritter as mw

def createDfFromExcel(input_file,filter_file,outPutColumnSheetName,filterColumnName,tobeSaved, OutPutFileName,inPutColumnSheetName=None):
    input = pd.read_excel(input_file.get(),sheet_name=inPutColumnSheetName)
    output = pd.read_excel(r"ColumnNames.xlsx",sheet_name=outPutColumnSheetName)
    site_list = getFilterColumnValueList(filter_file)
    output_df = pd.DataFrame(input[input[filterColumnName].isin(site_list)], \
         columns=  output["ColumnNames"].values.tolist()).reset_index(drop=True).sort_values(
             by=filterColumnName,
             ascending=False,
             kind="mergesort")
    
    if tobeSaved:
        saveToExcel(output_df,OutPutFileName)
    else:
        return output_df


def getFilterColumnValueList(filter_file_path):
    site_list = pd.read_excel (filter_file_path.get())
    return site_list["Site_List"].values.tolist()

def saveToExcel(data,outFileName):
    mw.WriteTOSheet("Output_Sheet_original.xlsx", outFileName,data)
    # data.to_excel(outFileName+".xlsx",index=False)