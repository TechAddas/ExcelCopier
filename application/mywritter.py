import pandas as pd
import openpyxl

def WriteTOSheet(filename, sheetname, data):
    writer = pd.ExcelWriter(filename, engine='openpyxl',mode='a',if_sheet_exists="replace")
    data.to_excel(writer, sheet_name=sheetname, index=False)
    writer.save()

def removeSheet(FileName, sheetName):
    workbook=openpyxl.load_workbook(FileName)
    sheets=workbook.get_sheet_names()
    if sheetName in sheets:
        std=workbook.get_sheet_by_name(sheetName)
        workbook.remove_sheet(std)
    workbook.save(FileName)