import pandas as pd
import openpyxl

outputfilename="Output_Sheet_original.xlsx"

def makeoutputsheetblank(outputsheetname,rowindex,colindex):
    myworkbook=openpyxl.load_workbook(outputfilename)
    worksheet= myworkbook[outputsheetname]
    for  rowindexvalue in range(0,rowindex):
        for index in range(1,colindex):
            worksheet.cell(row=int(rowindexvalue)+4,column=index).value=""
    myworkbook.save(outputfilename)
    myworkbook.close()

