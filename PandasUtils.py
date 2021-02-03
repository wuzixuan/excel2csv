from pandas import pandas as pd
import numpy as np

def open_excel(fileName):
    excelFile = pd.read_excel(fileName, sheet_name=None)
    return excelFile
    #print(excelFile.keys())

def get_sheet_name_list(fileName):
    excelFile = pd.read_excel(fileName, sheet_name=None)
    return excelFile.keys()

def excel_convers_to_csv(fileName, sheetName):
    excelFile:pd.DataFrame = pd.read_excel(fileName, sheet_name=sheetName)
    excelFile.to_csv(fileName+".csv", index=False)
