from pandas import pandas as pd
import pyarrow
import openpyxl
import xlrd
import tkinter

#获取要导入的数据
#如果只有一个sheet，则sheetName留空即可
def getData(fileName,sheetName):

    if sheetName != None:
        # 读取excel
        df:pd.DataFrame = pd.read_excel(fileName, sheet_name=sheetName)
    else:
        df:pd.DataFrame = pd.read_excel(fileName)
    return df





def saveFile(filePath,df):
    #with open(filePath, 'x', encoding='utf-8') as f:
    #    f.write(file)
    df.apply(str)
    df.to_parquet(filePath + "/example_fp.parquet", engine="pyarrow",index=False)

#
if __name__ == '__main__':
    df:pd.DataFrame  = getData("C:/Users/wuzixuan/Desktop/2021-01工作材料/催收投诉/每周投诉数据汇报（1130-1227）.xlsx","数据")
    #df2parquet = df.to_parquet
    print(df)
    saveFile("C:/Users/wuzixuan/Desktop",df)


