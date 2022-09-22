import os
import pandas as pd
from tabula.io import read_pdf
import msoffcrypto
import io
from xlsx2csv import Xlsx2csv

temp = io.BytesIO()

def data(file):
    df = read_pdf(file, pages='all')
    list1 = [] 
    for item in df:
        for info in item.values:
            list1.append(info)
    df = pd.DataFrame(list1)
    return df

def data_excel(file, header):
    df = pd.read_excel(file, header=[header])
    return df


files = [os.path.join(root, name)
             for root, dirs, files in os.walk(os.getcwd())
             for name in files
             if name.endswith((".pdf"))]

# for file in files:
#     print(data(file))

x_files = [os.path.join(root, name)
             for root, dirs, files in os.walk(os.getcwd())
             for name in files
             if name.endswith((".xlsx",".xls"))]

# for file in x_files:
#     print(data_excel(file,2))
# # print(data_excel('C:\C-STEP FOLDER\MAX&MIN VOLTAGE\\2018\DECEMBER-2018.xlsx',2))

# for root, dirs, files in os.walk(os.getcwd()):
#     for name in files:
#         if(name.endswith(('.xlsx'))):
#             xl = New-Object -comobject "Excel.Application"
#             # repeat this for every file concerned
#             wb = xl.Workbooks.open(os.path.join(root, name),3)
#             wb.SaveAs(os.path.join(root, 'root', name))
#             wb.Close(False)

# print(x_files[0])

# print(data_excel(x_files[0],2)).

# with open(x_files[0], 'rb') as f:
#     excel = msoffcrypto.OfficeFile(f)
#     excel.load_key('128')
#     excel.decrypt(temp)

# print(data_excel(temp, 2))


Xlsx2csv(x_files[0], outputencoding="utf-8").convert("C:/myfile.csv")

