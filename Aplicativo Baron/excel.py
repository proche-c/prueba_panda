# OPCION 1 --> OPENPYXL

import openpyxl

#c_path = "archivos/comprobacion.xlsx"
#wb = openpyxl.load_workbook(c_path)
#recibos = {}
#print(wb.sheetnames)

# OPCION 2 --> PANDA

import pandas as pd

df_recibos = pd.ExcelFile("archivos/comprobacion.xlsx")

print(df_recibos.sheet_names)

df = df_recibos.parse('Sheet1')

print(df[0])