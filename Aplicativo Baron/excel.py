import openpyxl
import pandas as pd

df_recibos = pd.ExcelFile("archivos/comprobacion.xlsx")

print(df_recibos.sheet_names)

df1 = df_recibos.parse('Sheet1')

print(df1.values[0])
print(df1.values[0][1])