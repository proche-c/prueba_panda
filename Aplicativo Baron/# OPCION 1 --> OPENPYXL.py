# OPCION 1 --> OPENPYXL

import openpyxl
import pandas as pd

path = "archivos/comprobacion.xlsx"

df_recibos = pd.read_excel(path)
length = len(df_recibos)
recibos = []
ind = 1
for ind in range(length - 1) :
    recibo = {'df_recibos'}

#print(df_recibos.info)