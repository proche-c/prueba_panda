import openpyxl
import pandas as pd
from datetime import datetime

# Defino el path en el que tengo mi fichero excel. Mas tarde incorporare interfaz grafica para importarlo
path = "archivos/comprobacion.xlsx"

# Funcion que decide que comision aplicar segun el año de Primer efecto
def choose_com(year_r, month_r, year_first, month_first):
    if year_r == year_first or (year_r == year_first + 1 and month_r < month_first):
        return 1
    elif (year_r == year_first + 1 and month_r >= month_first) or (year_r == year_first + 2 and month_r < month_first):
        return 2
    elif (year_r == year_first + 2 and month_r >= month_first) or (year_r == year_first + 3 and month_r < month_first):
        return 3
    elif (year_r == year_first + 3 and month_r >= month_first) or (year_r == year_first + 4 and month_r < month_first):
        return 4
    elif (year_r == year_first + 4 and month_r >= month_first) or (year_r == year_first + 5 and month_r < month_first):
        return 5
    else:
        return 6

# Creo un dataframe a partir del fichero excel, aunque voy a optimizarlo un poco
df_recibos = pd.read_excel(path, header=1)

# col establece los campos del dataframe, num_col nos da la dimension
col = df_recibos.columns.values
num_col = len(df_recibos.columns.values)
length = len(df_recibos)

#ahora voy a eliminar las columnas que no necesito y quedarme solo con las que quiero. Para ello, defino una lista col
#que contendra solo las columnas con las que voy a trabajar, las que tienen un campo con nombre
col = []

#defino una funcion que quite las columnas sin titulo y guarde las que si en una lista
def clean_fields(col):
    for i_col in range(num_col):
        if df_recibos.columns.values[i_col][0:7] != 'Unnamed' :
            col.append(i_col)
    return col

# Redefino el nuevo Dataframe eliminando las columnas vacias
col = clean_fields(col)
df_recibos = pd.read_excel(path, header=1, usecols=col)

# Con los datos del DataFrame creado, construyo la lista de diccionarios recibos[]
# Cada elemento de la lista es una fila del Excel
recibos = []

def built_dic(df_recibos):
    length = len(df_recibos)
    for ind in range(length) :
        recibo = {df_recibos.columns.values[0] : df_recibos.values[ind][0],
        df_recibos.columns.values[1] : df_recibos.values[ind][1], 
        df_recibos.columns.values[2] : df_recibos.values[ind][2],
        df_recibos.columns.values[3] : df_recibos.values[ind][3],
        df_recibos.columns.values[4] : df_recibos.values[ind][4],
        df_recibos.columns.values[5] : df_recibos.values[ind][5],
        df_recibos.columns.values[6] : df_recibos.values[ind][6],
        df_recibos.columns.values[7] : df_recibos.values[ind][7],
        df_recibos.columns.values[8] : df_recibos.values[ind][8],
        df_recibos.columns.values[9] : df_recibos.values[ind][9],
        df_recibos.columns.values[10] : df_recibos.values[ind][10],
        df_recibos.columns.values[11] : df_recibos.values[ind][11],
        df_recibos.columns.values[12] : df_recibos.values[ind][12],
        df_recibos.columns.values[13] : df_recibos.values[ind][13],
        df_recibos.columns.values[14] : df_recibos.values[ind][14],
        df_recibos.columns.values[15] : df_recibos.values[ind][15],
        df_recibos.columns.values[16] : df_recibos.values[ind][16],
        df_recibos.columns.values[17] : df_recibos.values[ind][17],
        df_recibos.columns.values[18] : df_recibos.values[ind][18],
        df_recibos.columns.values[19] : df_recibos.values[ind][19],
        df_recibos.columns.values[20] : df_recibos.values[ind][20],
        df_recibos.columns.values[21] : df_recibos.values[ind][21],
        df_recibos.columns.values[22] : df_recibos.values[ind][22],
        'Poliza.Producto.Com6' : 0,
        'Comision Cobrada': 0,
        'Diferencia':0,
        'A deber':0}
        recibos.append(recibo)
        ind = ind + 1
    return recibos

recibos = built_dic(df_recibos)

# Ahora añado la comision del sexto año
def add_Com6(recibos):
    for recibo in recibos:
        if recibo['Alias'][0:5] == 'ZMP(1':
            recibo['Poliza.Producto.Com6'] = 0.14
        elif recibo['Alias'][0:5] == 'ZMP(2':
            recibo['Poliza.Producto.Com6'] = 0.1
        elif recibo['Alias'][0:5] == 'ZMP(3':
            recibo['Poliza.Producto.Com6'] = 0.07
        else:
            recibo['Poliza.Producto.Com6'] = recibo['Poliza.Producto.Com5']
    return recibos

recibos = add_Com6(recibos)

# Creo la funcion imprimir recibos para cuando quiera sacar por pantalla de terminal
def imprimir_recibos(recibos, imprimir):
    print("******************")
    print("******************")
    for recibo in recibos:
        for clave, valor in recibo.items() :
            print(clave, ": ", recibo[clave])
        print("******************")
        print("******************")

#imprimir_recibos(recibos, 0)

# Ahora separo la lista en 2, las que tienen comisiones anomalas y las que no
recibos_n = []
recibos_ca = []

for recibo in recibos:
    if type(recibo['Comision reducida']) == float or recibo['Comision reducida'] == 'Tabla':
        recibos_n.append(recibo)
    else:
        recibos_ca.append(recibo)

# Ahora calculo la diferencia entre comision pactada y cobrada, que dependera en muchos casos de la antiguedad de la poliza
def obtener_desv(recibos, anomalia):
    for recibo in recibos:
        if anomalia == True and recibo['Comision reducida'][0:] != 'Referencia':
            com_a = float(recibo['Comision reducida'][0:-1]) / 100
        else:
            com_a = 1
        com = choose_com(recibo['Fecha efecto'].year, recibo['Fecha efecto'].month, recibo['Poliza.FechaPrimerEfecto'].year, recibo['Poliza.FechaPrimerEfecto'].month)
        recibo['Comision Cobrada'] = recibo['Comisión correduría'] / recibo['Prima neta']
        if com == 1:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com1'] * com_a - recibo['Comisión correduría']
        elif com == 2:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com2'] * com_a - recibo['Comisión correduría']
        elif com == 3:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com3'] * com_a - recibo['Comisión correduría']
        elif com == 4:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com4'] * com_a - recibo['Comisión correduría']
        elif com == 5:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com5'] * com_a - recibo['Comisión correduría']
        else:
            recibo['Diferencia'] = recibo['Prima neta'] * recibo['Poliza.Producto.Com6'] * com_a - recibo['Comisión correduría']
        if recibo['Diferencia'] > 0.03 and recibo['Prima neta'] > 0:
            recibo['A deber'] = recibo['Diferencia']
    return recibos

recibos_n = obtener_desv(recibos_n, False)
recibos_ca = obtener_desv(recibos_ca, True)

# Clasifico los elemnetos de las listas dependiendo de si la comision pagada es correcta
recibos_a_deber = []
recibos_a_deber_ca = []
recibos_correctos = []

#print(type(fecha_muestra['Poliza.FechaPrimerEfecto']))
for recibo in recibos_n:
    if recibo['A deber'] > 0:
        recibos_a_deber.append(recibo)
    else:
        recibos_correctos.append(recibo)

for recibo in recibos_ca:
    if recibo['A deber'] > 0:
        recibos_a_deber_ca.append(recibo)
    else:
        recibos_correctos.append(recibo)

# Convierto las listas en dataframe y exporto a Excel
df_recibos_correctos = pd.DataFrame(recibos_correctos)
df_recibos_a_deber = pd.DataFrame(recibos_a_deber)
df_recibos_a_deber_cr = pd.DataFrame(recibos_a_deber_ca)

df_recibos_correctos['Poliza.FechaPrimerEfecto'] = df_recibos_correctos['Poliza.FechaPrimerEfecto'].astype(str)
df_recibos_correctos['Poliza.FechaEfecto'] = df_recibos_correctos['Poliza.FechaEfecto'].astype(str)
df_recibos_correctos['Fecha efecto'] = df_recibos_correctos['Fecha efecto'].astype(str)

with pd.ExcelWriter("archivos/resultado.xlsx") as writer:
    df_recibos_correctos.to_excel(writer, sheet_name="Correctos")
    df_recibos_a_deber.to_excel(writer, sheet_name="A deber")
    df_recibos_a_deber_cr.to_excel(writer, sheet_name="Com. Reducidas")