import pandas as pd
import openpyxl

path = "archivos/58.xml"

with open(path, "r") as f:
    df = pd.read_xml(f.read())

print(df.info)
print(df.shape)
print(df.columns.values)