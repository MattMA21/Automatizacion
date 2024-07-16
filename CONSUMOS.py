import pandas as pd
import os

file_path = r'C:\xampp\htdocs\Automatizacion\CONSUMOS2024xlsx'
if not os.path.exists(file_path):
    raise FileNotFoundError(f"El archivo no se encuentra en la ruta especificada: {file_path}")

xls = pd.ExcelFile(file_path)

print("Hojas en el archivo Excel:", xls.sheet_names)

tablas_dinamicas = {}

for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name)
    
    print(f"\nPrimeros registros de la hoja {sheet_name}:")
    print(df.head())
    
 
    pivot_table = pd.pivot_table(df, 
    index=df.columns[0],  
    aggfunc='sum') 
    
    tablas_dinamicas[sheet_name] = pivot_table

output_path = r'C:\xampp\htdocs\Automatizacion\resultado_tablas_CONSUMOS2024.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, pivot_table in tablas_dinamicas.items():
        pivot_table.to_excel(writer, sheet_name=sheet_name)

print(f"Resultados exportados a {output_path}")
