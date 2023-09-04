import pandas as pd

# Leer los archivos xlsx
df_cierre = pd.read_excel('CIERRE.xlsx', usecols=['DNI', 'IMPORTE'])
df_distribucion = pd.read_excel('DISTRIBUCION.xlsx', usecols=['DNI', 'OPERADOR'])

# Comparar las dos columnas DNI
df_final = pd.merge(df_cierre, df_distribucion, on='DNI')

# Crear un nuevo archivo xlsx
df_final.to_excel('FINAL_CIERRE.xlsx', index=False)
