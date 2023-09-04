import pandas as pd
import datetime

# Leer el archivo viernes y abrir la hoja del mes concurrente
mes_actual = datetime.datetime.now().strftime('%B')
df_viernes = pd.read_excel('M:/PAGOS/MONITOR/Financieras/2023_financieras.xlsx', sheet_name=mes_actual)
cant_sem = len(df_viernes.columns)

print(cant_sem)

if cant_sem == 4:
    import base4
if cant_sem == 5:
    import base5

print('FINALIZADO')