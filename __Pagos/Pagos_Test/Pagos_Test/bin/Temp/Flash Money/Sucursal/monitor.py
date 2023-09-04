import pandas as pd
import datetime
import os

cliente_cdpxlsx = 'FLASH'
cliente = 'FLASHMONEY'
cliente_temp = 'Flash Money'

mes_actual = datetime.datetime.now().strftime('%B')

meses = {'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo', 'April': 'Abril',
         'May': 'Mayo', 'June': 'Junio', 'July': 'Julio', 'August': 'Agosto',
         'September': 'Septiembre', 'October': 'Octubre', 'November': 'Noviembre',
         'December': 'Diciembre'}

mes_actual_espanol = meses[mes_actual]

if mes_actual_espanol == 'Enero':
    mes_actual_espanol_carpeta = '01 Enero'
    mes_actual_espanol_carpetacdp = '1. Enero'
    mes_actual_espanol_mayusc = 'ENERO'

if mes_actual_espanol == 'Febrero':
    mes_actual_espanol_carpeta = '02 Febrero'
    mes_actual_espanol_carpetacdp = '2. Febrero'
    mes_actual_espanol_mayusc = 'FEBRERO'

if mes_actual_espanol == 'Marzo':
    mes_actual_espanol_carpeta = '03 Marzo'
    mes_actual_espanol_carpetacdp = '3. Marzo'
    mes_actual_espanol_mayusc = 'MARZO'

if mes_actual_espanol == 'Abril':
    mes_actual_espanol_carpeta = '04 Abril'
    mes_actual_espanol_carpetacdp = '4. Abril'
    mes_actual_espanol_mayusc = 'ABRIL'

if mes_actual_espanol == 'Mayo':
    mes_actual_espanol_carpeta = '05 Mayo'
    mes_actual_espanol_carpetcadp = '5. Mayo'
    mes_actual_espanol_mayusc = 'MAYO'

if mes_actual_espanol == 'Junio':
    mes_actual_espanol_carpeta = '06 Junio'
    mes_actual_espanol_carpetacdp = '6. Junio'
    mes_actual_espanol_mayusc = 'JUNIO'

if mes_actual_espanol == 'Julio':
    mes_actual_espanol_carpeta = '07 Julio'
    mes_actual_espanol_carpetacdp = '7. Julio'
    mes_actual_espanol_mayusc = 'JULIO'

if mes_actual_espanol == 'Agosto':
    mes_actual_espanol_carpeta = '08 Agosto'
    mes_actual_espanol_carpetacdp = '8. Agosto'
    mes_actual_espanol_mayusc = 'AGOSTO'

if mes_actual_espanol == 'Septiembre':
    mes_actual_espanol_carpeta = '09 Septiembre'
    mes_actual_espanol_carpetacdp = '9. Septiembre'
    mes_actual_espanol_mayusc = 'SEPTIMEBRE'

if mes_actual_espanol == 'Octubre':
    mes_actual_espanol_carpeta = '10 Octubre'
    mes_actual_espanol_carpetacdp = '10. Octubre'
    mes_actual_espanol_mayusc = 'OCTUBRE'

if mes_actual_espanol == 'Noviembre':
    mes_actual_espanol_carpeta = '11 Noviembre'
    mes_actual_espanol_carpetacdp = '11. Noviembre'
    mes_actual_espanol_mayusc = 'NOVIEMBRE'

if mes_actual_espanol == 'Diciembre':
    mes_actual_espanol_carpeta = '12 Diciembre'
    mes_actual_espanol_carpetacdp = '12. Diciembre'
    mes_actual_espanol_mayusc = 'DICIEMBRE'

# Leer los archivos CENTRAL y ALTAS
if os.path.isfile(F'M:/Central de pagos/2023/{mes_actual_espanol_carpetacdp}/Macro Central de Pagos {mes_actual_espanol_mayusc}.xlsm'):
    print('Leyendo CENTRAL DE PAGOS')
    central = pd.read_excel(F'M:/Central de pagos/2023/{mes_actual_espanol_carpetacdp}/Macro Central de Pagos {mes_actual_espanol_mayusc}.xlsm', usecols=['DNI', ' Importe Total ', 'Cliente'])
    central = central.rename(columns={' Importe Total ': 'IMPORTE'})
    central = central[central['Cliente'].str.contains(cliente_cdpxlsx)]
    
    if os.path.isfile(F'H:/Casos/2023/{mes_actual_espanol_carpeta}/{cliente}/PAGOS/{cliente} PAGOS ACUMULADOS.xlsx'):
        print('Leyendo PAGOS de ALTAS/CASOS')
        altas = pd.read_excel(f'H:/Casos/2023/{mes_actual_espanol_carpeta}/{cliente}/PAGOS/{cliente} PAGOS ACUMULADOS.xlsx', usecols=['DOCUMENTO', 'IMPORTECONSOLIDADO'])
        altas = altas.rename(columns={'DOCUMENTO': 'DNI'})
        altas = altas.rename(columns={'IMPORTECONSOLIDADO': 'IMPORTE'})

        # Combinar los dos DataFrames usando outer join en la columna DNI
        central = central.drop(columns="Cliente")
        df_merge = pd.merge(central, altas, on='DNI', how='outer', suffixes=('_CENTRAL', '_ALTAS'))
        #print(df_merge)
        # Seleccionar solo las filas donde falta un valor en la columna IMPORTE_CENTRAL o IMPORTE_ALTAS
        dni_diff = df_merge[(df_merge['IMPORTE_CENTRAL'].isna()) | (df_merge['IMPORTE_ALTAS'].isna())]['DNI']
        #print(dni_diff)
        # Crear un DataFrame FINAL_2 solo con la columna DNI y la columna IMPORTE de CENTRAL o ALTAS según corresponda
        df_final = pd.DataFrame(columns=['DNI', 'IMPORTE'])
        df_final['TOTAL'] = ''
        df_final['DNI'] = dni_diff
        df_final.loc[df_final['DNI'].isin(central['DNI']), 'IMPORTE'] = central.loc[central['DNI'].isin(df_final['DNI']), 'IMPORTE']
        df_final.loc[df_final['DNI'].isin(altas['DNI']), 'IMPORTE'] = altas.loc[altas['DNI'].isin(df_final['DNI']), 'IMPORTE']

        # Leer el archivo DISTRIBUCION y seleccionar solo la columna DNI
        df_distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/DISTRIBUCION.xlsx', usecols=['DNI'])
        #df_distribucion['DNI'] = df_distribucion['DNI'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

        # Eliminar las filas de FINAL_2 que no se encuentran en DISTRIBUCION
        df_final = df_final[df_final['DNI'].isin(df_distribucion['DNI'])]

        # Seleccionar solo las columnas de interés
        df_final = df_final[['DNI', 'IMPORTE', 'TOTAL']]

        # Eliminar las filas que contienen valores nulos en alguna de las dos columnas
        df_final = df_final.dropna(subset=['DNI', 'IMPORTE'])
        df_final = df_final[['DNI', 'IMPORTE', 'TOTAL']]

        df_final.loc[0, 'TOTAL'] = df_final['IMPORTE'].sum()

        print(df_final)
        # Guardar el DataFrame FINAL_2 en un archivo xlsx
        df_final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx', index=False)

        # Leer el archivo xlsx
        df_final2 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx')

        # Calcular la suma de la columna IMPORTE
        total_final2 = df_final2['IMPORTE'].sum()

        # Crear una nueva columna llamada TOTAL y asignar el valor de la suma
        df_final2.loc[0, 'TOTAL'] = total_final2

        # Guardar el archivo con la columna TOTAL agregada
        df_final2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx', index=False)
        
    else:
        print('NO EXISTE ARCHIVO EN ALTAS CASOS')
        central = central.drop(columns="Cliente")
        # Crear un DataFrame FINAL_2 solo con la columna DNI y la columna IMPORTE de CENTRAL o ALTAS según corresponda
        df_final = pd.DataFrame(columns=['DNI', 'IMPORTE', 'TOTAL'])
        df_final['TOTAL'] = ''
        df_final['DNI'] = central['DNI']
        df_final.loc[df_final['DNI'].isin(central['DNI']), 'IMPORTE'] = central.loc[central['DNI'].isin(df_final['DNI']), 'IMPORTE']

        # Leer el archivo DISTRIBUCION y seleccionar solo la columna DNI
        df_distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/DISTRIBUCION.xlsx', usecols=['DNI'])
        #df_distribucion['DNI'] = df_distribucion['DNI'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

        # Eliminar las filas de FINAL_2 que no se encuentran en DISTRIBUCION
        df_final = df_final[df_final['DNI'].isin(df_distribucion['DNI'])]

        # Seleccionar solo las columnas de interés
        df_final = df_final[['DNI', 'IMPORTE', 'TOTAL']]

        # Eliminar las filas que contienen valores nulos en alguna de las dos columnas
        df_final = df_final.dropna(subset=['DNI', 'IMPORTE'])
        df_final = df_final[['DNI', 'IMPORTE', 'TOTAL']]

        df_final.loc[0, 'TOTAL'] = df_final['IMPORTE'].sum()

        print(df_final)
        # Guardar el DataFrame FINAL_2 en un archivo xlsx
        df_final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx', index=False)

        # Leer el archivo xlsx
        df_final2 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx')

        # Calcular la suma de la columna IMPORTE
        total_final2 = df_final2['IMPORTE'].sum()

        # Crear una nueva columna llamada TOTAL y asignar el valor de la suma
        df_final2.loc[0, 'TOTAL'] = total_final2

        # Guardar el archivo con la columna TOTAL agregada
        df_final2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/MONITOR.xlsx', index=False)
        
else:
    print('NO EXISTE MACRO CENTRAL DE PAGOS')