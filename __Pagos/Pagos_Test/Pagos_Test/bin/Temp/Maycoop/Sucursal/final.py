import pandas as pd
import datetime
import os

cliente = 'MAYCOOP'
cliente_temp = 'Maycoop'

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



if os.path.isfile(F'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/VALIDOS.xlsx'):

    print('LEYENDO CIERRE')
    # Leemos el archivo VALIDOS y eliminamos la columna OPERADOR

    # Leer los archivos xlsx
    df_cierre = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/CIERRE.xlsx', usecols=['DNI', 'IMPORTE'])
    df_distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/DISTRIBUCION.xlsx', usecols=['DNI', 'OPERADOR'])

    # Comparar las dos columnas DNI
    df_final = pd.merge(df_cierre, df_distribucion, on='DNI')

    # Crear un nuevo archivo xlsx
    df_final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/FINAL_CIERRE.xlsx', index=False)
    



#if os.path.isfile(F'M:/Central de pagos/2023/{mes_actual_espanol_carpetacdp}/Macro Central de Pagos {mes_actual_espanol_mayusc}.xlsm'):
else:
    print('Leyendo CENTRAL DE PAGOS...')
    df_cdp = pd.read_excel(F'M:/Central de pagos/2023/{mes_actual_espanol_carpetacdp}/Macro Central de Pagos {mes_actual_espanol_mayusc}.xlsm', sheet_name='Base', usecols=['DNI', ' Importe Total '])
    df_cdp = df_cdp.rename(columns={' Importe Total ': 'IMPORTE'})

    df_cdp['TOTAL'] = ''
    df_cdp['STATUS'] = ''
    df_cdp.loc[0, 'TOTAL'] = df_cdp['IMPORTE'].sum()
    df_cdp = df_cdp[['DNI', 'STATUS', 'IMPORTE', 'TOTAL']]
    df_cdp['DNI'] = df_cdp['DNI'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))
    #print(df_cdp)


    if os.path.isfile(f'H:/Casos/2023/{mes_actual_espanol_carpeta}/{cliente}/PAGOS'):
        #Leo los pagos de ALTAS->CASOS
        directory = (f'H:/Casos/2023/{mes_actual_espanol_carpeta}/{cliente}/PAGOS')  # Puedes cambiarlo al directorio que necesites

        # Lista de archivos que contienen la palabra 'LOAN'
        archivos = 'NULL'
        archivos = [archivo for archivo in os.listdir(directory) if 'LOAN' in archivo]
        # Inicializar la variable de la sumatoria
        total_importe = 0

        num_filas_final_casos = 0
        # Iterar sobre los archivos y sumar los valores de la columna 'IMPORTECONSOLIDADO'
        for archivo in archivos:
            df_casos = pd.read_excel(os.path.join(directory, archivo))
            total_importe += df_casos['IMPORTECONSOLIDADO'].sum()

            num_filas_final_casos += len(df_casos)
        # Lee el archivo más reciente de la lista usando pandas
        df_final_casos = df_casos['IMPORTECONSOLIDADO']
        #total_casos = df_final_casos.loc[0, 'TOTAL']
        #print(total_importe)
        #print(num_filas_final_casos)


        # Leemos el archivo VALIDOS y eliminamos la columna OPERADOR
        #validos_casos = total_importe
        validos_casos = df_casos.drop(columns=['NUMEROCONVENIO', 'NUMEROOPERACION', 'COMPANIA',	'FECHAPAGO', 'IMPORTECAPITAL', 'IMPORTEHONORARIOS', 'IMPORTECOMISIONCANALPAGO', 'OBSERVACION', 'CANALPAGO', 'NUMERORECIBO'])
        # Nos quedamos únicamente con la columna DNI y eliminamos caracteres no numéricos

        validos_casos['DOCUMENTO'] = validos_casos['DOCUMENTO'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))
        validos_casos['DNI'] = validos_casos['DOCUMENTO']
        validos_casos['TOTAL'] = ''
        validos_casos['STATUS'] = ''
        #validos_casos['OPERADOR'] = ''
        validos_casos.loc[0, 'TOTAL'] = total_importe
        validos_casos = validos_casos.rename(columns={'IMPORTECONSOLIDADO': 'IMPORTE'})
        validos_casos = validos_casos[['DNI', 'STATUS', 'IMPORTE', 'TOTAL']]
        #print(validos_casos)

        # Leemos el archivo DISTRIBUCION y nos quedamos con las columnas DNI y OPERADOR
        distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/DISTRIBUCION.xlsx', usecols=['DNI', 'OPERADOR'])
        distribucion['DNI'] = distribucion['DNI'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

        #validos_casos['OPERADOR'] = distribucion['OPERADOR']

        # Realizamos la comparación de los valores de la columna DNI y nos quedamos con los valores iguales
        final = pd.merge(df_cdp, validos_casos, distribucion, on='DNI')
        final.loc[0, 'TOTAL'] = final['IMPORTE'].sum()
        #print(final)

        # Cambio de posición de la columna TOTAL con la columna OPERADOR
        final = final[["DNI", "STATUS", "IMPORTE", "OPERADOR", "TOTAL"]]

        # Guardamos los datos en un nuevo archivo xlsx llamado FINAL
        final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/FINAL.xlsx', index=False)

    else:
        print('NO EXISTE ARCHIVO EN ALTAS/CASOS')


        # Leemos el archivo DISTRIBUCION y nos quedamos con las columnas DNI y OPERADOR
        distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/DISTRIBUCION.xlsx', usecols=['DNI', 'OPERADOR'])
        distribucion['DNI'] = distribucion['DNI'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

        # Realizamos la comparación de los valores de la columna DNI y nos quedamos con los valores iguales
        final = pd.merge(df_cdp, distribucion, on='DNI')
        #final = pd.concat([final, distribucion[~distribucion['DNI'].isin(final['DNI'])]])
        final.loc[0, 'TOTAL'] = final['IMPORTE'].sum()
        

        # Cambio de posición de la columna TOTAL con la columna OPERADOR
        final = final[["DNI", "STATUS", "IMPORTE", "OPERADOR", "TOTAL"]]

        #final['OPERADOR'].fillna('NO DISTRIBUIDO', inplace=True)

        #print(final)
        # Guardamos los datos en un nuevo archivo xlsx llamado FINAL
        final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{cliente_temp}/Sucursal/FINAL.xlsx', index=False)

