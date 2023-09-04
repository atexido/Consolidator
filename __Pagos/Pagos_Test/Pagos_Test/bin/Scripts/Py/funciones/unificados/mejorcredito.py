import pandas as pd 
import numpy as np

def unificado_mejorcredito(empresa, unificado, i):
    df_unificado = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{unificado}', usecols=['DNI', 'IMPORTE'])
    #Dejo df_filtro1 con los 'Falta Comprobante' y realizar su validación
    #y creo df_filtro2 con los demás para luego concatenar df_filtro1 con df_filtro2
    df_filtro1 = df_concat[df_concat['STATUS']=='Falta Comprobante']
    #Los "Falta Comprobante" son los que van a chocar con el unificado
    df_filtro2 = df_concat[df_concat['STATUS']!='Falta Comprobante']
    df_filtro1 = df_filtro1[['DNI', 'STATUS', 'Operador']]

    lista_dnis = list(df_filtro1['DNI'])
    df_nuevo = pd.DataFrame({})

    for index_dfu in range(len(df_unificado)):
        if df_unificado.iloc[index_dfu]['DNI'] in lista_dnis:
            registro = pd.DataFrame({
                'DNI': [df_unificado.iloc[index_dfu]['DNI']],
                'STATUS': ['Válido por Unificado'],
                'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                'Operador': [i]
                })
            df_nuevo = pd.concat([df_nuevo, registro])
            #Elimino el dni de la lista
            lista_dnis.remove(df_unificado.iloc[index_dfu]['DNI'])

    df_resto = pd.DataFrame({'DNI': lista_dnis})
    df_resto['IMPORTE'] = np.nan
    df_resto['STATUS'] = 'Falta Comprobante'
    df_resto['Operador'] = i

    df_unificado = pd.concat([df_nuevo, df_resto])

    #-------------------------------------------------------#
    # lista_unificado = list(df_unificado['DNI'])
    # lista_df_filtro1 = list(df_filtro1['DNI'])

    # dni_estado = []
    # for dni in lista_df_filtro1:
    #     if dni in lista_unificado:
    #         dni_estado.append('Válido por Unificado')
    #         lista_unificado.remove(dni)
    #     else:
    #         dni_estado.append('Falta Comprobante')

    # df_status_excel = pd.DataFrame({'DNI': lista_df_filtro1, 'STATUS': dni_estado,})
    #-------------------------------------------------------#
    df_concat = pd.concat([df_unificado, df_filtro2])
    df_concat['DNI'] = df_concat['DNI'].apply(int)
    print('Concat luego de unir los que no chocaron contra los nuevos resultados')
    print(df_concat)

    # #✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔