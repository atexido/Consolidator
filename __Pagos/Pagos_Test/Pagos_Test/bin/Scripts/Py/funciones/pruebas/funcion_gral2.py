import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np
import os
import re
import msvcrt

indice = 0
def funcion(empresa):
    os.system('cls')
    os.remove(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_sin_gestion.xlsx')
    os.remove(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx')
    formatos_validos = ['jpg', 'JPG', 'jpeg', 'JPEG', 'pdf', 'jfif', 'PDF', 'html', 'png', 'PNG']

    lista_operadores = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/')
    lista_operadores = list(map(lambda x: x.title(), lista_operadores)) #Mayúscula las primeras letras
    
    try:
        lista_operadores.remove('Thumbs.Db')
    except:
        pass

    lista_valida_operadores = []
    lista_invalida_archivos = []
    df_dni_inicial = pd.DataFrame({'DNI': []})
    df_comprobantes = pd.DataFrame({'DNI': []})

    for index, i in enumerate(lista_operadores):
        lista_archivos = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}')
        lista_archivos = list(map(lambda x: x.title(), lista_archivos)) #2022, Comprobantes, Operador.xlsx
        lista_archivos = list(map(lambda x: x.split('.')[0], lista_archivos)) #Quito el .xlsx
        
        if i in lista_archivos:
            print(f'{i} √')
            lista_valida_operadores.append(i)

            cantidad_excels = 0
            if '~$' not in lista_archivos: #Evita los archivos abierto
                cantidad_excels += 1
                try:
                    df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', usecols=['DNI', 'IMPORTE'])
                except:
                    df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', skiprows=1, usecols=['DNI', 'IMPORTE'])
                df = df.dropna()
                df['Operador'] = f'{i}'
                df_dni_inicial = pd.concat([df_dni_inicial, df])

            if cantidad_excels != 1:
                print(f'Cantidad de archivos Excel: {cantidad_excels}')

            #-------------------- Comprobantes --------------------#
            try:
                lista_comprobantes = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/Comprobantes')
            except:
                print(f'Falta carpeta Comprobantes de: {i}')

            #Filtro por los formatos de interés
            lista_validos = []
            lista_invalidos = []                                                                                                         
            for comprobante in lista_comprobantes:
                extension = comprobante.split('.')[-1]
                if extension in formatos_validos:
                    lista_validos.append(comprobante)
                else:
                    lista_invalidos.append(comprobante)
            
            try:
                lista_invalidos.remove('Thumbs.db')
            except:
                pass 

            if len(lista_invalidos) > 0:
                print(f'Archivos no válidos: {lista_invalidos}')

            #Tomo los valores numéricos
            patron = '[0-9]+'
            lista_patron = []
            for valido in lista_validos:
                valido = re.findall(patron, valido)
                lista_patron.append(valido)

            #Primer valor encontrado (DNI)
            lista_2 = []
            for patron in lista_patron:
                if len(patron)>0:
                    lista_2.append(patron[0])
            

            df_2 = pd.DataFrame({
                'Operador': f'{i}',
                'DNI': lista_2
            })

            #Data de comprobantes. Col DNI y Col del operador
            df_comprobantes = pd.concat([df_comprobantes, df_2]) 
            print('-' * 70)

        else:
            lista_invalida_archivos.append(i)
            
    if len(lista_invalida_archivos) > 0:
        print(f'Lista de archivos descartados: {lista_invalida_archivos}')

    lista_dfs = []

#Iteración sobre los operadores
# ---------------------------------------------------------------------- #
    df_contra_distribucion = []
    for i in lista_valida_operadores:
        #Dni que figuran en el excel
        df_dni2 = df_dni_inicial[df_dni_inicial['Operador'].str.contains(f'{i}')]
        patron = '[0-9]+'
        def validez(i):
            i = str(i)
            valido = re.findall(patron, i)
            valido = [i for i in valido if i != '0']
            return int(''.join(valido))

        df_dni2['DNI'] = df_dni2['DNI'].apply(validez)
        
        def normalizador(numero):
            numero = str(numero)
            if ',' in numero or '.' in numero:
                numero = numero.replace('$', '')
                numero = numero.replace(',', '')
                numero = numero.split('.')[0]
                return int(numero)
            else:
                try:
                    return int(numero)
                except:
                    return 0

        try:
            df_dni2['IMPORTE'] = df_dni2['IMPORTE'].apply(normalizador)
        except:
            print(f'Fallo en la columna DNI de {i}')

        #Comprobantes
        print(f'-------- {i} --------')
        df_comprobantes2 = df_comprobantes[df_comprobantes['Operador'].str.contains(f'{i}')]
        df_comprobantes2['DNI'] = df_comprobantes2['DNI'].apply(validez)
        
        lista_df_dni2 = list(df_dni2['DNI'])
        lista_df_comprobantes2 = list(df_comprobantes2['DNI'])

        dni_estado = []
        for dni in lista_df_dni2:
            if dni in lista_df_comprobantes2:
                dni_estado.append('Válido')
                lista_df_comprobantes2.remove(dni)
            else:
                dni_estado.append('Falta Comprobante')
    
        df_status_excel = pd.DataFrame({'DNI': lista_df_dni2,
                                        'STATUS': dni_estado,
                                        })

        df_status_excel['Operador'] = f'{i}'

        lista_df_dni2 = list(df_dni2['DNI'])
        lista_df_comprobantes2 = list(df_comprobantes2['DNI'])

        comp_estado = []
        for comprobante in lista_df_comprobantes2:
            if comprobante in lista_df_dni2:
                comp_estado.append(np.nan)
                lista_df_dni2.remove(comprobante)
            else:
                comp_estado.append('Falta Excel')

        df_status_comp = pd.DataFrame({'DNI': lista_df_comprobantes2,
                                       'STATUS': comp_estado,
                                      })

        df_status_comp['Operador'] = f'{i}'

        df_concat = pd.concat([df_status_excel, df_status_comp])
        df_concat.dropna(inplace=True)

        try:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', usecols=['DNI', 'IMPORTE'])
        except:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', skiprows=1, usecols=['DNI', 'IMPORTE'])

        df_concat['STATUS'] = df_concat['STATUS'].apply(str)
        df_validos = df_concat[df_concat['STATUS'] == 'Válido']
        df_no_validos = df_concat[df_concat['STATUS'] != 'Válido']

        excel3 = df[['DNI', 'IMPORTE']]
        excel3.dropna(inplace=True)

        excel3['DNI'] = excel3['DNI'].apply(validez)
        excel3['IMPORTE'] = excel3['IMPORTE'].apply(normalizador)

        lista_validos = list(df_validos['DNI'])        

        df_nuevo = pd.DataFrame({
            'DNI': [],
            'IMPORTE': []
        })

        for indice_df in range(len(excel3)):
            if excel3.iloc[indice_df]['DNI'] in lista_validos:
                registro = pd.DataFrame({
                    'DNI': [excel3.iloc[indice_df]['DNI']],
                    'IMPORTE': [excel3.iloc[indice_df]['IMPORTE']],
                    })
                df_nuevo = pd.concat([df_nuevo, registro])
                #Elimino el dni de la lista
                lista_validos.remove(excel3.iloc[indice_df]['DNI'])
        df_nuevo = df_nuevo.astype(int)

        print('resultante de los comprobantes contra el excel')
        df_nuevo['STATUS'] = 'Válido'
        df_nuevo['Operador'] = i
        print(df_nuevo)
    
        df_concat = pd.concat([df_nuevo, df_no_validos])
        print('Primer Concat. Los válidos con su importe + No válidos sin importe')
        print(df_concat)
        print('-' * 100)

        #Hasta acá están los VALIDOS con su importe
        #y en nan los FALTA COMPROBANTE y FALTA EXCEl

        if empresa == 'MejorCredito':
            lista_unificados = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal')
            lista_unificados = [i.lower() for i in lista_unificados] #Todo a minúsculas

            #--------------- Archivos ---------------#
            #El archivo debe tener la palabra "unificado" o "cierre" según corresponda
            unificado = [i for i in lista_unificados if 'unificado' in i]
            unificado = unificado[0]
            cierre = [i for i in lista_unificados if 'cierre' in i]
            cierre = cierre[0]
            #----------------------------------------#

            if cierre == '': 
                print('NO HAY CIERRE')           
                df_unificado = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{unificado}', usecols=['DNI', 'IMPORTE'])
                #Dejo df_filtro1 con los 'Falta Comprobante' y realizar su validación
                #y creo df_filtro2 con los demás para luego concatenar df_filtro1 con df_filtro2
                df_filtro1 = df_concat[df_concat['STATUS']=='Falta Comprobante']
                #Los "Falta Comprobante" son los que van a chocar con el unificado
                df_filtro2 = df_concat[df_concat['STATUS']!='Falta Comprobante']
                df_filtro1 = df_filtro1[['DNI', 'STATUS', 'Operador']]

                #---#
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
        
                df_resto = pd.DataFrame({
                    'DNI': lista_dnis
                })
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

                # print('Estado de los "Falta Comprobante luego del choque')
                # print(df_status_excel)

                df_concat = pd.concat([df_unificado, df_filtro2])
                df_concat['DNI'] = df_concat['DNI'].apply(int)
                print('Concat luego de unir los que no chocaron contra los nuevos resultados')
                print(df_concat)

                #✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔
                #Cruce con Cierre
                
                #Tengo que separar por los "Falta Comprobante" para chocar
                #con el Cierre
                #Y por otro lado los demás a ser concatenados luego del resultado
                #del choque con el cierre

                df_concat['STATUS'] = df_concat['STATUS'].astype(str)
                #Separo por válidos y no válidos
                searchfor = ['Válido', 'Válido por Unificado']
                df_validos = df_concat[df_concat['STATUS'].str.contains('|'.join(searchfor))]
                df_no_validos = df_concat[~df_concat['STATUS'].str.contains('|'.join(searchfor))]

                print('Cruce con cierre:')
                print('Válidos:')
                print(df_validos) 
                print('No válidos:')
                print(df_no_validos)

            else:
                print('HAY CIERRE')
                try:
                    #-------------------- Validacion de los no_validos con el cierre --------------------#
                    df_cierre1 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{cierre}', sheet_name='COBRANZA MEJOR CREDITO', usecols=['doc.', 'fecha', 'est'])
                    df_cierre2 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{cierre}', sheet_name='RENDICION MAS COBRANZAS', usecols=['DNI', 'Fecha de Pago', ' Importe Total '])
                    df_cierre1 = df_cierre1.iloc[:-1] #Elimino el total calculado a mano
                    df_cierre2 = df_cierre2.iloc[:-1] #Elimino el total calculado a mano

                    df_cierre1.rename({'doc.': 'DNI', 'est': 'IMPORTE', 'fecha': 'Fecha de Pago'}, axis=1, inplace=True)
                    df_cierre2.rename({' Importe Total ': 'IMPORTE'}, axis=1, inplace=True)
                    df_cierre1['DF'] = 1 #Distinguir de dónde viene
                    df_cierre2['DF'] = 2 #Distinguir de dónde viene

                    df_cierre = pd.concat([df_cierre1, df_cierre2]) #Concateno en un único df 
                    df_cierre['DNI'] = df_cierre['DNI'].apply(int)
                except:
                    print('No hay archivo de Cierre')
                
                df_concat['STATUS'] = df_concat['STATUS'].astype(str)
                #Separo por válidos y no válidos
                df_validos = df_concat[df_concat['STATUS']=='Válido']
                df_no_validos = df_concat[df_concat['STATUS']!='Válido']
              
                lista_dnis = list(df_no_validos['DNI'])
                df_nuevo = pd.DataFrame({})

                for index_dfc in range(len(df_cierre)):
                    if df_cierre.iloc[index_dfc]['DNI'] in lista_dnis:
                        registro = pd.DataFrame({
                            'DNI': [df_cierre.iloc[index_dfc]['DNI']],
                            'STATUS': ['Válido por Cierre'],
                            'IMPORTE': [df_cierre.iloc[index_dfc]['IMPORTE']],
                            'Operador': [i]
                            })
                        df_nuevo = pd.concat([df_nuevo, registro])
                        #Elimino el dni de la lista
                        lista_dnis.remove(df_cierre.iloc[index_dfc]['DNI'])
        
                df_resto = pd.DataFrame({
                    'DNI': lista_dnis
                })
                
                df_resto = df_resto.merge(df_no_validos, on='DNI', how='inner')
                df_cierre2 = pd.concat([df_validos, df_nuevo, df_resto])
                df_cierre2['STATUS'].replace('Falta Comprobante', 'Pago NO Imputado - Falta Comprobante', inplace=True)
                df_cierre2['STATUS'].replace('Falta Excel', 'Pago NO Imputado - Falta Excel', inplace=True)
                print('Resultante del df_cierre2')
                print(df_cierre2)

        
                #df_validos = df_cierre2[df_cierre2['STATUS']=='Válido por Cierre']
            
            # ------------------ x --------------------- #
            # df_concat['STATUS'] = df_concat['STATUS'].astype(str)
            # searchfor = ['Válido', 'Válido por Unificado']
            # df_validos = df_concat[df_concat['STATUS'].str.contains('|'.join(searchfor))]
            # df_no_validos = df_concat[~df_concat['STATUS'].str.contains('|'.join(searchfor))]
            # def normalizador(numero):
            #     numero = str(numero)
            #     if ',' in numero or '.' in numero or '$' in numero:
            #         numero = numero.replace('$', '')
            #         numero = numero.replace(',', '')
            #         numero = numero.split('.')[0]
            #         return int(numero)
            #     else:
            #         return int(numero)
            # ------------------ x --------------------- #

            # searchfor = ['Válido', 'Válido por Unificado', 'Válido por Cierre']
            # df_final2 = df_concat[df_concat['STATUS'].str.contains('|'.join(searchfor))]
            # df_final3 = df_concat[~df_concat['STATUS'].str.contains('|'.join(searchfor))]            

            # try:
            #     df_merge['IMPORTE'] = df_merge['IMPORTE'].apply(normalizador)
            
            #     df_merge.sort_values('IMPORTE', inplace=True, ascending=False)
            #     df_merge.drop_duplicates('DNI', keep='last', inplace=True)
            #     df_merge.reset_index(drop=True, inplace=True)

            #     sumatoria = df_merge['IMPORTE'].sum()
            #     lista = []
            #     for ix in range(len(df_merge)):
            #         if ix == 0:
            #             lista.append(sumatoria)
            #         else:
            #             lista.append(np.nan)
            #     nuevo_df = pd.DataFrame({'TOTAL': lista})
            #     df_merge2 = pd.concat([df_merge, nuevo_df], axis=1)

            #     df_merge2 = pd.concat([df_merge2, df_final3])
                
            #     df_concat = df_merge2.copy()
            #     df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
            #     lista_dfs.append(df_concat)
            # except:
            #     pass 

        #-------------------- Procedimiento sin ser MejorCredito --------------------#
        else:
            #------------------------------------------------------------#
            #df_final2 = df_concat[df_concat['STATUS'].str.contains('Válido')]
            df_final2 = df_concat[df_concat['STATUS']=='Válido']
            df_final3 = df_concat[df_concat['STATUS']!='Válido']
            #De df_final2 obtengo el importe al chocar contra excel
            excel3 = df[['DNI', 'IMPORTE']] #Elimino la columna de operador
            
            excel3.dropna(inplace=True)
            excel3['DNI'] = excel3['DNI'].apply(validez)

            df_merge = df_final2.merge(excel3, on='DNI', how='left')
            def normalizador(numero):
                numero = str(numero)
                if ',' in numero or '.' in numero or '$' in numero:
                    numero = numero.replace('$', '')
                    numero = numero.replace(',', '')
                    numero = numero.split('.')[0]
                    return int(numero)
                else:
                    return int(numero)

            #print(df_merge)
            
            df_merge['IMPORTE'] = df_merge['IMPORTE'].apply(normalizador)

            #Acomodo los valores de mayor a menor para luego eliminar los 
            #valores duplicados y quedarme con el primer registro que 
            #corresponde al importe más alto.
            df_merge.sort_values('IMPORTE', inplace=True, ascending=False)
            df_merge.drop_duplicates('DNI', keep='last', inplace=True)
            df_merge.reset_index(drop=True, inplace=True)

            sumatoria = df_merge['IMPORTE'].sum()
            lista = []
            for ix in range(len(df_merge)):
                if ix == 0:
                    lista.append(sumatoria)
                else:
                    lista.append(np.nan)
            nuevo_df = pd.DataFrame({'TOTAL': lista})
            df_merge2 = pd.concat([df_merge, nuevo_df], axis=1)

            df_merge2 = pd.concat([df_merge2, df_final3])

            df_concat = df_merge2.copy()
            #------------------------------------------------------------#

            df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
            lista_dfs.append(df_concat)

        #df_validos = df_validos[['DNI', 'STATUS', 'IMPORTE', 'Operador']]
        #df_contra_distribucion.append(df_cierre2)
        #print('Lo que va a df_contra_distribucion')
        #print(df_cierre2)
        sumatoria = df_cierre2['IMPORTE'].sum()
        lista = []
        for ix in range(len(df_cierre2)):
            if ix == 0:
                lista.append(sumatoria)
            else:
                lista.append(np.nan)

        nuevo_df = pd.DataFrame({'TOTAL': lista})
        #print(len(nuevo_df))
        nuevo_df.reset_index(drop=True, inplace=True)
        df_cierre2.reset_index(drop=True, inplace=True)
        #print(len(df_cierre2))
        df_cierre2 = pd.concat([df_cierre2, nuevo_df], axis=1)
        #df_merge2 = pd.concat([df_merge2, df_final3])

        df_cierre2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
        df_contra_distribucion.append(df_cierre2)

#Choque contra la distribución
# ----------------------------------------------------------------------------------------------- #
    #df_contra_distribucion va a chocar contra la distribución
    #Se genera un excel luego de analizar todos los Operadores
    df_validos_gral = pd.concat([df for df in df_contra_distribucion])
    
    #df_validos_gral.to_excel('df_validos_gral.xlsx', index=False)
    #---14/12---#
    lista_dnis = list(df_validos_gral['DNI'])
    df_nuevo = pd.DataFrame({})
    #Quiero los cierres quitando los válidos
    for index_dfc in range(len(df_cierre)):
        if df_cierre.iloc[index_dfc]['DNI'] not in lista_dnis:
            registro = pd.DataFrame({
                'DNI': [df_cierre.iloc[index_dfc]['DNI']],
                'IMPORTE': [df_cierre.iloc[index_dfc]['IMPORTE']],
                # 'SALDO': [df_cierre.iloc[index_dfc]['SALDO']],
                # 'Operador': 'Melanie Gonzalez'
                })
            df_nuevo = pd.concat([df_nuevo, registro])
            #Elimino el dni de la lista
            #lista_dnis.remove(df.iloc[index_dfc]['DNI'])
    #df_resto = pd.DataFrame({'DNI': lista_dnis})
    #df_nuevo son los registros del cierre que no son válidos

    #Choque contra la distribución
    df_distribucion = pd.read_excel('M:/FEDERICO/DISTRIBUCION MEJOR CREDITO 0112.xlsx',
    usecols=['DNI', 'OPERADOR'])
    #Quito los no gestionables de la distribución
    df_distribucion2 = df_distribucion[df_distribucion['OPERADOR'] != 'MARINA']

    #Los registros del cierre que están en la distribución
    try:
        df_merge = df_nuevo.merge(df_distribucion2, on='DNI', how='inner')
        df_merge.rename({'OPERADOR': 'ASIGNADO'}, axis=1, inplace=True)
        if os.path.exists(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_sin_gestion.xlsx'):
            a_concatenar = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_sin_gestion.xlsx')
            df_concat = pd.concat([a_concatenar, df_merge])
            df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_sin_gestion.xlsx', index=False)
        else:
            df_merge.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_sin_gestion.xlsx', index=False)
    except:
        print('No hay archivo generado')

    #Busco los registros del cierre que NO están en la distribución
    lista_dnis = list(df_distribucion['DNI']) #dnis de la distribución
    df_nuevo2 = pd.DataFrame({})
    for index_dfc in range(len(df_nuevo)):
        if df_nuevo.iloc[index_dfc]['DNI'] not in lista_dnis:
            registro = pd.DataFrame({
                'DNI': [df_nuevo.iloc[index_dfc]['DNI']],
                'IMPORTE': [df_nuevo.iloc[index_dfc]['IMPORTE']],
                #'SALDO': [df_nuevo.iloc[index_dfc]['SALDO']],
                #'Operador': 'Melanie Gonzalez'
                })
            df_nuevo2 = pd.concat([df_nuevo2, registro])
            #Elimino el dni de la lista
            #lista_dnis.remove(df.iloc[index_dfc]['DNI'])
    #df_resto = pd.DataFrame({'DNI': lista_dnis})
    #df_nuevo son los registros del cierre que no son válidos
    #print(df_nuevo2.shape)
    if os.path.exists(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_dni_no_distribuido.xlsx'):
        a_concatenar = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_dni_no_distribuido.xlsx')
        df_concat = pd.concat([a_concatenar, df_nuevo2])
        df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_dni_no_distribuido.xlsx', index=False)
        #print(f"Archivo 'Break_{empresa}' concatenado correctamente")
    else:
        df_nuevo2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_dni_no_distribuido.xlsx', index=False)
        #print(f"Archivo 'Break_{empresa}' creado correctamente")
    #df_nuevo2.to_excel('M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados/df_dni_no_distribuido.xlsx', index=False)

    #-------------------- Excel de duplicados --------------------#
    df_todos = df_validos_gral.copy()
    #df_todos = pd.concat([x for x in lista_dfs])
    if df_todos.empty:
        df_dupli = pd.DataFrame({'DNI': [],
                                'STATUS': [],
                                'Operador': [],
                                'Repetido': []})
                                
        df_dupli.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Duplicados.xlsx', index=False)
    else:
        df_todos_validos = df_todos[df_todos['STATUS'].str.contains('Válido')]
        df_dupli = df_todos_validos[df_todos_validos['DNI'].duplicated(keep=False)]
        
        if df_dupli.empty:
            df_dupli = pd.DataFrame({'DNI': [],
                                    'STATUS': [],
                                    'Operador': [],
                                    'Repetido': []})
            df_dupli.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Duplicados.xlsx', index=False)
        else:
            df_dupli = df_dupli[['STATUS', 'Operador', 'DNI', 'IMPORTE']]
            df_dupli.sort_values(by=['DNI', 'Operador'], inplace=True)

            def duplicado(x, y):
                global indice
                try:
                    #if (x == df_dupli.iloc[indice+1][0]) and (y != df_dupli.iloc[indice+1][2]):
                    if (x == df_dupli.iloc[indice+1][2]) and (y != df_dupli.iloc[indice+1][1]):
                        indice += 1
                        return f'Repetido con {df_dupli.iloc[indice][1]}'
                    elif (x == df_dupli.iloc[indice-1][2]) and (y != df_dupli.iloc[indice-1][1]):
                        indice += 1
                        return f'Repetido con {df_dupli.iloc[indice-2][1]}'
                    else:
                        indice += 1
                except:
                    if (x == df_dupli.iloc[indice-1][2]) and (y != df_dupli.iloc[indice-1][1]):
                        indice+=1
                        return f'Repetido con {df_dupli.iloc[indice-2][1]}'
            df_dupli['Repetido'] = df_dupli.apply(lambda x: duplicado(x.DNI, x.Operador), axis=1) 
            df_dupli = df_dupli[['DNI', 'STATUS', 'Operador', 'IMPORTE', 'Repetido']]
            df_dupli.dropna(inplace=True)

            df_dupli.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Duplicados.xlsx', index=False)
            
    print(f"{empresa} actualizado.")
    print('')
    print('')
    print('Presiona cualquier tecla para cerrar la ventana... ')
    msvcrt.getch()