import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np
import os
import re
import msvcrt

indice = 0
def funcion(empresa):
    os.system('cls')
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
                #df es el excel 
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
    for i in lista_valida_operadores:
        #Dni que figuran en el excel
        df_dni2 = df_dni_inicial[df_dni_inicial['Operador'].str.contains(f'{i}')]

        #try:
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

        #df_dni2['IMPORTE'] = df_dni2['IMPORTE'].apply(validez)
        try:
            df_dni2['IMPORTE'] = df_dni2['IMPORTE'].apply(normalizador)
        except:
            print(f'Fallo en la columna DNI de {i}')

        #Comprobantes
        df_comprobantes2 = df_comprobantes[df_comprobantes['Operador'].str.contains(f'{i}')]
        df_comprobantes2['DNI'] = df_comprobantes2['DNI'].apply(validez)
        #df_comprobantes2['DNI'] = df_comprobantes2['DNI'].apply(int)

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
        
        #print(df_concat)
        # print(f'\nContador Inicial: {len(df_concat)}')
        # print(df_concat['STATUS'].value_counts())

        try:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', usecols=['DNI', 'IMPORTE'])
        except:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Financieras/{empresa}/{i}/{i}.xlsx', skiprows=1, usecols=['DNI', 'IMPORTE'])

        #A partir de acá realizar la validación con el Unificado
        #Archivo de unificado funciona sólo para MejorCredito
        if empresa == 'MejorCredito':
            lista_unificados = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal')
            if len(lista_unificados) == 1:
                for unificado in lista_unificados:
                    df_unificado = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{unificado}', usecols=['DNI', 'MONTO'])
                
                #Realizar la validación
                #Dejo df_filtro1 con los 'Falta Comprobante' y realizar su validación
                #y creo df_filtro2 con los demás para luego concatenar df_filtro1 con df_filtro2
                #df_filtro1 = df_concat[df_concat['STATUS'].str.contains('Falta Comprobante')]
                df_filtro1 = df_concat[df_concat['STATUS']=='Falta Comprobante']
                #df_filtro2 = df_concat[~df_concat['STATUS'].str.contains('Falta Comprobante')]
                df_filtro2 = df_concat[df_concat['STATUS']!='Falta Comprobante']

                df_filtro1 = df_filtro1[['DNI', 'STATUS', 'Operador']]

                lista_unificado = list(df_unificado['DNI'])
                lista_df_filtro1 = list(df_filtro1['DNI'])

                dni_estado = []
                for dni in lista_df_filtro1:
                    if dni in lista_unificado:
                        dni_estado.append('Válido por Unificado')
                        lista_unificado.remove(dni)
                    else:
                        dni_estado.append('Falta Comprobante')

                df_status_excel = pd.DataFrame({'DNI': lista_df_filtro1,
                                                'STATUS': dni_estado,
                                                })
                
                df_concat = pd.concat([df_status_excel, df_filtro2])
                df_concat['Operador'] = f'{i}'

                #print(df_concat.info())
                df_concat['STATUS'] = df_concat['STATUS'].astype(str)
                searchfor = ['Válido', 'Válido por Unificado']
                df_validos = df_concat[df_concat['STATUS'].str.contains('|'.join(searchfor))]
                df_no_validos = df_concat[~df_concat['STATUS'].str.contains('|'.join(searchfor))]

                #Validacion de los no_validos con el cierre
                #df_cierre = pd.read_excel(, usecols=['DNI', 'MONTO'])
                df_cierre1 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Cierre/FINAL NOVIEMBRE 2022.xlsx', sheet_name='COBRANZA MEJOR CREDITO', usecols=['doc.', 'fecha', 'est'])
                df_cierre2 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Cierre/FINAL NOVIEMBRE 2022.xlsx', sheet_name='RENDICION MAS COBRANZAS', usecols=['DNI', 'Fecha de Pago', ' Importe Total '])
                df_cierre1 = df_cierre1.iloc[:-1] #Elimino el total calculado a mano
                df_cierre2 = df_cierre2.iloc[:-1] #Elimino el total calculado a mano

                df_cierre1.rename({'doc.': 'DNI', 'est': 'IMPORTE', 'fecha': 'Fecha de Pago'}, axis=1, inplace=True)
                df_cierre2.rename({' Importe Total ': 'IMPORTE'}, axis=1, inplace=True)
                df_cierre1['DF'] = 1
                df_cierre2['DF'] = 2

                df_cierre = pd.concat([df_cierre1, df_cierre2])
                df_cierre['DNI'] = df_cierre['DNI'].apply(int)

                #Etapa de validación
                #--------------------#
                lista_cierre = list(df_cierre['DNI'])
                df_no_validos = list(df_no_validos['DNI'])

                dni_estado = []
                for dni in df_no_validos:
                    if dni in lista_cierre:
                        dni_estado.append('Válido por Cierre')
                        lista_cierre.remove(dni)
                    else:
                        dni_estado.append('Falta Comprobante')

                df_status_excel = pd.DataFrame({'DNI': df_no_validos,
                                                'STATUS': dni_estado,
                                                })
                
                df_concat = pd.concat([df_status_excel, df_filtro2])
                df_concat['Operador'] = f'{i}'

                #print(df_concat.info())
                df_concat['STATUS'] = df_concat['STATUS'].astype(str)
                searchfor = ['Válido', 'Válido por Unificado']
                df_validos = df_concat[df_concat['STATUS'].str.contains('|'.join(searchfor))]
                df_no_validos = df_concat[~df_concat['STATUS'].str.contains('|'.join(searchfor))]
                #--------------------#
                def normalizador(numero):
                    numero = str(numero)
                    if ',' in numero or '.' in numero or '$' in numero:
                        numero = numero.replace('$', '')
                        numero = numero.replace(',', '')
                        numero = numero.split('.')[0]
                        return int(numero)
                    else:
                        return int(numero)

                excel3 = df[['DNI', 'IMPORTE']] #Elimino la columna de operador
                excel3.dropna(inplace=True)
                excel3['DNI'] = excel3['DNI'].apply(validez)

                #Antes del merge tienen que tener valor en $ los válidos y válido por unificado
                #Separo los que tienen Válido y Válido por Unificado

                df_merge = df_validos.merge(excel3, on='DNI', how='left')
                df_merge['IMPORTE'] = df_merge['IMPORTE'].apply(normalizador)

                #Todos los válidos con su importe. Realizar la suma
                #df_merge['IMPORTE'] = df_merge['IMPORTE'].apply(int)
                sumatoria = df_merge['IMPORTE'].sum()
                lista = []
                for ix in range(len(df_merge)):
                    if ix == 0:
                        lista.append(sumatoria)
                    else:
                        lista.append(np.nan)
                nuevo_df = pd.DataFrame({'TOTAL': lista})
                df_merge2 = pd.concat([df_merge, nuevo_df], axis=1)

                df_final = pd.concat([df_merge2, df_no_validos])

                df_concat = df_final.copy()
                df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
                lista_dfs.append(df_concat)
            else:
                print(f'Más de 1 archivo en la carpeta Sucursal - Unificado de {empresa}')









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

    #-------------------- Excel de duplicados --------------------#
    df_todos = pd.concat([x for x in lista_dfs])
    
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
            #df_dupli.dropna(inplace=True)

            df_dupli.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Duplicados.xlsx', index=False)
            
    print(f"{empresa} actualizado.")
    print('')
    print('')
    print('Presiona cualquier tecla para cerrar la ventana... ')
    msvcrt.getch()