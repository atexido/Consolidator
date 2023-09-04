import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np
import os
import re
import msvcrt

indice = 0
def funcion(empresa):
    os.system('cls')
    try:
        os.remove(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_sin_gestion.xlsx')
        os.remove(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx')
    except:
        pass

    # ------------------------------------------------------------------------------------------- #
    formatos_validos = ['jpg', 'JPG', 'jpeg', 'JPEG', 'pdf', 'jfif', 'PDF', 'html', 'png', 'PNG']
    # ------------------------------------------------------------------------------------------- #

    lista_operadores = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/')
    if len(lista_operadores) == 0:
        print('Sin carpetas creadas')
        print('')
        print('Presiona cualquier tecla para cerrar la ventana... ')
        msvcrt.getch()
        return 0
    
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
        lista_archivos = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}')
        lista_archivos = list(map(lambda x: x.title(), lista_archivos)) #2022, Comprobantes, Operador.xlsx
        lista_archivos = list(map(lambda x: x.split('.')[0], lista_archivos)) #Quito el .xlsx
        
        if i in lista_archivos:
            print(f'{i} √')
            lista_valida_operadores.append(i)

            cantidad_excels = 0
            if '~$' not in lista_archivos: #Evita los archivos abierto
                cantidad_excels += 1
                try:
                    df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}/{i}.xlsx', usecols=['DNI', 'IMPORTE'])
                except:
                    df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}/{i}.xlsx', skiprows=1, usecols=['DNI', 'IMPORTE'])
                df = df.dropna()
                df['Operador'] = f'{i}'
                df_dni_inicial = pd.concat([df_dni_inicial, df])

            if cantidad_excels != 1:
                print(f'Cantidad de archivos Excel: {cantidad_excels}')

            #-------------------- Comprobantes --------------------#
            try:
                lista_comprobantes = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}/Comprobantes')
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

# Iteración sobre los operadores
# ------------------------------------------------------------------------------------------ #
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
        # def normalizador(numero):
        #     numero = str(numero)
        #     if ',' in numero or '.' in numero or '$' in numero:
        #         numero = numero.replace('$', '')
        #         numero = numero.replace(',', '')
        #         numero = numero.split('.')[0]
        #         return int(numero)
        #     else:
        #         try:
        #             return int(numero)
        #         except:
        #             return 0
        def normalizador(numero):
            numero = str(numero)
            if len(numero) > 0: 
                valido = re.findall(patron, numero)
                if len(valido) > 0:
                    return(int(valido[0]))
                else:
                    return 0
            else:
                return 0

        try:
            df_dni2['IMPORTE'] = df_dni2['IMPORTE'].apply(normalizador)
        except:
            print(f'Fallo en la columna DNI de {i}')

        #Comprobantes
        df_comprobantes2 = df_comprobantes[df_comprobantes['Operador'].str.contains(f'{i}')]
        df_comprobantes2['DNI'] = df_comprobantes2['DNI'].apply(validez)
        
        lista_df_dni2 = list(df_dni2['DNI'])
        #print(lista_df_dni2)
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
        #print(df_status_excel)

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
        #print(f'df_concat de {i}')
        #print(df_concat)

        #Coloco el importe
        try:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}/{i}.xlsx', usecols=['DNI', 'IMPORTE'])
        except:
            df = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/BackUp/Servicios/{empresa}/{i}/{i}.xlsx', skiprows=1, usecols=['DNI', 'IMPORTE'])

        #Separo por válidos y no válidos
        df_concat['STATUS'] = df_concat['STATUS'].apply(str)
        df_validos = df_concat[df_concat['STATUS'] == 'Válido']
        df_no_validos = df_concat[df_concat['STATUS'] != 'Válido']

        try:
            excel3 = df[['DNI', 'IMPORTE']]
        except:
            print(f'Fallo en el excel de: {i}. Falta las columnas DNI e IMPORTE')

        excel3.dropna(inplace=True)
        #print('excel3 previo a normalizar')
        #print(excel3)

        excel3['DNI'] = excel3['DNI'].apply(validez)
        excel3['IMPORTE'] = excel3['IMPORTE'].apply(normalizador)

        lista_validos = list(df_validos['DNI'])  
        #print(f'Lista de válidos de {i}: {lista_validos}')

        df_nuevo = pd.DataFrame({
            'DNI': [],
            'IMPORTE': []
        })
        #print('Excel 3')
        #print(excel3)
        for indice_df in range(len(excel3)):
            #print(excel3.iloc[indice_df]['DNI'])
            if excel3.iloc[indice_df]['DNI'] in lista_validos:
                registro = pd.DataFrame({
                    #'DNI': [excel3.iloc[indice_df]['DNI']],
                    'DNI': [int(excel3.iloc[indice_df]['DNI'])],
                    #'IMPORTE': [excel3.iloc[indice_df]['IMPORTE']],
                    'IMPORTE': [int(excel3.iloc[indice_df]['IMPORTE'])],
                    })
                #print(registro)
                df_nuevo = pd.concat([df_nuevo, registro])
                #print(df_nuevo)
                #Elimino el dni de la lista
                lista_validos.remove(excel3.iloc[indice_df]['DNI'])
        #df_nuevo = df_nuevo.astype(int)
        #print(f'df_nuevo luego de convertir DNI a int')
        #print(df_nuevo)
        #print(df_nuevo.info())

        #print('resultante de los comprobantes contra el excel')
        df_nuevo['STATUS'] = 'Válido'
        df_nuevo['Operador'] = i

        #print(f'df_nuevo de: {i}')
        #print(df_nuevo)
    
        df_concat = pd.concat([df_nuevo, df_no_validos])
        #print('Primer Concat. Los válidos con su importe + No válidos sin importe')
        #print(df_concat)
        lista = set(list(df_concat['DNI']))
        #print(f'Cantidad de únicos: {len(lista)}')
        df_operador = df_concat.copy()
        #print(f'df_operador de {i}')
        #print(df_operador)
        
        #print('-' * 100)
        #Hasta acá están los VALIDOS con su importe y nan FALTA COMPROBANTE y FALTA EXCEl
        #print(df_operador)
        
# --------------- Verificación si hay archivos de Unificado o Cierre --------------- #
        lista_archivos = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal')
        #Valido que haya archivos dentro de la carpeta
        if len(lista_archivos) > 0:
            lista_archivos = [i.lower() for i in lista_archivos] #Todo a minúsculas

            #--------------- Archivos ---------------#
            #Cada archivo debe tener la palabra "unificado", "cierre" o "distribucion" según corresponda
            #Cada archivo debe tener la columna DNI e IMPORTE
            unificado = [i for i in lista_archivos if 'unificado' in i]
            try:
                unificado = unificado[0]
            except:
                unificado = ''

            cierre = [i for i in lista_archivos if 'cierre' in i]
            try:
                cierre = cierre[0]
            except:
                cierre = ''

            # distribucion = [i for i in lista_archivos if 'distribucion' in i]
            # try:
            #     distribucion = distribucion[0]
            # except:
            #     distribucion = ''
            #----------------------------------------#
            if cierre == '': 
                #print('NO HAY CIERRE')      
                df_unificado = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{unificado}', usecols=['DNI', 'IMPORTE'])
                #Dejo df_filtro1 con los 'Falta Comprobante' y realizar su validación
                #y creo df_filtro2 con los demás para luego concatenar df_filtro1 con df_filtro2}

                #No realizo el filtro de Falta Comprobante
                #Hago una comparación de todos (inluyendo los válidos)
                # ------------------------------------ #
                #df_filtro0 = df_concat[['DNI', 'STATUS', 'Operador']]
                # ------------------------------------ #

                #df_filtro1 = df_concat[df_concat['STATUS']=='Falta Comprobante']
                #Los "Falta Comprobante" son los que van a chocar con el unificado
                #df_filtro2 = df_concat[df_concat['STATUS']!='Falta Comprobante']
                #df_filtro1 = df_filtro1[['DNI', 'STATUS', 'Operador']]

                #lista_dnis = list(df_filtro1['DNI'])
                
                lista_dnis = list(df_operador['DNI']) #Todos los dni de df_concat
                #print(lista_dnis)
                #Estos dnis van a chocar con el unificado:
                #Puedo separar por Válido, y !Válido
                #Al conjunto de los Válidos. Si están en el unificado: Válido.
                #Sino, Pago NO Imputado

                #Al conjunto de los !Válidos. Si están: Válido por Unificado
                #Sino, Pago No Imputado
                #Luego concatenar estos 2 dfs para el resultado final

                #Relación 1 a 1. Para cada comprobante debe tener su registro en excel

                #2 cuotas. Las 2 deben estar en el excel
                #Si tiene 1 comprobante:
                #x válido y x Falta Comprobante
                #El válido, obtiene la sumatoria, el dni se elimina de la lista,
                #Para que aparezca el falta comprobante
                #Luego en el final, se tiene que eliminar este último registro

                #2 comprobantes y 1 excel

                df_nuevo = pd.DataFrame({})

                #for dni i lista_dnis:
                    #if dni in df_unificado['DNI']:

                # for index_dfu in range(len(df_unificado)):
                #     #Si el DNI está en la lista_dnis

                #     if df_unificado.iloc[index_dfu]['DNI'] in lista_dnis:
                #         dni = df_unificado.iloc[index_dfu]['DNI']
                #         importe = df_unificado[df_unificado['DNI'] == dni]
                #         suma_importe = importe['IMPORTE'].sum()

                #         registro = pd.DataFrame({
                #             'DNI': [dni],
                #             'STATUS': ['Válido por Unificado'],
                #             #'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                #             'IMPORTE': [suma_importe],
                #             'Operador': [i]
                #             })
                #         df_nuevo = pd.concat([df_nuevo, registro])
                #         #Elimino el dni de la lista
                #         lista_dnis.remove(df_unificado.iloc[index_dfu]['DNI'])

                    # else:
                    #     dni = df_unificado.iloc[index_dfu]['DNI']
                    #     registro = pd.DataFrame({
                    #         'DNI': [dni],
                    #         'STATUS': ['Pago No Imputado'],
                    #         #'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                    #         'IMPORTE': [np.nan],
                    #         'Operador': [i]
                    #         })
                    #     df_nuevo = pd.concat([df_nuevo, registro])

                #Tengo que iterar la lista. Si está el DNI: Válido por unificado con la sumatoria
                #Sino: "Pago NO Imputado"
                #print(f'Cantidad de lista_dnis {len(lista_dnis)}')
                contador = 0
                lista_dni_unificado = list(df_unificado['DNI'])

                for doc in lista_dnis:
                    if doc in lista_dni_unificado:
                        contador += 1
                        importe = df_unificado[df_unificado['DNI'] == doc] 
                        suma_importe = importe['IMPORTE'].sum()

                        registro = pd.DataFrame({
                            'DNI': [doc],
                            'STATUS': ['Válido por Unificado'],
                            #'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                            'IMPORTE': [suma_importe],
                            'Operador': [i]
                            })

                        df_nuevo = pd.concat([df_nuevo, registro])
                        #lista_dnis.remove(doc)
                    else:
                        contador += 1
                        registro = pd.DataFrame({
                            'DNI': [doc],
                            'STATUS': ['Pago NO Imputado'],
                            #'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                            'IMPORTE': [np.nan],
                            'Operador': [i]
                            })

                        df_nuevo = pd.concat([df_nuevo, registro])
                        #lista_dnis.remove(doc)

                #print(f'contador: {contador}')
                # df_resto = pd.DataFrame({'DNI': lista_dnis})
                # df_resto['IMPORTE'] = np.nan
                # df_resto['STATUS'] = 'Falta Comprobante'
                # df_resto['Operador'] = i
                # df_unificado = pd.concat([df_nuevo, df_resto])
                #print('df_nuevo')
                df_nuevo.drop_duplicates(keep='first', inplace=True)
                try:
                    df_nuevo.sort_values(by='DNI', inplace=True)
                except:
                    pass
                #print(df_nuevo)

                df_concat = df_nuevo.copy()
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

                #No necesito concatenar porque paso todos los dni con el unificado
                #df_concat = pd.concat([df_unificado, df_filtro2])

                try:
                    df_concat['DNI'] = df_concat['DNI'].apply(int)
                    #print('Concat luego de unir los que no chocaron contra los nuevos resultados')
                    #print(df_concat)
                    #Quitar válidos si ya está Válido por Unificado
                    lista = set(list(df_concat['DNI'])) #Dnis únicos
                    df_concat2 = pd.DataFrame({})

                    for dni in lista: 
                        reg = df_concat[df_concat['DNI'] == dni]
                        if len(reg) > 1: 
                            registro = reg[reg['STATUS']!='Válido']
                            df_concat2 = pd.concat([df_concat2, registro])
                        else:
                            df_concat2 = pd.concat([df_concat2, reg])
                    df_concat = df_concat2.copy()   
                except:
                    pass 

                try:   
                    df_concat.sort_values(by='STATUS', ascending=False, inplace=True)              
                except:
                    pass

                df_operador = df_concat.copy()
                
            else:
                #print('HAY CIERRE')
                #Cada empresa tiene su cierre. Debe tener la columna DNI e IMPORTE
                # if empresa == 'MejorCredito':
                #     df_cierre = cierre_mejorcredito(empresa, cierre)
                # elif empresa == 'Argenpesos':
                #     df_cierre = cierre_argenpesos(empresa, cierre)
                # elif empresa == 'Consumax':
                #     df_cierre = cierre_consumax(empresa, cierre)
                # elif empresa == 'Credipesos':
                #     df_cierre = cierre_credipesos(empresa, cierre)
                # elif empresa == 'Credisol':
                #     df_cierre = cierre_credisol(empresa, cierre)
                # elif empresa == 'CredNow':
                #     df_cierre = cierre_crednow(empresa, cierre)
                # elif empresa == 'Cristal Cash':
                #     df_cierre = cierre_cristalcash(empresa, cierre)
                # elif empresa == 'Kalima':
                #     df_cierre = cierre_kalima(empresa, cierre)
                # elif empresa == 'Qida':
                #     df_cierre = cierre_qida(empresa, cierre)

                #Cierre a mano con las columnas de DNI e IMPORTE
                df_cierre = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{cierre}', usecols=['DNI', 'IMPORTE'])

#Nuevo código cierre 
#----------------------------------------------------------------------------------------------#
                try:
                    lista_dnis = list(df_operador['DNI']) #Todos los dni de df_concat
                    #print(lista_dnis) 

                    df_nuevo = pd.DataFrame({})

                    #print(f'Cantidad de lista_dnis {len(lista_dnis)}')
                    contador = 0
                    lista_dni_cierre = list(df_cierre['DNI'])

                    for doc in lista_dnis:
                        if doc in lista_dni_cierre:
                            contador += 1
                            importe = df_cierre[df_cierre['DNI'] == doc]
                            suma_importe = importe['IMPORTE'].sum()

                            registro = pd.DataFrame({
                                'DNI': [doc],
                                'STATUS': ['Válido por Cierre'],
                                'IMPORTE': [suma_importe],
                                'Operador': [i]
                                })

                            df_nuevo = pd.concat([df_nuevo, registro])
                            #lista_dnis.remove(doc)
                        else:
                            contador += 1
                            registro = pd.DataFrame({
                                'DNI': [doc],
                                'STATUS': ['Pago NO Imputado'],
                                #'IMPORTE': [df_unificado.iloc[index_dfu]['IMPORTE']],
                                'IMPORTE': [np.nan],
                                'Operador': [i]
                                })

                            df_nuevo = pd.concat([df_nuevo, registro])
                            #lista_dnis.remove(doc)

                    #print(f'contador: {contador}')
                    # df_resto = pd.DataFrame({'DNI': lista_dnis})
                    # df_resto['IMPORTE'] = np.nan
                    # df_resto['STATUS'] = 'Falta Comprobante'
                    # df_resto['Operador'] = i
                    # df_unificado = pd.concat([df_nuevo, df_resto])
                    #print('df_nuevo')
                    df_nuevo.drop_duplicates(keep='first', inplace=True)
                    try:
                        df_nuevo.sort_values(by='DNI', inplace=True)
                    except:
                        pass
                    #print(df_nuevo)

                    df_concat = df_nuevo.copy()
                except:
                    nulo = pd.DataFrame({'DNI': [], 'STATUS': [], 'IMPORTE': [], 'Operador': []})
                    #df_contra_distribucion.append(nulo)
                    df_concat = nulo.copy()

                try:
                    df_concat['DNI'] = df_concat['DNI'].apply(int)
                    #print('Concat luego de unir los que no chocaron contra los nuevos resultados')
                    #print(df_concat)
                    #Quitar válidos si ya está Válido por Unificado
                    lista = set(list(df_concat['DNI'])) #Dnis únicos
                    df_concat2 = pd.DataFrame({})

                    for dni in lista: 
                        reg = df_concat[df_concat['DNI'] == dni]
                        if len(reg) > 1: 
                            registro = reg[reg['STATUS']!='Válido']
                            df_concat2 = pd.concat([df_concat2, registro])
                        else:
                            df_concat2 = pd.concat([df_concat2, reg])
                    df_concat = df_concat2.copy()   
                except:
                    pass 

                try:   
                    df_concat.sort_values(by='STATUS', ascending=False, inplace=True)              
                except:
                    pass

                df_operador = df_concat.copy()

#----------------------------------------------------------------------------------------------#
                #------------------------------------------------------------------#
                # df_concat['STATUS'] = df_concat['STATUS'].astype(str)
                # #Separo por válidos y no válidos
                # df_validos = df_concat[df_concat['STATUS']=='Válido']
                # df_no_validos = df_concat[df_concat['STATUS']!='Válido']
                
                # lista_dnis = list(df_no_validos['DNI'])
                # df_nuevo = pd.DataFrame({})

                # for index_dfc in range(len(df_cierre)):
                #     if df_cierre.iloc[index_dfc]['DNI'] in lista_dnis:
                #         registro = pd.DataFrame({
                #             'DNI': [df_cierre.iloc[index_dfc]['DNI']],
                #             'STATUS': ['Válido por Cierre'],
                #             'IMPORTE': [df_cierre.iloc[index_dfc]['IMPORTE']],
                #             'Operador': [i]
                #             })
                #         df_nuevo = pd.concat([df_nuevo, registro])
                #         #Elimino el dni de la lista
                #         lista_dnis.remove(df_cierre.iloc[index_dfc]['DNI'])
        
                # df_resto = pd.DataFrame({ 'DNI': lista_dnis})
                
                # df_resto = df_resto.merge(df_no_validos, on='DNI', how='inner')
                # df_cierre2 = pd.concat([df_validos, df_nuevo, df_resto])
                # df_cierre2['STATUS'].replace('Falta Comprobante', 'Pago NO Imputado - Falta Comprobante', inplace=True)
                # df_cierre2['STATUS'].replace('Falta Excel', 'Pago NO Imputado - Falta Excel', inplace=True)
                # print('Resultante del df_cierre2')
                # print(df_cierre2)

                # df_operador = df_cierre2.copy()
        else:
            mensaje_sin_archivos = 'SIN UNIFICADO / CIERRE / DISTRIBUCIÓN'
            df_contra_distribucion.append(df_operador)
            #NO HAY ARCHIVOS

            #Que sólo se muestren los válidos e inválidos por Excel-Comprobante
        #------------------------------------------------------------#
        df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
        #print('df_concat guardado:')
        #print(df_concat)

        lista_dfs.append(df_concat)

        try:
            sumatoria = df_operador['IMPORTE'].sum()
            lista = []
            for ix in range(len(df_operador)):
                if ix == 0:
                    lista.append(sumatoria)
                else:
                    lista.append(np.nan)

            nuevo_df = pd.DataFrame({'TOTAL': lista})
            nuevo_df.reset_index(drop=True, inplace=True)
            df_operador.reset_index(drop=True, inplace=True)
            df_operador = pd.concat([df_operador, nuevo_df], axis=1)

            df_operador.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/df {i}.xlsx', index=False)
            df_contra_distribucion.append(df_operador)
        except:
            nulo = pd.DataFrame({'DNI': [], 'STATUS': [], 'IMPORTE': [], 'Operador': []})
            df_contra_distribucion.append(nulo)

# ----------------------------------------------------------------------------------------------- #
#Choque contra la distribución
    lista_archivos = os.listdir(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal')

    #Valido que haya archivos dentro de la carpeta
    if len(lista_archivos) > 0:
        lista_archivos = [i.lower() for i in lista_archivos] #Todo a minúsculas

        distribucion = [i for i in lista_archivos if 'distribucion' in i]
        try:
            distribucion = distribucion[0]
        except:
            distribucion = ''

        unificado = [i for i in lista_archivos if 'unificado' in i]
        try:
            unificado = unificado[0]
        except:
            unificado = ''

        cierre = [i for i in lista_archivos if 'cierre' in i]
        try:
            cierre = cierre[0]
        except:
            cierre = ''

        if unificado != '' and cierre == '':
            print('\nCHOQUE CON UNIFICADO')
        elif unificado == '' and cierre != '':
            print('\nCHOQUE CON CIERRE')
        elif unificado != '' and cierre != '':
            print('\nCHOQUE CON CIERRE')

        if distribucion != '':
            print('CHOQUE CON DISTRIBUCIÓN')
            #df_contra_distribucion va a chocar contra la distribución
            #Se genera un excel luego de analizar todos los Operadores
            df_validos_gral = pd.concat([df for df in df_contra_distribucion])
            
            lista_dnis = list(df_validos_gral['DNI']) #Lista de dni válidos
            df_nuevo = pd.DataFrame({})

            #Quiero los cierres quitando los válidos
            for index_dfc in range(len(df_cierre)):
                if df_cierre.iloc[index_dfc]['DNI'] not in lista_dnis:
                    registro = pd.DataFrame({
                        'DNI': [df_cierre.iloc[index_dfc]['DNI']],
                        'IMPORTE': [df_cierre.iloc[index_dfc]['IMPORTE']],
                        })
                    df_nuevo = pd.concat([df_nuevo, registro])

            #Choque contra la distribución
            #Lectura del archivo de distribución
            print('Leyendo archivo de distribución...')
            df_distribucion = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{distribucion}', usecols=['DNI', 'OPERADOR'])
            
            #Quito los no gestionables de la distribución
            df_distribucion2 = df_distribucion[df_distribucion['OPERADOR'] != 'MARINA']

            #Los registros del cierre que están en la distribución
            #try:
            df_merge = df_nuevo.merge(df_distribucion2, on='DNI', how='inner')
            df_merge.rename({'OPERADOR': 'ASIGNADO'}, axis=1, inplace=True)
            
            df_merge.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_sin_gestion.xlsx', index=False)
            print('')
            print('df_sin_gestion √')
            
            #Busco los registros del cierre que NO están en la distribución
            lista_dnis = list(df_distribucion['DNI']) #dnis de la distribución
            df_nuevo2 = pd.DataFrame({})
            for index_dfc in range(len(df_nuevo)):
                if df_nuevo.iloc[index_dfc]['DNI'] not in lista_dnis:
                    registro = pd.DataFrame({
                        'DNI': [df_nuevo.iloc[index_dfc]['DNI']],
                        'IMPORTE': [df_nuevo.iloc[index_dfc]['IMPORTE']],
                        })
                    df_nuevo2 = pd.concat([df_nuevo2, registro])

            # if os.path.exists(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx'):
            #     a_concatenar = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx')
            #     df_concat = pd.concat([a_concatenar, df_nuevo2])
            #     print('CONCAT!')
            #     df_concat.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx', index=False)
            #else:
            df_nuevo2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/df_dni_no_distribuido.xlsx', index=False)
            print('df_dni_no_distribuido √')
        else:
            df_validos_gral = pd.concat([df for df in df_contra_distribucion])
    else:
        print('\nSIN UNIFICADO / CIERRE / DISTRIBUCION')
        df_validos_gral = pd.concat([df for df in df_contra_distribucion])
        
    #-------------------- Excel de duplicados --------------------#
    df_validos_gral2 = df_validos_gral.query('STATUS not in ["Pago NO Imputado", "Falta Comprobante", "Falta Excel"]')
    df_validos_gral2.drop('TOTAL', axis=1, inplace=True) #Borro los subtotales de cada operador
    sumatoria = df_validos_gral2['IMPORTE'].sum()
    lista = []
    for ix in range(len(df_validos_gral2)):
        if ix == 0:
            lista.append(sumatoria)
        else:
            lista.append(np.nan)

    nuevo_df = pd.DataFrame({'TOTAL': lista})
    nuevo_df.reset_index(drop=True, inplace=True)
    df_validos_gral2.reset_index(drop=True, inplace=True)
    df_validos_gral2 = pd.concat([df_validos_gral2, nuevo_df], axis=1)

    df_validos_gral2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/Unificado de válidos.xlsx', index=False)
    df_validos_gral2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/VALIDOS.xlsx', index=False)
    print('Unificado de válidos creado √')

    #Unificado de Pagos NO Imputados
    #df_no_validos_gral2 = df_validos_gral.query('STATUS == "Pago NO Imputado"')
    df_no_validos_gral2 = df_validos_gral.query('STATUS == "Pago NO Imputado" or STATUS in ["Falta Excel", "Falta Comprobante"]')
    df_no_validos_gral2.drop(['IMPORTE', 'TOTAL'], axis=1, inplace=True) #Borro los subtotales de cada operador
    df_no_validos_gral2.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/Unificado de NO válidos.xlsx', index=False)
    print('Unificado de NO válidos creado √')

#--------------Total-----------#
    df_concat_final = pd.concat([df_validos_gral2, df_no_validos_gral2])
    df_concat_final.to_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Resultados/Unificado Válidos y NO válidos.xlsx', index=False)
    print('Unificado de Válidos y No Válidos creado √')
    
    df_todos = df_validos_gral.copy()    
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
    
    print('')
    print(f"{empresa} actualizado.")
    print('')
    print('')
    print('Presiona cualquier tecla para cerrar la ventana... ')
    msvcrt.getch()