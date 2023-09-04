import pandas as pd 

def cierre_argenpesos(empresa, cierre):
    df_cierre1 = pd.read_excel(f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/{empresa}/Sucursal/{cierre}', sheet_name='Total', usecols=['Documento', 'ImportePago'])
    #df_cierre1 = df_cierre1.iloc[:-1] #Elimino el total calculado a mano
    df_cierre1.rename({'Documento': 'DNI', 
                       'ImportePago': 'IMPORTE', 
                       }, axis=1, inplace=True)
    #'Fecha': 'Fecha de Pago'
    df_cierre1['DNI'] = df_cierre1['DNI'].apply(int)
    df_cierre = df_cierre1.copy()
    return df_cierre