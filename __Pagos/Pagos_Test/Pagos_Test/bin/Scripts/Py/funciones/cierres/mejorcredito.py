import pandas as pd 

def cierre_mejorcredito(empresa, cierre):
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

    return df_cierre