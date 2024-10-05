import pandas as pd
import datetime as dt
import openpyxl, yaml, re, sys
from os import path

if sys.platform == 'Windows':
    with open('./config.yaml', 'r') as f:
        config = yaml.safe_load(f)
elif sys.platform == 'linux':
    with open('./config_linux.yaml', 'r') as f:
        config = yaml.safe_load(f)



#ruta_pagos='./data/RegistroPagos.xlsx'
ruta_transferencias = config['origen_pagos']

def get_transferencias_periodo(meses_interes):
    transferencias = pd.read_excel(ruta_transferencias, sheet_name='Transferencias')
    #Seleccionar solo los movimientos correspondiente al periodo deseado
    transferencias=transferencias[transferencias['Fecha'].dt.month.isin(meses_interes)]

    # preparar la columna REFERENCIA
    transferencias['Referencia'] = transferencias['Referencia'].astype(int).astype(str).str.zfill(8)

    #Preparar las columnas de cuentas
    transferencias['CuentaOrigen'] = transferencias['CuentaOrigen'].astype(int).astype(str).str.zfill(4)
    transferencias['CuentaDestino'] = transferencias['CuentaDestino'].astype(int).astype(str).str.zfill(4)

    return transferencias


