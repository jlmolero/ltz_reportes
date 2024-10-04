import os
import openpyxl
import re
import pandas as pd
import yaml
from basicas import latino_a_numero

######################
#
# Cuentas
#
######################
with open('./config.yaml', 'r') as f:
    config = yaml.safe_load(f)

cuentas = {}
for cuenta, valores in config['cuentas'].items():
    cuentas[cuenta] = {
        'ruta': valores['ruta'],
        'banco': valores['banco']
    }

######################
# FUNCION PARA OBTENER LOS MOVIMIENTOS DE UNA CUENTA
######################
def get_movimientos(cuenta, meses_interes=None):
    ruta = cuentas[cuenta]['ruta']
    banco = cuentas[cuenta]['banco']
    movimientos = pd.DataFrame()
    for file in os.listdir(ruta):
        if file.endswith(".xlsx"):
            wb = openpyxl.load_workbook(os.path.join(ruta, file))
            for sheet in wb:
                if sheet.title != 'Resumen':
                    df = pd.read_excel(os.path.join(ruta, file), sheet_name=sheet.title)
                    movimientos = pd.concat([movimientos, df], ignore_index=True)
    #
    # PREPARAR MOVIEMIENTOS
    #
    if banco == 'BANCRECER':
        # convertir la columna 'FECHA' a tipo datetime
        movimientos['FECHA'] = pd.to_datetime(movimientos['FECHA'], dayfirst=True)

        #Seleccionar solo los movimientos correspondiente al periodo deseado
        if meses_interes is not None:
            movimientos = movimientos[movimientos['FECHA'].dt.month.isin(meses_interes)]

        # preparar la columna REFERENCIA
        movimientos['REFERENCIA'] = movimientos['REFERENCIA'].astype(str).str.zfill(8)

        #Crear La Columna LOTE cuyos valores serán los mismos de la columna REFERENCIA
        movimientos['lote'] = movimientos['REFERENCIA']

        #Crear la columna tipoOperacion
        movimientos['tipoOperacion'] = ''
        for index, row in movimientos.iterrows():
            if row['DESCRIPCION']=="N/D CREDITO INM.OB-G200076496":
                movimientos.at[index, 'tipoOperacion'] = 'TRASFERENCIA LTZ OTROS BANCOS'
            elif row['DESCRIPCION'].startswith('N/D CREDITO INM.'):
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS A PROVEEDORES'
            elif row['DESCRIPCION'].startswith('COM.CREDITO INM.OB'):
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES PAGOS A PROVEEDORES'
            elif row['DESCRIPCION'].startswith('N/C CREDITO INM.'):
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS RECIBIDOS'
            elif row['DESCRIPCION'].startswith('COMISION'):
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES VARIAS'
            elif row['DESCRIPCION']=="COM. POR EMISION EDO. DE CTA.":
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES VARIAS'
            elif row['DESCRIPCION']=="COM. POR MANTENIMIENTO DE CTA.":
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES VARIAS'


        #Convertir las columnas 'DEBITOS' y 'CREDITOS' a tipo float
        movimientos['DEBITOS'] = movimientos['DEBITOS'].apply(latino_a_numero)
        movimientos['CREDITOS'] = movimientos['CREDITOS'].apply(latino_a_numero)
        movimientos['SALDO'] = movimientos['SALDO'].apply(latino_a_numero)
        
        #Crear 2 nuevas columnas, una llamada monto y otra llama tipoMovimiento
        movimientos['monto'] = 0.0
        movimientos['tipoMovimiento'] = ''
        for index, row in movimientos.iterrows():
            if row['DEBITOS'] > 0 and row['CREDITOS'] == 0:
                movimientos.loc[index, 'monto'] = -row['DEBITOS']
                movimientos.loc[index, 'tipoMovimiento'] = 'Nota de Débito'
            elif row['DEBITOS'] == 0 and row['CREDITOS'] > 0:
                movimientos.loc[index, 'monto'] = row['CREDITOS']
                movimientos.loc[index, 'tipoMovimiento'] = 'Nota de Crédito'
        
        #Eliminar las columnas 'DEBITOS' y 'CREDITOS'
        movimientos = movimientos.drop(['DEBITOS', 'CREDITOS'], axis=1)
        
        #Cambiar el nombre de las columnas
        movimientos = movimientos.rename(columns={'FECHA': 'fecha', 'SALDO': 'saldo', 'REFERENCIA': 'referencia', 'DESCRIPCION': 'concepto'})

        #reordenar las columnas
        movimientos = movimientos[['fecha', 'referencia', 'lote', 'concepto', 'monto', 'saldo', 'tipoOperacion']]

    elif banco == 'BDV':
        # elmiar las columnas 'rif' y 'numeroCuenta'
        movimientos = movimientos.drop(['rif', 'numeroCuenta'], axis=1)

        # Eliminar todos los registros cuyo valor para la columnas 'concepto' sea "SALDO INICIAL"
        movimientos = movimientos[movimientos['tipoMovimiento'] != 'Saldo Inicial']

        # convertir la columna 'fecha' a tipo datetime
        movimientos['fecha'] = pd.to_datetime(movimientos['fecha'], dayfirst=True)

        #Seleccionar solo los movimientos correspondiente al periodo deseado
        if meses_interes is not None:
            movimientos = movimientos[movimientos['fecha'].dt.month.isin(meses_interes)]
        
        #Convertir las columnas de 'monto' y saldo' a tipo float
        movimientos['monto'] = movimientos['monto'].apply(latino_a_numero)
        movimientos['saldo'] = movimientos['saldo'].apply(latino_a_numero)

        #Extraer los ultimos 8 digitos de la columna 'referencia' y convertirlos a tipo str
        movimientos['referencia'] = movimientos['referencia'].astype(str).apply(lambda x: x[:-1] if x.endswith(' ') else x)
        movimientos['referencia'] = movimientos['referencia'].str[-8:].str.zfill(8)

        #Crear la columna 'lote'
        movimientos['lote'] = ''
        movimientos['lote'] = movimientos.apply(lambda row: row['concepto'][-8:] if 'LOTE' in row['concepto'] else row['referencia'], axis=1)

        #crear la columna tipoOperacion
        movimientos['tipoOperacion'] = ''
        for index, row in movimientos.iterrows():
            if row['concepto']=="PAGO RECIBIDO OTROS BANCOS 0168 G200076496": #Transferencias internas desde Bancrecer
                movimientos.at[index, 'tipoOperacion'] = 'TRANSFERENCIA RECIBIDA LTZ BANCRECER'
            elif row['concepto'].startswith("PAGO RECIBIDO BDV G200076496"): #Transferencias internas desde BDV
                movimientos.at[index, 'tipoOperacion'] = 'TRANSFERENCIA RECIBIDA LTZ BDV'
            elif row['concepto'].startswith('COMISION PAGO A PROVEEDORES'): # Comisiones de Pagos a Proveedores
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES PAGOS A PROVEEDORES'
            elif row['concepto'].startswith('PAGO A PROVEEDORES'): # Pagos a Proveedores
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS A PROVEEDORES'
            elif row['concepto'].startswith('COM MANTENIMIENTO DE CUENTA'): # Comisiones de Mantenimiento de cuenta
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES VARIAS'
            elif row['concepto'].startswith("PAGO IMPUESTOS INTERNET"): # Pagos de Impuestos
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS AL SENIAT'
            elif row['concepto'].startswith('PAGO BANAVIH'): # Pagos a BANAVIH
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS BANAVIH'
            elif row['concepto'].startswith('DEVUELTA PAGO PROVEEDORES'): # Reintegros
                movimientos.at[index, 'tipoOperacion'] = 'REINTEGROS TRANSFERENCIAS FALLIDAS'
            elif row['concepto'].startswith('TRANSF PROPIA BDV G200076496'): #Transferencias internas emitida a BDV
                movimientos.at[index, 'tipoOperacion'] = 'TRASFERENCIA INTERNA LTZ EMITIDA BDV'
            elif row['concepto'].startswith('PAGO RECIBIDO'): # Pagos de otros bancos
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS RECIBIDOS OTROS BANCOS'
            elif row['concepto'].startswith('TRANSF RECIBIDA BDV'): #Pagos de BDV
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS RECIBIDOS BDV'
            elif row['concepto'].startswith('PAGO PROVEEDOR CTAS PROPIAS'): #Transferencias emitidas internas bdv
                movimientos.at[index, 'tipoOperacion'] = 'TRANSFERENCIAS EMITIDA LTZ BDV'
            elif row['concepto'].startswith('COM PAGO A PROVEED CTAS PROPIA'): #Comision Transferencias emitidas internas bdv
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES TRANSFERENCIA EMITIDA LTZ BDV'
            elif row['concepto'].startswith('PAGO IVSS'): #PAgo seguro social
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS SEGURO SOCIAL'
            elif row['concepto'].startswith('DOMICILIACION'): #Pagos Domiciliados
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS DOMICILIADOS'
            elif row['concepto'].startswith('ABONO INTERESES LIQUIDACION'): #Abono intereses de liquidacion
                movimientos.at[index, 'tipoOperacion'] = 'ABONOS DE INTERESES'
            elif row['concepto'].startswith('PAGOMOVIL'): #Pago Movil
                movimientos.at[index, 'tipoOperacion'] = 'PAGOS MOVIL'
            elif row['concepto'].startswith('COBRO COMISION PAG MOVIL'): #Pago Movil
                movimientos.at[index, 'tipoOperacion'] = 'COMISIONES VARIAS'
            elif row['tipoMovimiento']=='Deposito':
                movimientos.at[index, 'tipoOperacion'] = 'DEPOSITOS RECIBIDOS'
            

        #reordenar las columnas
        movimientos = movimientos[['fecha', 'referencia', 'lote', 'concepto', 'monto', 'saldo', 'tipoOperacion']]


    return movimientos

######################
# FUNCION PARA OBTENER UN RESUMEN DE UNA CUENTA
######################
def get_resumen_cuenta(cuenta, meses_interes=None):
    movimientos = get_movimientos(cuenta,meses_interes)
    resumen_cuenta = movimientos.groupby(['tipoOperacion'])['monto'].sum()
    return resumen_cuenta

mov = get_movimientos('0171',[9])
#print(mov.to_string(index=False))
print(mov)
#print(mov[mov['concepto'].str.contains('G200076496')])
#print(mov[mov['tipoOperacion'] == ''])
#print(mov[mov['tipoOperacion'].str.startswith('TRANSFERENCIA RECIBIDA')])
resumen= get_resumen_cuenta('4363')
print(resumen)