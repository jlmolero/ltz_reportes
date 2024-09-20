import pandas as pd
import datetime as dt
import openpyxl

ruta_pagos='./data/RegistroPagos.xlsx'

# Define un diccionario que mapea los valores de la columna Cuenta a los valores de la columna FuenteDeFinanciamiento
cuenta_fuente_financiamiento = {
    '4363': 'Recursos por Operaciones',
    '6597': 'Recursos por Operaciones',
    '0171': 'Situado Constitucional',
    'PATRIA': 'Situado Constitucional'
}

partida_descripcion = {
    '4.01':'GASTOS DE PERSONAL',
    '4.02':'MATERIALES, SUMINISTROS Y MERCANCÍAS',
    '4.03':'SERVICIOS NO PERSONALES',
    '4.04':'ACTIVOS REALES',
    '4.05':'ACTIVOS  FINANCIEROS',
    '4.06':'GASTOS DE DEFENSA Y SEGURIDAD DEL ESTADO',
    '4.07':'TRANSFERENCIAS Y DONACIONES',
    '4.08':'OTROS GASTOS',
    '4.09':'ASIGNACIONES NO DISTRIBUIDAS',
    '4.10':'SERVICIO DE LA DEUDA PÚBLICA',
    '4.11':'DISMINUCIÓN DE PASIVOS',
    '4.12':'DISMINUCIÓN DEL PATRIMONIO',
    '4.98':'RECTIFICACIONES AL PRESUPUESTO'
}

meses={
        'enero': 1,
        'febrero': 2,
        'marzo': 3,
        'abril': 4,
        'mayo': 5,
        'junio': 6,
        'julio': 7,
        'agosto': 8,
        'septiembre': 9,
        'octubre': 10,
        'noviembre': 11,
        'diciembre': 12
    }

def gastos_periodo(meses_interes):
    # ejecucion_gasto.py


    meses={
        'enero': 1,
        'febrero': 2,
        'marzo': 3,
        'abril': 4,
        'mayo': 5,
        'junio': 6,
        'julio': 7,
        'agosto': 8,
        'septiembre': 9,
        'octubre': 10,
        'noviembre': 11,
        'diciembre': 12
    }

    if meses_interes == [1, 2, 3]:
        periodo = 'Primer Trimestre'
    elif meses_interes == [4, 5, 6]:
        periodo = 'Segundo Trimestre'
    elif meses_interes == [7, 8, 9]:
        periodo = 'Tercer Trimestre'
    elif meses_interes == [10, 11, 12]:
        periodo = 'Cuarto Trimestre'
    else:
        periodo = meses_interes

    datos = pd.read_excel(ruta_pagos, sheet_name="Pagos") # Importar datos
    
    #Eliminar registros
    datos= datos[datos['NroFactura']!='ANULADA'] # Eliminar ordenes de pago anuladas
    datos=datos[datos['OrdenPago'].notnull()] # Eliminar ordenes de pago vacías
    datos = datos[datos['CodigoPartida']!='4.03.18.01.00'] # Eliminar registros de ordemes de pagos de retencion de IVA
    datos = datos[datos['CodigoPartida']!='4.03.18.99.00'] # Eliminar registros de ordemes de pagos de retencion de islr y sedatez


    # Seleccionar solo los registros que corresponden al periodo deseado
    datos = datos[datos['Fecha'].dt.month.isin(meses_interes)]

    #crear una columna que contendra las diferentes fuentes de financimiento segun la cuenta de donde sale los recursos
    fuente_financiamiento = datos['Cuenta'].map(cuenta_fuente_financiamiento)
    #Asignar la fuente de financiamiento a las columnas donde ya venga especificada
    datos = datos.assign(FuenteDeFinanciamiento=datos['FuenteDeFinanciamiento'].fillna(fuente_financiamiento))
    print(datos)

    #Seleccionar las columnas deseadas para la tabla de datos
    datos=datos[['N', 'Fecha', 'Cuenta', 'OrdenPago', 'DocBeneficiario', 'Beneficiario', 'Descripcion', 'MontoSinIVA', 'IVA', 'FuenteDeFinanciamiento']]
    
    
    return datos


