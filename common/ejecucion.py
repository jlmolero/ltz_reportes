import pandas as pd
import datetime as dt
import openpyxl
import yaml

with open('./config.yaml', 'r') as f:
    config = yaml.safe_load(f)



#ruta_pagos='./data/RegistroPagos.xlsx'
ruta_pagos = config['origen_pagos']
#ruta_pagos = r"\\dc2\administracion\2024\CUADROS\Registro de Pagos.xlsx"

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

def periodo_meses(meses_interes):
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
    elif meses_interes == [1, 2,3,4,5,6,7,8,9,10,11,12]:
        periodo = 'Año Completo' 
    else:
        if len(meses_interes) == 1:
            periodo=list(meses.keys())[list(meses.values()).index(meses_interes[0])]
        else:
            periodo=''
            for mes, num in meses.items():
                if num in meses_interes:
                    periodo=periodo + mes + ', '
            periodo=periodo[:-2]

    return(periodo)

def gastos_periodo(meses_interes):
    # ejecucion_gasto.py

    """
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
    """
    datos = pd.read_excel(ruta_pagos, sheet_name="Pagos") # Importar datos
    
    #Eliminar registros
    datos= datos[datos['NroFactura']!='ANULADA'] # Eliminar ordenes de pago anuladas
    datos=datos[datos['OrdenPago'].notnull()] # Eliminar ordenes de pago vacías
    datos = datos[datos['CodigoPartida']!='4.03.18.01.00'] # Eliminar registros de ordemes de pagos de retencion de IVA
    datos = datos[datos['CodigoPartida']!='4.03.18.99.00'] # Eliminar registros de ordemes de pagos de retencion de islr y sedatez


    # Seleccionar solo los registros que corresponden al periodo deseado
    datos = datos[datos['Fecha'].dt.month.isin(meses_interes)]

    #Sustituir valores - en la columna IVA
    #datos['IVA'] = datos['IVA'].replace('-', 0)
    datos['IVA'] = datos['IVA'].apply(lambda x: 0 if x == '-' else x)


    #crear una columna que contendra las diferentes fuentes de financimiento segun la cuenta de donde sale los recursos
    fuente_financiamiento = datos['Cuenta'].map(cuenta_fuente_financiamiento)
    #Asignar la fuente de financiamiento a las columnas donde ya venga especificada
    datos = datos.assign(FuenteDeFinanciamiento=datos['FuenteDeFinanciamiento'].fillna(fuente_financiamiento))
    

    #Seleccionar las columnas deseadas para la tabla de datos
    datos=datos[['Fecha', 'Cuenta', 'OrdenPago', 'Beneficiario', 'CodigoPartida', 'DescripcionPartida', 'Descripcion', 'MontoSinIVA', 'IVA', 'FuenteDeFinanciamiento']]
    
    
    return datos

def informe_ejecucion_gasto(datos, meses_interes):
    iva_periodo= datos[datos['Fecha'].dt.month.isin(meses_interes)].pivot_table(values='IVA', index=None, columns='FuenteDeFinanciamiento', aggfunc='sum').infer_objects(copy=False).fillna(0).reset_index()
    #Reindexamos las columnas para garantizar que tenga valores en todas las fuentes de financiamiento
    iva_periodo = iva_periodo.reindex(columns=['Recursos por Operaciones', 'Situado Constitucional'], fill_value=0)

    fila_iva_mes = pd.DataFrame({'CodigoPartida':['4.03.18.01.00'], 'DescripcionPartida':['Impuesto al valor Agregado'], 'Recursos por Operaciones':[iva_periodo.loc[0, 'Recursos por Operaciones']], 'Situado Constitucional':[iva_periodo.loc[0, 'Situado Constitucional']]}).fillna(0)

    informe_gasto_ejecutado = datos.groupby(['CodigoPartida', 'DescripcionPartida', 'FuenteDeFinanciamiento'])['MontoSinIVA'].sum().unstack('FuenteDeFinanciamiento').fillna(0).rename_axis(None, axis=1).reset_index()
    # Reindexamos las columnas para garantizar que tenga valores en todas las fuentes de financiamiento
    informe_gasto_ejecutado = informe_gasto_ejecutado.reindex(columns=['CodigoPartida', 'DescripcionPartida', 'Recursos por Operaciones', 'Situado Constitucional'], fill_value=0)

    informe_gasto_ejecutado = pd.concat([informe_gasto_ejecutado, fila_iva_mes], ignore_index=True).sort_values(by='CodigoPartida')

    # Ahora voy a agregar una columna con el valor total de 'MontoSinIVA' para cada agrupación
    informe_gasto_ejecutado['Total'] = informe_gasto_ejecutado.iloc[:, 2:].sum(axis=1)

    #Calcular subtotales    
    informe_gasto_ejecutado_subtotales = informe_gasto_ejecutado.groupby(informe_gasto_ejecutado['CodigoPartida'].str[:4]).agg({'Recursos por Operaciones':'sum', 'Situado Constitucional':'sum', 'Total': 'sum'}).rename_axis('CodigoPartida').reset_index()
    informe_gasto_ejecutado_subtotales['CodigoPartida'] = informe_gasto_ejecutado_subtotales['CodigoPartida']# + '.00.00.00'
    informe_gasto_ejecutado_subtotales['DescripcionPartida'] = informe_gasto_ejecutado_subtotales['CodigoPartida'].map(partida_descripcion)

    informe_gasto_ejecutado = pd.concat([informe_gasto_ejecutado, informe_gasto_ejecutado_subtotales], ignore_index=True).sort_values(by='CodigoPartida')

 
    #return informe_gasto_ejecutado
    return informe_gasto_ejecutado

def informe_ejecucion_cuenta(datos, meses_interes):
    
    iva_periodo= datos[datos['Fecha'].dt.month.isin(meses_interes)].pivot_table(values='IVA', index=None, columns='Cuenta', aggfunc='sum').infer_objects(copy=False).fillna(0).reset_index()
    # Reindexamos el dataframe para garantizar que tenga valores para todas las cuentas
    iva_periodo = iva_periodo.reindex(columns=['0171', '2633', '4363', '6597', 'PATRIA'], fill_value=0.0)

    
    fila_iva_mes = pd.DataFrame({'CodigoPartida':['4.03.18.01.00'], 'DescripcionPartida':['Impuesto al valor Agregado'],
                                 '0171':[iva_periodo.loc[0, '0171']],
                                 '2633':[iva_periodo.loc[0, '2633']],
                                 '4363':[iva_periodo.loc[0,'4363']],
                                 '6597':[iva_periodo.loc[0,'6597']],
                                 'PATRIA':[iva_periodo.loc[0,'PATRIA']]
                                 }).fillna(0)

    
    informe_gasto_ejecutado = datos.groupby(['CodigoPartida', 'DescripcionPartida', 'Cuenta'])['MontoSinIVA'].sum().unstack('Cuenta').fillna(0).rename_axis(None, axis=1).reset_index()
    # Reindexamos el dataframe para garantizar que tenga valores para todas las Cuentas
    informe_gasto_ejecutado = informe_gasto_ejecutado.reindex(columns=['CodigoPartida', 'DescripcionPartida', '0171', '2633', '4363', '6597', 'PATRIA'], fill_value=0.0)
    informe_gasto_ejecutado = pd.concat([informe_gasto_ejecutado, fila_iva_mes], ignore_index=True).sort_values(by='CodigoPartida')
    
    # Ahora voy a agregar una columna con el valor total de 'MontoSinIVA' para cada agrupación
    informe_gasto_ejecutado['Total'] = informe_gasto_ejecutado.iloc[:, 2:].sum(axis=1)

    #Calcular subtotales    
    informe_gasto_ejecutado_subtotales = informe_gasto_ejecutado.groupby(informe_gasto_ejecutado['CodigoPartida'].str[:4]).agg({'0171':'sum', '2633':'sum', '4363':'sum', '6597':'sum', 'PATRIA':'sum', 'Total': 'sum'}).rename_axis('CodigoPartida').reset_index()
    
    #informe_gasto_ejecutado_subtotales['CodigoPartida'] = informe_gasto_ejecutado_subtotales['CodigoPartida']# + '.00.00.00'
    informe_gasto_ejecutado_subtotales['DescripcionPartida'] = informe_gasto_ejecutado_subtotales['CodigoPartida'].map(partida_descripcion)
    informe_gasto_ejecutado = pd.concat([informe_gasto_ejecutado, informe_gasto_ejecutado_subtotales], ignore_index=True).sort_values(by='CodigoPartida')

    
    return informe_gasto_ejecutado

