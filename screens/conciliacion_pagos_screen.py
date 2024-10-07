import PySimpleGUI as sg
import pandas as pd
from common.cuentas import get_movimientos, exportar_conciliacion
from common.ejecucion import gastos_periodo
from common.transferencias import get_transferencias_periodo
import os

def conciliacion_pagos(meses_seleccionados, cuenta):
    
    #Carcagar movimientos bancarios de la cuenta
    movimientos=get_movimientos(cuenta,meses_seleccionados)
    suma_comisiones=movimientos[movimientos['tipoOperacion']=='COMISIONES PAGOS A PROVEEDORES']['monto'].sum()
    #movimientos=movimientos[movimientos['tipoOperacion']!='COMISIONES PAGOS A PROVEEDORES']
    movimientos=movimientos[['fecha','concepto','referencia','lote','monto','tipoOperacion']]
    movimientos.reset_index(drop=True, inplace=True)
    #Cargar transferencias internas de la cuenta
    transferencias=get_transferencias_periodo(meses_seleccionados)
    transferencias=transferencias[transferencias['CuentaOrigen']==cuenta]
    transferencias=transferencias[['Fecha','CodTransf','CuentaDestino','Descripcion','Referencia','Monto','OP_Relacionada','RazonTransferencia']]
    #cargar pagos de la cuenta
    gastos=gastos_periodo(meses_seleccionados)
    gastos=gastos[gastos['Cuenta']==cuenta]
    gastos=gastos[['Fecha','OrdenPago','Referencia','Beneficiario','MontoPagado','totalRetenido']]
    

    pagos = pd.DataFrame()
    recibidos = pd.DataFrame()

    #pago={}
    #print(movimientos)
    #########################
    # CREAR TABLA DE PAGOS RECIBIDOS POR LA CUENTA
    ########################
    
    for index, movimiento in movimientos.iterrows():
        if movimiento['monto']>0:
            recibido={}
            recibido['fecha']=movimiento['fecha']
            recibido['concepto']=movimiento['concepto']
            recibido['referencia']=movimiento['referencia']
            recibido['monto']=movimiento['monto']
            recibido['tipoOperacion']=movimiento['tipoOperacion']
            recibidos = pd.concat([recibidos, pd.DataFrame(recibido, index=[0])], ignore_index=True)
            movimientos.drop(index=movimiento.name, inplace=True)

    #########################
    # CREAR TABLA DE PAGOS QUE MATCHEAN ENTRE MOVIMIENTOS Y GASTOS
    ########################

        elif movimiento['tipoOperacion']=='PAGOS A PROVEEDORES':
            pago={}
            pago['fecha']=movimiento['fecha']
            referencia_match = gastos[gastos['Referencia']==movimiento['lote']]
            #print(referencia_match)
            if not referencia_match.empty:
                pago['ordenPago']=referencia_match['OrdenPago'].iloc[0]
                pago['beneficiario']=referencia_match['Beneficiario'].iloc[0]
                pago['montoOrdenPago']=referencia_match['MontoPagado'].iloc[0]
                pago['lote']=movimiento['lote']
                pago['referencia']=movimiento['referencia']
                pago['montoOperacion']=movimiento['monto']

                pagos = pd.concat([pagos, pd.DataFrame(pago, index=[0])], ignore_index=True)
                movimientos.drop(index=movimiento.name, inplace=True)
                gastos.drop(referencia_match.index[0], inplace=True)
                

    #########################
    # CREAR TABLA DE COMISIONES QUE MATCHEAN ENTRE PAGOS Y MOVIMIENTOS
    # #########################
      
    comisiones = pd.DataFrame()

    for index, pago in pagos.iterrows():
        comision = {}
        mov_match = movimientos[movimientos['lote']==pago['referencia']]
        if not mov_match.empty:
            #pago['comision']=mov_match['monto'].iloc[0]
            #pago['comisionRef']=mov_match['referencia'].iloc[0]
            comision['fecha']=pago['fecha']
            comision['ordenPago']=pago['ordenPago']
            comision['referencia']=pago['referencia']
            comision['montoComision']=mov_match['monto'].iloc[0]

            comisiones = pd.concat([comisiones, pd.DataFrame(comision, index=[0])], ignore_index=True)
            
            movimientos.drop(index=mov_match.index[0], inplace=True)

    
    
    ############################
    # CREAR TABLA de RETENCIONES QUE MATCHEAN ENTRE MOVIMIENTOS Y TRANSFERENCIAS
    ########################
    retenciones = pd.DataFrame()
    for index, movimiento in movimientos.iterrows():
        retencion = {}
        transf_mov_match = transferencias[transferencias['Referencia']==movimiento['lote']]
        
        razonTransferencia = "\n".join(transf_mov_match['RazonTransferencia'].tolist())
        

        if not transf_mov_match.empty and "Retención" in razonTransferencia:
        #if not transf_mov_match.empty:
            retencion['fecha']=movimiento['fecha']
            retencion['ordenPago']=transf_mov_match['OP_Relacionada'].iloc[0]
            retencion['descripcion']=transf_mov_match['Descripcion'].iloc[0]
            retencion['referencia']=movimiento['lote']
            retencion['montoRetencion']=transf_mov_match['Monto'].iloc[0]
            retencion['montoOperacion']=movimiento['monto']

            retenciones = pd.concat([retenciones, pd.DataFrame(retencion, index=[0])], ignore_index=True)

            movimientos.drop(index=movimiento.name, inplace=True)
            transferencias.drop(index=transf_mov_match.index[0], inplace=True)
            #transferencias.drop(index=transf_mov_match.name, inplace=True)

    #########################
    # CREAR TABLA DE COMISIONES POR TRASNFERENCIA DE RETENCIONES
    ########################
    comisiones_ret=pd.DataFrame()

    for index, retencion in retenciones.iterrows():
        comision_ret={}
        ret_match = movimientos[movimientos['referencia']==retencion['referencia']]
        

        if not ret_match.empty:
            comision_ret['fecha']=ret_match['fecha'].iloc[0] 
            comision_ret['ordenPago']=retencion['ordenPago']
            comision_ret['referencia']=ret_match['referencia'].iloc[0]
            comision_ret['montoComision']=ret_match['monto'].iloc[0]

            comisiones_ret = pd.concat([comisiones_ret, pd.DataFrame(comision_ret, index=[0])], ignore_index=True)
            
            movimientos.drop(index=ret_match.index[0], inplace=True)

    #########################
    # CREAR UN DATAFRAME CONSOLIDADO DE PAGOS
    ########################
    consolidado_pagos = pd.DataFrame()
    for index, pago in pagos.iterrows():
        comision_match = comisiones[comisiones['ordenPago']==pago['ordenPago']]
        retencion_match = retenciones[retenciones['ordenPago']==pago['ordenPago']]
        if not comisiones_ret.empty:
            comisiones_ret_match = comisiones_ret[comisiones_ret['ordenPago']==pago['ordenPago']]
        else:
            comisiones_ret_match = pd.DataFrame()


        consolidado_pago={}
        consolidado_pago['fecha']=pago['fecha']
        consolidado_pago['ordenPago']=pago['ordenPago']
        consolidado_pago['beneficiario']=pago['beneficiario']
        consolidado_pago['montoOP']=pago['montoOrdenPago']
        consolidado_pago['pagado']=pago['montoOperacion']
        consolidado_pago['lote']=pago['lote']
        if not comision_match.empty:
            consolidado_pago['comision']=comision_match['montoComision'].iloc[0]
            consolidado_pago['comisionRef']=comision_match['referencia'].iloc[0]
        else:
            consolidado_pago['comision']=0
            consolidado_pago['comisionRef']=''
        if not retencion_match.empty:
            consolidado_pago['retCalculada']=retencion_match['montoRetencion'].iloc[0]
            consolidado_pago['retPagada']=retencion_match['montoOperacion'].iloc[0]
            consolidado_pago['retRef']=retencion_match['referencia'].iloc[0]
        else:
            consolidado_pago['retCalculada']=0
            consolidado_pago['retPagada']=0
            consolidado_pago['retRef']=''
        if not comisiones_ret_match.empty:
            consolidado_pago['retComision']=comisiones_ret_match['montoComision'].iloc[0]
        else:
            consolidado_pago['retComision']=0

        




        consolidado_pagos = pd.concat([consolidado_pagos, pd.DataFrame(consolidado_pago, index=[0])], ignore_index=True)


    os.system('cls' if os.name == 'nt' else 'clear')
    """
    print('Pagos Recibidos:')
    print(recibidos)
    print('Pagos Identificados:')
    print(pagos)
    print('Comisiones Identificadas:')
    print(comisiones)
    print('Retenciones Identificadas:')
    print(retenciones)
    print('Comisiones por Retenciones:')
    print(comisiones_ret)
    #print(movimientos)
    print('Gastos Huerfanos:')
    print(gastos)
    print('Transferencias Huerfanas:')
    print(transferencias)
    print('Movimientos Huerfanos:')
    print(movimientos)
    print('Consolidado de Pagos:')
    print(consolidado_pagos)
    #print(suma_comisiones)
    #print(gastos)
    
    #print(transferencias)
    """
    layout=[
        [sg.Text("Conciliación de Cuenta:", font=("Helvetica", 16, "bold")),
         sg.Text(cuenta, font=("Helvetica", 16, "bold"))],
        [sg.Text("Informes Para el Año 2024", font=("Helvetica", 12))],
        [sg.Text("Conciliación de Pagos", font=("Helvetica", 12))],
        [sg.Button("Regresar", key="-REGRESAR-", font=("Helvetica", 12)), sg.Push(),
        sg.Button("Exportar a Excel", key="-EXPORTAR-", font=("Helvetica", 12))],
    ]
    
    window=sg.Window("Conciliación de Pagos", layout)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == "-REGRESAR-":
            from screens.mainscreen import mainscreen
            window.close()
            mainscreen()
        elif event == "-EXPORTAR-":
            destino=sg.popup_get_folder(message="Seleccionar carpeta para exportar", title="Carpeta de Destino")
            exportar_conciliacion(cuenta, meses_seleccionados, destino, consolidado_pagos, movimientos, gastos, transferencias)

    window.close()