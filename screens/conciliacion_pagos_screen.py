import PySimpleGUI as sg
import pandas as pd
from common.cuentas import get_movimientos
from common.ejecucion import gastos_periodo

def conciliacion_pagos(meses_seleccionados, cuenta):
    
    movimientos=get_movimientos(cuenta,meses_seleccionados)
    suma_comisiones=movimientos[movimientos['tipoOperacion']=='COMISIONES PAGOS A PROVEEDORES']['monto'].sum()
    movimientos=movimientos[movimientos['tipoOperacion']!='COMISIONES PAGOS A PROVEEDORES']




    gastos=gastos_periodo(meses_seleccionados)
    gastos_deseados=gastos[gastos['Cuenta']==cuenta]
    
    gastos_deseados=gastos_deseados[['Fecha','OrdenPago','Referencia','Beneficiario','MontoPagado','totalRetenido']]
    print(movimientos)
    print(suma_comisiones)
    #print(gastos)
    print(gastos_deseados)

    layout=[
        [sg.Text("Conciliación de Pagos", font=("Helvetica", 16, "bold"))],
        [sg.Text("Informes Para el Año 2024", font=("Helvetica", 12))],
        [sg.Text("Conciliación de Pagos", font=("Helvetica", 12))],
        [sg.Button("Regresar", key="-REGRESAR-", font=("Helvetica", 12))],
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
            return
    window.close()