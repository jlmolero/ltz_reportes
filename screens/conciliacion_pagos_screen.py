import PySimpleGUI as sg
import pandas as pd
from common.cuentas import get_movimientos
from common.ejecucion import gastos_periodo

def conciliacion_pagos(meses_seleccionados):
    cuenta = '6597'
    movimientos=get_movimientos(cuenta,meses_seleccionados)
    gastos=gastos_periodo(meses_seleccionados)
    gastos_deseados=gastos[gastos['Cuenta']==cuenta]
    print(movimientos)
    #print(gastos)
    print(gastos_deseados)

    layout=[
        [sg.Text("Conciliación de Pagos", font=("Helvetica", 16, "bold"))],
        [sg.Text("Informes Para el Año 2024", font=("Helvetica", 12))],
        [sg.Text("Conciliación de Pagos", font=("Helvetica", 12))],
    ]
    
    window=sg.Window("Conciliación de Pagos", layout)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
    window.close()