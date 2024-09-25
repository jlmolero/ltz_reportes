#import PySimpleGUI as sg
import PySimpleGUI as sg
from common.ejecucion import gastos_periodo, informe_ejecucion_cuenta, periodo_meses
from common.basicas import latinizar
import pandas as pd





def ejecucion_cuenta(meses_seleccionados):

    periodo=f"{periodo_meses(meses_seleccionados)} 2024"
    pagos = gastos_periodo(meses_seleccionados)
    informe = informe_ejecucion_cuenta(pagos, meses_seleccionados)
    
    total_0171= informe[informe['CodigoPartida'].str.len() == 4]['0171'].sum()
    total_2633 = informe[informe['CodigoPartida'].str.len() == 4]['2633'].sum()
    total_4363 = informe[informe['CodigoPartida'].str.len() == 4]['4363'].sum()
    total_6597 = informe[informe['CodigoPartida'].str.len() == 4]['6597'].sum()
    total_PATRIA = informe[informe['CodigoPartida'].str.len() == 4]['PATRIA'].sum()
    total_general = informe[informe['CodigoPartida'].str.len() == 4]['Total'].sum()
    
    lista_informe = informe.values.tolist()
    #Mostrar los numero con el formato latino
    for i in range(len(lista_informe)):
        lista_informe[i][2] = latinizar(lista_informe[i][2])
        lista_informe[i][3] = latinizar(lista_informe[i][3])
        lista_informe[i][4] = latinizar(lista_informe[i][4])
        lista_informe[i][5] = latinizar(lista_informe[i][5])
        lista_informe[i][6] = latinizar(lista_informe[i][6])
        lista_informe[i][7] = latinizar(lista_informe[i][7])

    #Hacer que las filas de subtotales aparezcan con un color diferente
    colores_filas=[]
    for i, row in enumerate(lista_informe):
        if len(row[0]) == 4:            
            colores_filas.append((i, "black", "#8d7a64"))
    

    tabla_pagos=[sg.Table(values=lista_informe, headings=["Código de Partida", "Descripción de Partida", "0171", "2633", "4363", "6597", "PATRIA", "Total"],
                         col_widths=[15, 50, 8, 8, 8, 8, 8, 8], auto_size_columns=False, justification="center", num_rows=24, font=("Helvetica", 12), key="-TABLA-",
                         row_colors=colores_filas)
                         ]
    
    
    fila_totales = [sg.Push(),
                    sg.Text(latinizar(total_0171), font=("Helvetica", 12, "bold"), size=(8,1)),
                    sg.Text(latinizar(total_2633), font=("Helvetica", 12, "bold"), size=(8,1)),
                    sg.Text(latinizar(total_4363), font=("Helvetica", 12, "bold"), size=(8,1)),
                    sg.Text(latinizar(total_6597), font=("Helvetica", 12, "bold"), size=(8,1)),
                    sg.Text(latinizar(total_PATRIA), font=("Helvetica", 12, "bold"), size=(8,1)),
                    sg.Text(latinizar(total_general), font=("Helvetica", 12, "bold"),size=(10,1))]    

    layout = [
        [sg.Text("Ejecución de Gasto por Cuenta Bancaria", font=("Helvetica", 16, "bold"))],
        [sg.Text("Periodo:", font=("Helvetica", 14)), sg.Text(periodo, font=("Helvetica", 14, "bold"))],
        tabla_pagos,
        fila_totales,
        [sg.Push(),sg.Button("Exportar a Excel", key="-EXCEL-", font=("Helvetica", 12))]
        
        ]

    
    
    window=sg.Window("Ejecución de Gasto por Cuenta Bancaria", layout, resizable=True)



    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        
        elif event == "-EXCEL-":
            ruta=sg.popup_get_folder("Seleccionar Destino")
            print(ruta)

    window.close()