#import PySimpleGUI as sg
import PySimpleGUI as sg
from common.ejecucion import gastos_periodo, informe_ejecucion_gasto
import pandas as pd





def mainscreen():

    tabla_pagos=[sg.Table(values=[], headings=["Código de Partida", "Descripción de Partida", "Recursos por Operaciones", "Situado Constitucional", "Total"],
                         col_widths=[15, 50, 12, 12, 12], auto_size_columns=False, justification="center", num_rows=12, font=("Helvetica", 12), key="-TABLA-")
                         ]

    layout = [
        [sg.Text("Prueba")],
        [sg.Combo(["Año Completo","Primer Trimestre", "Segundo Trimestre", "Tercer Trimestre", "Cuarto Trimestre"],
                  default_value="", key="-TRIMESTRES-", readonly=True, font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Ene", key="-1-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Feb", key="-2-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Mar", key="-3-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Abr", key="-4-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("May", key="-5-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Jun", key="-6-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Jul", key="-7-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Ago", key="-8-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Sep", key="-9-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Oct", key="-10-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Nov", key="-11-", font=("Helvetica", 12), enable_events=True),
    sg.Checkbox("Dic", key="-12-", font=("Helvetica", 12), enable_events=True),
    sg.Push(), sg.Button("Consultar", key="-CONSULTAR-")
    ],
        tabla_pagos,
        [sg.Push(),sg.Button("Exportar a Excel", key="-EXCEL-", font=("Helvetica", 12))]
        
        ]

    
    
    window=sg.Window("Colección de Informes", layout, resizable=True)



    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == "-1-" or event == "-2-" or event == "-3-" or event == "-4-" or event == "-5-" or event == "-6-" or event == "-7-" or event == "-8-" or event == "-9-" or event == "-10-" or event == "-11-" or event == "-12-":
            meses_seleccionados = []
            window["-TRIMESTRES-"].update(value="")
            for i in range(1,12):
                if values[f"-{i}-"]:
                    meses_seleccionados.append(i)

        elif event == "-CONSULTAR-":
            pagos=gastos_periodo(meses_seleccionados)
            informe=informe_ejecucion_gasto(pagos, meses_seleccionados)
            lista_informe=informe.values.tolist()

            window["-TABLA-"].update(values=lista_informe)

        elif event == "-TRIMESTRES-":
            
            if values["-TRIMESTRES-"] == "Primer Trimestre":
                meses_seleccionados = [1, 2, 3]
                
                for i in range(1,4):
                    window[f"-{i}-"].update(value=True)
                for j in range(4,13):
                    window[f"-{j}-"].update(value=False)

            elif values["-TRIMESTRES-"] == "Segundo Trimestre":
                meses_seleccionados = [4, 5, 6]
                for i in range(4,7):
                    window[f"-{i}-"].update(value=True)
                for j in range(7,13):
                    window[f"-{j}-"].update(value=False)
                for k in range(1,4):
                    window[f"-{k}-"].update(value=False)

            elif values["-TRIMESTRES-"] == "Tercer Trimestre":
                meses_seleccionados = [7, 8, 9]
                for i in range(7,10):
                    window[f"-{i}-"].update(value=True)
                for j in range(10,13):
                    window[f"-{j}-"].update(value=False)
                for k in range(1,7):
                    window[f"-{k}-"].update(value=False)

            elif values["-TRIMESTRES-"] == "Cuarto Trimestre":
                meses_seleccionados = [10, 11, 12]
                for i in range(10,13):
                    window[f"-{i}-"].update(value=True)
                for j in range(1,10):
                    window[f"-{j}-"].update(value=False)

            elif values["-TRIMESTRES-"] == "Año Completo":
                meses_seleccionados = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
                for i in range(1,13):
                    window[f"-{i}-"].update(value=True)
        
        elif event == "-EXCEL-":
            ruta=sg.popup_get_folder("Seleccionar Destino")
            print(ruta)

    window.close()