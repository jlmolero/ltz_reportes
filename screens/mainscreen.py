#import PySimpleGUI as sg
import PySimpleGUI as sg
from common.ejecucion import gastos_periodo




meses= [
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
    sg.Checkbox("Dic", key="-12-", font=("Helvetica", 12), enable_events=True)]
def mainscreen():



    layout = [
        [sg.Text("Prueba")],
        meses,
        [sg.Button("Consultar", key="-CONSULTAR-")]
        
    ]

    
    
    window=sg.Window("Colección de Informes", layout, resizable=True)



    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == "-1-" or event == "-2-" or event == "-3-" or event == "-4-" or event == "-5-" or event == "-6-" or event == "-7-" or event == "-8-" or event == "-9-" or event == "-10-" or event == "-11-" or event == "-12-":
            meses_seleccionados = []
            for i in range(1,12):
                if values[f"-{i}-"]:
                    meses_seleccionados.append(i)
        elif event == "-CONSULTAR-":
            pagos=gastos_periodo(meses_seleccionados)

    window.close()