#Función para convertir el numero float en un string con el formato de numeros en español
import numpy


def latinizar(numero):
    if type(numero) is numpy.float64 or type(numero) is float:
        cadena="{:,.2f}".format(numero).replace('.', '#').replace(',', '.').replace('#', ',')
        return cadena
    else:
        return ''
    
def latino_a_numero(numero_latino):
    try:
        numero = float(numero_latino.replace('.', '').replace(',', '.'))
        return numero
    except:
        print('No se pudo convertir el numero')
        return ''
    
latino='1.234.567,89'
numero=latino_a_numero(latino)
print(type(numero))