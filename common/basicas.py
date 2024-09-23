#Función para convertir el numero float en un string con el formato de numeros en español
import numpy


def latinizar(numero):
    if type(numero) is numpy.float64 or type(numero) is float:
        cadena="{:,.2f}".format(numero).replace('.', '#').replace(',', '.').replace('#', ',')
        return cadena
    else:
        return ''