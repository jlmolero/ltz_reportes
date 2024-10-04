import pandas as pd
from common.cuentas import get_movimientos

meses_seleccionados= [9]
movimientos=get_movimientos('6597',meses_seleccionados)
print(movimientos)