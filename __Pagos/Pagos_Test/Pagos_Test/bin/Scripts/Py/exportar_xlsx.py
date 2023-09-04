import shutil
from getpass import getuser
usuario = getuser()

carpeta = 'Resultados'
origen = f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/MejorCredito/Resultados'
destino = f'C:/Users/{usuario}/Desktop/{carpeta}'

try:
    shutil.copytree(origen, destino)
except:
    shutil.copytree(origen, destino, dirs_exist_ok=True)