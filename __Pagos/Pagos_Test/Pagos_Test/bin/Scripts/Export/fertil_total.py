import shutil
from getpass import getuser
usuario = getuser()

origen = f'M:/PAGOS/__Pagos/Pagos_Test/Pagos_Test/bin/Temp/Fertil/Resultados/Unificado Válidos y NO válidos.xlsx'
destino = f'C:/Users/{usuario}/Desktop/'

shutil.copy(origen, destino)