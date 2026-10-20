"""Prepara una copia de los archivos reales de DARCA II sin sobrescribir datos."""
from pathlib import Path
import shutil
root=Path(__file__).resolve().parents[1]
if any((root/n).exists() for n in ('data','output')):
    raise SystemExit('Ya existe data/ u output/. Use una copia nueva del repositorio para evitar mezclar registros.')
shutil.copytree(root/'fuentes_darca',root/'data')
shutil.copytree(root/'examples',root/'output')
print('Archivos de DARCA II preparados en data/ y output/.')
