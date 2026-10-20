"""Genera fuentes ficticias desde cero. Nunca sobrescribe fuentes existentes."""
from pathlib import Path
from datetime import date
import json
from openpyxl import Workbook
from PIL import Image, ImageDraw


def crear(root):
    mapping = json.loads((root/'config/mapeo_wbs.json').read_text(encoding='utf-8'))
    presupuesto = root/'data/presupuesto/presupuesto_demo.xlsx'
    reportes = [root/f'data/reportes/avance_demo_{dia}.xlsx' for dia in (10,17,24)]
    if presupuesto.exists() or any(p.exists() for p in reportes):
        raise SystemExit('Ya existen archivos demo. Use una copia nueva del repositorio para repetir la preparación.')
    for d in ('data/presupuesto','data/reportes','data/historico','assets','output'):
        (root/d).mkdir(parents=True,exist_ok=True)
    wb=Workbook();ws=wb.active;ws.title='Ppto Ajustado'
    ws.append(['DEMOSTRACIÓN: TODOS LOS DATOS SON FICTICIOS'])
    ws.append([]);ws.append([])
    ws.append(['Ítem','Concepto','Descripción','Unidad','Cantidad','Precio unitario','Valor total'])
    directo=0
    for i, codigos in sorted(mapping.items(),key=lambda x:int(x[0])):
        item=int(i);cantidad=192 if item==11 else 100*len(codigos);precio=10000+item*1000
        valor=cantidad*precio;directo+=valor
        ws.append([item,f'Actividad demostrativa {item:02}',f'Trabajo ficticio para demostrar el cálculo del ítem {item}.','m²',cantidad,precio,valor])
    for label,value in [('Administración',.10),('Imprevistos',.04),('Utilidad',.05),('IVA sobre utilidad',.19),('Neto obra',directo),('Valor total contractual',directo*1.1995)]:
        ws.append([label,value])
    wb.save(presupuesto);wb.close()
    for dia,porcentaje,path in zip((10,17,24),(.10,.20,.15),reportes):
        wb=Workbook();ws=wb.active;ws.title='DATOS'
        ws.append(['DATOS FICTICIOS — NO UTILIZAR PARA COBRO']);ws.append([]);ws.append([])
        ws.append(['Fecha Corte','ID','WBS','Actividad','Tramo','Unidad','Cantidad Total','Cantidad Ejecutada Semana','Observación'])
        ident=0
        for item,codigos in sorted(mapping.items(),key=lambda x:int(x[0])):
            for wbs in codigos:
                ident+=1;total=192 if int(item)==11 else 100
                ws.append([date(2026,1,dia),ident,wbs,f'Actividad demostrativa {int(item):02}',f'TRAMO {wbs.split(".")[1]}','m²',total,total*porcentaje,'Ejemplo sintético'])
        wb.save(path);wb.close()
    logo=root/'assets/logo_empresa.png'
    if not logo.exists():
        im=Image.new('RGB',(640,110),'#DCAD37');draw=ImageDraw.Draw(im)
        draw.text((35,28),'PROYECTO DEMOSTRATIVO',fill='#173047',font_size=30)
        im.save(logo)
    return presupuesto


if __name__=='__main__':
    print(crear(Path(__file__).resolve().parents[1]))
