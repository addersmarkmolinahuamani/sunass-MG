import openpyxl
import json
from datetime import datetime, date

print("Cargando Excel para extraer Estudios Tarifarios (ET) y RCD...")
wb = openpyxl.load_workbook('Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx', data_only=True)

def clean_val(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).replace('\xa0', ' ').strip()
    if s in ('', '-', 'None', 'nan'):
        return None
    return v if isinstance(v, (int, float)) else s

# 1. ESTUDIOS TARIFARIOS
sheet_et = 'ET' if 'ET' in wb.sheetnames else 'Estudios Tarifarios'
ws_et = wb[sheet_et]
data_et = []
for r in range(3, ws_et.max_row + 1):
    cod = ws_et.cell(r, 2).value
    ep = ws_et.cell(r, 3).value
    if cod is None and ep is None:
        continue
    cod_str = str(cod).strip()
    if cod_str.isdigit():
        cod_str = cod_str.zfill(3)
    elif cod_str == '008-052':
        cod_str = '008'
        
    cell_link = ws_et.cell(r, 15)
    link_val = clean_val(cell_link.value)
    if cell_link.hyperlink and cell_link.hyperlink.target:
        link_val = cell_link.hyperlink.target

    data_et.append({
        "cod_ep": cod_str,
        "ep": clean_val(ws_et.cell(r, 3).value),
        "region": clean_val(ws_et.cell(r, 4).value),
        "regimen": clean_val(ws_et.cell(r, 5).value),
        "tamano": clean_val(ws_et.cell(r, 6).value),
        "periodo": clean_val(ws_et.cell(r, 7).value),
        "fecha_aprobacion": clean_val(ws_et.cell(r, 8).value),
        "quinquenio": clean_val(ws_et.cell(r, 9).value),
        "inicio_periodo": clean_val(ws_et.cell(r, 10).value),
        "fin_periodo": clean_val(ws_et.cell(r, 11).value),
        "estado": clean_val(ws_et.cell(r, 12).value),
        "tipo": clean_val(ws_et.cell(r, 13).value),
        "comentario": clean_val(ws_et.cell(r, 14).value),
        "link": link_val,
        "tipo_servicio": clean_val(ws_et.cell(r, 16).value)
    })

with open('estudios_tarifarios.js', 'w', encoding='utf-8') as f:
    f.write('// Base de datos de Estudios Tarifarios de las Empresas Prestadoras (SUNASS)\n')
    f.write('// Generado automáticamente desde Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx\n')
    f.write('const db_estudios_tarifarios = ')
    json.dump(data_et, f, ensure_ascii=False, indent=2)
    f.write(';\n')

print(f"Extraídos {len(data_et)} Estudios Tarifarios -> estudios_tarifarios.js")

# 2. RESOLUCIONES DE CONSEJO DIRECTIVO (RCD)
ws_rcd = wb['RCD']
data_rcd = []
for r in range(3, ws_rcd.max_row + 1):
    cod = ws_rcd.cell(r, 2).value
    ep = ws_rcd.cell(r, 3).value
    if cod is None and ep is None:
        continue
    cod_str = str(cod).strip()
    if cod_str.isdigit():
        cod_str = cod_str.zfill(3)
    elif cod_str == '008-052':
        cod_str = '008'

    cell_link = ws_rcd.cell(r, 18)
    link_val = clean_val(cell_link.value)
    if cell_link.hyperlink and cell_link.hyperlink.target:
        link_val = cell_link.hyperlink.target

    data_rcd.append({
        "cod_ep": cod_str,
        "ep": clean_val(ws_rcd.cell(r, 3).value),
        "region": clean_val(ws_rcd.cell(r, 4).value),
        "regimen": clean_val(ws_rcd.cell(r, 5).value),
        "tamano": clean_val(ws_rcd.cell(r, 6).value),
        "periodo": clean_val(ws_rcd.cell(r, 7).value),
        "rcd_num": clean_val(ws_rcd.cell(r, 8).value),
        "fecha_publicacion": clean_val(ws_rcd.cell(r, 9).value),
        "quinquenio": clean_val(ws_rcd.cell(r, 10).value),
        "inicio_periodo": clean_val(ws_rcd.cell(r, 11).value),
        "fin_periodo": clean_val(ws_rcd.cell(r, 12).value),
        "estado": clean_val(ws_rcd.cell(r, 13).value),
        "tipo": clean_val(ws_rcd.cell(r, 14).value),
        "rebalanceo": clean_val(ws_rcd.cell(r, 15).value),
        "titulo": clean_val(ws_rcd.cell(r, 16).value),
        "comentario": clean_val(ws_rcd.cell(r, 17).value),
        "link": link_val,
        "tipo_servicio": clean_val(ws_rcd.cell(r, 19).value)
    })

with open('resoluciones_cd.js', 'w', encoding='utf-8') as f:
    f.write('// Base de datos de Resoluciones de Consejo Directivo (RCD) de las Empresas Prestadoras (SUNASS)\n')
    f.write('// Generado automáticamente desde Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx\n')
    f.write('const db_resoluciones_cd = ')
    json.dump(data_rcd, f, ensure_ascii=False, indent=2)
    f.write(';\n')

def generar_estudios_tarifarios():
    pass

if __name__ == '__main__':
    generar_estudios_tarifarios()

