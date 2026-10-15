import pandas as pd
import json
import re
import datetime

def normalizar_cod(c):
    if c is None or pd.isna(c):
        return None
    s = str(c).strip().split('-')[0]
    nums = re.sub(r'\D', '', s)
    return str(int(nums)).zfill(3) if nums else str(c).strip()

def normalizar_per(p):
    if not p or pd.isna(p):
        return ''
    s = str(p).strip().upper()
    return s

def extraer_periodo_base(p):
    if not p:
        return ''
    s = str(p).strip().upper()
    m = re.search(r'\b(I{1,3}|IV|V)\s*PR\b', s)
    if m:
        return m.group(0)
    m2 = re.search(r'\b(I{1,3}|IV|V)\b', s)
    return f"{m2.group(1)} PR" if m2 else s

def normalizar_servicio(s):
    if not s or pd.isna(s):
        return 'Agua y saneamiento'
    st = str(s).strip()
    if 'subterr' in st.lower():
        return 'Aguas subterráneas'
    return 'Agua y saneamiento'

def formatear_fecha(d):
    if pd.isna(d) or d is None:
        return None
    if isinstance(d, (pd.Timestamp, datetime.datetime, datetime.date)):
        return d.strftime('%Y-%m-%d')
    return str(d).strip()

def generar_estudios_tarifarios():
    excel_path = 'Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx'
    sheet_name = 'Estudios Tarifarios'
    js_output_path = 'estudios_tarifarios.js'
    
    print(f"Leyendo hoja '{sheet_name}' de '{excel_path}'...")
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=1)
    df = df.dropna(subset=['CODIGO EP'])
    
    registros = []
    for _, row in df.iterrows():
        cod_ep_raw = row['CODIGO EP']
        cod_ep_norm = normalizar_cod(cod_ep_raw)
        ep = str(row['EP']).strip() if pd.notna(row['EP']) else ''
        
        # Manejo de caracteres en columnas con tildes
        region = ''
        for col_cand in ['REGIÓN', 'REGIN', 'REGION']:
            if col_cand in df.columns and pd.notna(row[col_cand]):
                region = str(row[col_cand]).strip()
                break
                
        regimen = ''
        for col_cand in ['RÉGIMEN', 'RGIMEN', 'REGIMEN']:
            if col_cand in df.columns and pd.notna(row[col_cand]):
                regimen = str(row[col_cand]).strip()
                break
                
        tamano = ''
        for col_cand in ['TAMAÑO', 'TAMAO', 'TAMANO']:
            if col_cand in df.columns and pd.notna(row[col_cand]):
                tamano = str(row[col_cand]).strip()
                break
                
        periodo_reg = normalizar_per(row['PERIODO REGULATORIO'])
        periodo_base = extraer_periodo_base(periodo_reg)
        
        fecha_aprob = None
        for col_cand in ['FECHA DE APROBACIÓN', 'FECHA DE APROBACIN', 'FECHA DE APROBACION']:
            if col_cand in df.columns:
                fecha_aprob = formatear_fecha(row[col_cand])
                break
                
        periodo_quinquenio = str(row['PERIODO REGULATORIO_1']).strip() if pd.notna(row['PERIODO REGULATORIO_1']) else ''
        inicio_p = int(row['INICIO DE PERIODO REGULATORIO']) if pd.notna(row['INICIO DE PERIODO REGULATORIO']) else None
        fin_p = int(row['FIN DE PERIODO REGULATORIO']) if pd.notna(row['FIN DE PERIODO REGULATORIO']) else None
        
        estado = str(row['ESTADO']).strip() if pd.notna(row['ESTADO']) else ''
        tipo = str(row['TIPO']).strip() if pd.notna(row['TIPO']) else 'Definitivo'
        comentario = str(row['COMENTARIO']).strip() if pd.notna(row['COMENTARIO']) else None
        
        link = str(row['LINK']).strip() if pd.notna(row['LINK']) else ''
        if link.lower() in ('nan', 'none', '-'):
            link = ''
            
        tipo_servicio = normalizar_servicio(row['TIPO DE SERVICIO'])
        
        registros.append({
            'cod_ep': cod_ep_norm,
            'cod_ep_original': str(cod_ep_raw).strip(),
            'ep': ep,
            'region': region,
            'regimen': regimen,
            'tamano': tamano,
            'periodo': periodo_reg,
            'periodo_base': periodo_base,
            'fecha_aprobacion': fecha_aprob,
            'quinquenio': periodo_quinquenio,
            'inicio_periodo': inicio_p,
            'fin_periodo': fin_p,
            'estado': estado,
            'tipo': tipo,
            'comentario': comentario,
            'link': link,
            'tipo_servicio': tipo_servicio
        })
        
    print(f"Total registros procesados: {len(registros)}")
    con_link = sum(1 for r in registros if r['link'])
    sin_link = sum(1 for r in registros if not r['link'])
    print(f"Con link: {con_link}, Sin link: {sin_link}")
    
    with open(js_output_path, 'w', encoding='utf-8') as f:
        f.write("// Base de datos de Estudios Tarifarios de las Empresas Prestadoras (SUNASS)\n")
        f.write("// Generado automáticamente desde Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx\n")
        f.write("const db_estudios_tarifarios = ")
        json.dump(registros, f, ensure_ascii=False, indent=2)
        f.write(";\n")
        
    print(f"¡Éxito! Archivo '{js_output_path}' generado satisfactoriamente.")

if __name__ == '__main__':
    generar_estudios_tarifarios()
