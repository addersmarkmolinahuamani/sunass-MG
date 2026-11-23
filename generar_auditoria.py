import json
import re
import os
from datetime import datetime
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_ORIGEN = "Compilatorio ICI_ICG_EPS TOTALES_ACT EN PROCESO.xlsx"
EXCEL_REPORTE = "Auditoria_Calidad_Base_Datos_SUNASS.xlsx"
JS_SALIDA = "auditoria_datos.js"

def load_js_var(path, var_name=None):
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        txt = f.read()
    if var_name:
        m = re.search(rf'(?:const|var|let)\s+{var_name}\s*=\s*', txt)
    else:
        m = re.search(r'(?:const|var|let)\s+\w+\s*=\s*', txt)
    if not m:
        return []
    decoder = json.JSONDecoder()
    obj, _ = decoder.raw_decode(txt, m.end())
    return obj

def norm_cod(c):
    if c is None:
        return ''
    nums = re.sub(r'\D', '', str(c))
    return nums.zfill(3) if nums else str(c).strip()

def is_empty(v):
    if v is None:
        return True
    s = str(v).strip()
    return s in ('', '-', '–', '—', 'None', 'nan', 'NaN', 'null')

def es_doc_real(v):
    """Devuelve True si el texto representa un número de informe/resolución real (no nota ni placeholder)."""
    if is_empty(v):
        return False
    s = str(v).strip()
    u = s.upper()
    if any(k in u for k in [
        'FECHA TENTATIVA', 'PENDIENTE', 'SIN EVALUAC', 'NO EVALUAD',
        'SUSPENDID', 'EN PROCESO', 'EN CURSO', 'NO APLICA', 'N/A'
    ]):
        return False
    if u in ('X', 'HUARAL') or not re.search(r'\d', s):
        return False
    return True

def es_placeholder_sospechoso(v):
    """Detecta textos de relleno o errores de digitación en columnas de informes (ej. 'X', 'huaral')."""
    if is_empty(v):
        return False
    s = str(v).strip()
    u = s.upper()
    if any(k in u for k in [
        'FECHA TENTATIVA', 'PENDIENTE', 'SIN EVALUAC', 'NO EVALUAD',
        'SUSPENDID', 'EN PROCESO', 'EN CURSO', 'NO APLICA', 'N/A',
        'ARCHIV', 'IMPROCED', 'INFUNDAD'
    ]):
        return False
    if not re.search(r'\d', s):
        return True
    return False

def ejecutar_auditoria():
    print("1. Cargando bases de datos del aplicativo...")
    metas = load_js_var('datos_sunass.js', 'baseDatosNacional') or []
    informes = load_js_var('datos_sunass.js', 'matrizInformesNacional') or []
    et_list = load_js_var('estudios_tarifarios.js') or []
    rcd_list = load_js_var('resoluciones_cd.js') or []

    eps_names = {}
    bloques_etapas = defaultdict(set)
    por_anio = defaultdict(lambda: {
        'metas_por_etapa': defaultdict(list),
        'icg_por_etapa': {},
        'info': None
    })

    for r in metas:
        cod = norm_cod(r.get('Cod_EP'))
        nom = str(r.get('Nombre_EPS') or '').replace('\xa0', ' ').strip()
        nom = re.sub(r'\s+', ' ', nom)
        if cod and nom:
            eps_names[cod] = nom
        per = str(r.get('Periodo_Regulatorio') or '').strip()
        serv = str(r.get('Tipo_Servicio') or 'Agua y Saneamiento').strip()
        anio = str(r.get('Año_Regulatorio') or '').strip()
        etapa = str(r.get('Etapa') or '').strip()
        bloques_etapas[(cod, per, serv)].add(etapa)
        if anio in ('Año Base', 'Acumulado', ''):
            continue
        k = (cod, per, serv, anio)
        mp = str(r.get('Meta_Principal') or '')
        if 'Índice de Cumplimiento Global' in mp or mp.strip().upper() == 'ICG':
            if not is_empty(r.get('ICI')):
                por_anio[k]['icg_por_etapa'][etapa] = r.get('ICI')
        else:
            por_anio[k]['metas_por_etapa'][etapa].append(r)

    for inf in informes:
        cod = norm_cod(inf.get('Cod_EP'))
        nom = str(inf.get('Nombre_EPS') or '').replace('\xa0', ' ').strip()
        nom = re.sub(r'\s+', ' ', nom)
        if cod and nom and cod not in eps_names:
            eps_names[cod] = nom
        per = str(inf.get('Periodo_Regulatorio') or '').strip()
        serv = str(inf.get('Tipo_Servicio') or 'Agua y Saneamiento').strip()
        anio = str(inf.get('Año_Regulatorio') or '').strip()
        if anio in ('Año Base', 'Acumulado', ''):
            continue
        k = (cod, per, serv, anio)
        por_anio[k]['info'] = inf

    inf_cols = [
        ('Inf_Fisc_Inicial', 'FISCALIZACIÓN', 'Inf. Fisc. Inicial'),
        ('Inf_Fisc_Final', 'FISCALIZACIÓN', 'Inf. Fisc. Final'),
        ('Inf_Fiscalizacion', 'FISCALIZACIÓN', 'Inf. Fiscalización'),
        ('Inf_Fisc_Complementario', 'FISCALIZACIÓN', 'Inf. Fisc. Complementario'),
        ('Inf_Fisc_MedidasCorrectivas', 'MEDIDAS', 'Inf. Medidas Correctivas'),
        ('Inf_PAS_Instruccion', 'INSTRUCCIÓN', 'Inf. PAS Instrucción'),
        ('Inf_PAS_Complementario', 'INSTRUCCIÓN', 'Inf. PAS Complementario'),
        ('Inf_PAS_Decision', 'DECISIÓN', 'Res. PAS Decisión'),
        ('Recurso_Reconsideracion', 'RECONSIDERACIÓN', 'Recurso Reconsideración'),
        ('Recurso_Apelacion', 'APELACIÓN', 'Recurso Apelación'),
    ]

    hallazgos = []
    links_faltantes = []

    def add_h(severidad, categoria, cod, per, serv, anio, etapa, detalle, accion, tipo_visual='general'):
        hallazgos.append({
            'id': '',
            'severidad': severidad,
            'categoria': categoria,
            'cod_eps': cod,
            'nombre_eps': eps_names.get(cod, cod),
            'periodo': per,
            'servicio': serv,
            'anio': anio,
            'etapa': etapa,
            'detalle': detalle,
            'accion': accion,
            'tipo_visual': tipo_visual  # 'icg', 'etapa', 'fecha', 'excel', 'ici', 'informe'
        })

    print("2. Analizando consistencia por EPS, Periodo, Año y Etapa...")
    for k in sorted(por_anio.keys()):
        cod, per, serv, anio = k
        data = por_anio[k]
        inf = data['info'] or {}
        estado = str(inf.get('Estado_Evaluacion') or '').strip()
        rango = str(inf.get('Rango_Anio') or '').strip()
        icg_map = data['icg_por_etapa']
        metas_map = data['metas_por_etapa']
        etapas_bloque = bloques_etapas.get((cod, per, serv), set())
        es_multi_etapa = len(etapas_bloque - {'RESULTADO_ACTUAL', 'FISCALIZACIÓN'}) > 0

        infs_eval = []
        infs_por_etapa = defaultdict(list)

        for col, et_col, label in inf_cols:
            val = inf.get(col)
            if is_empty(val):
                continue
            val_s = str(val).strip().replace('\n', ' ')
            if es_placeholder_sospechoso(val):
                add_h(
                    'CRÍTICO',
                    'Texto inválido / error de digitación en Informes',
                    cod, per, serv, anio, et_col,
                    f"En '{label}' se registró el texto inválido '{val_s}' en lugar de un número de informe o resolución.",
                    f"Corregir o limpiar la celda de '{label}' en la hoja INFORMES TOTALES.",
                    'informe'
                )
                continue

            if es_doc_real(val):
                if et_col != 'MEDIDAS':
                    infs_eval.append(f"{label}: {val_s}")
                    infs_por_etapa[et_col].append(f"{label}: {val_s}")
                lnk = inf.get(col + '_Link')
                if is_empty(lnk):
                    links_faltantes.append({
                        'fuente': 'INFORMES TOTALES',
                        'cod_eps': cod,
                        'nombre_eps': eps_names.get(cod, cod),
                        'periodo': per,
                        'servicio': serv,
                        'anio': anio,
                        'tipo_doc': label,
                        'documento': val_s
                    })

        # A) Resultado Actual desactualizado respecto a última etapa con ICG
        orden_et = ['APELACIÓN', 'RECONSIDERACIÓN', 'DECISIÓN', 'INSTRUCCIÓN', 'FISCALIZACIÓN']
        ultima_et = next((e for e in orden_et if e in icg_map), None)
        if ultima_et:
            if 'RESULTADO_ACTUAL' not in icg_map:
                add_h(
                    'CRÍTICO',
                    'Resultado Actual sin ICG (existe en etapa)',
                    cod, per, serv, anio, 'RESULTADO_ACTUAL',
                    f"Falta el ICG en Resultado Actual, pero la etapa {ultima_et} sí registra ICG = {icg_map[ultima_et]}%.",
                    f"Copiar o vincular el ICG de {ultima_et} ({icg_map[ultima_et]}%) en el cuadro de Resultado Actual.",
                    'icg'
                )
            else:
                try:
                    v_act = float(icg_map['RESULTADO_ACTUAL'])
                    v_ult = float(icg_map[ultima_et])
                    if abs(v_act - v_ult) > 0.05:
                        add_h(
                            'CRÍTICO',
                            'ICG de Resultado Actual desactualizado',
                            cod, per, serv, anio, 'RESULTADO_ACTUAL',
                            f"Resultado Actual muestra ICG = {v_act}%, pero la última etapa registrada ({ultima_et}) tiene ICG = {v_ult}% (diferencia: {round(v_ult - v_act, 2):+.2f} pp).",
                            f"Actualizar el cuadro de Resultado Actual con los resultados vigentes de {ultima_et} (ICG = {v_ult}%).",
                            'icg'
                        )
                except (ValueError, TypeError):
                    pass

        # B) Informes reales cargados pero SIN ICG en Resultado Actual
        if len(infs_eval) > 0 and 'RESULTADO_ACTUAL' not in icg_map:
            icis_act = [r for r in metas_map.get('RESULTADO_ACTUAL', []) if not is_empty(r.get('ICI'))]
            if estado.upper() == 'SUSPENDIDO':
                tiene_pas_o_final = any(e in infs_por_etapa for e in ('INSTRUCCIÓN', 'DECISIÓN', 'RECONSIDERACIÓN', 'APELACIÓN')) or es_doc_real(inf.get('Inf_Fisc_Final'))
                if tiene_pas_o_final:
                    add_h(
                        'MEDIO',
                        'Año Suspendido con informes sin ICG',
                        cod, per, serv, anio, 'RESULTADO_ACTUAL',
                        f"Estado 'Suspendido' con informes registrados ({'; '.join(infs_eval[:2])}), pero sin ICG.",
                        "Confirmar si corresponde mantener sin ICG por suspensión o si falta registrarlo.",
                        'icg'
                    )
            else:
                det_extra = f" Además ya tiene {len(icis_act)} ICI de metas cargados." if icis_act else ""
                add_h(
                    'CRÍTICO',
                    'Informes cargados sin ICG en Resultado Actual',
                    cod, per, serv, anio, 'RESULTADO_ACTUAL',
                    f"Estado '{estado or 'Sin estado'}': cuenta con informes ({'; '.join(infs_eval[:3])}) pero NO tiene ICG en Resultado Actual.{det_extra}",
                    "Cargar/calcular el ICG y los ICI faltantes en la hoja de la EPS.",
                    'icg'
                )

        # C) Tiene ICI en Resultado Actual pero falta ICG y no tiene informes
        icis_act_cnt = sum(1 for r in metas_map.get('RESULTADO_ACTUAL', []) if not is_empty(r.get('ICI')))
        if icis_act_cnt > 0 and 'RESULTADO_ACTUAL' not in icg_map and len(infs_eval) == 0:
            add_h(
                'CRÍTICO',
                'ICI cargados sin ICG ni informes',
                cod, per, serv, anio, 'RESULTADO_ACTUAL',
                f"Tiene {icis_act_cnt} metas con ICI en Resultado Actual, pero falta el ICG y no tiene informes en 'INFORMES TOTALES'.",
                "Ingresar el ICG en la hoja de la EPS y registrar el informe de respaldo en 'INFORMES TOTALES'.",
                'icg'
            )

        # D) Tiene ICG en Resultado Actual pero SIN informes en INFORMES TOTALES
        if 'RESULTADO_ACTUAL' in icg_map and len(infs_eval) == 0:
            add_h(
                'ALTO',
                'ICG registrado sin informes de respaldo',
                cod, per, serv, anio, 'RESULTADO_ACTUAL',
                f"Registra ICG = {icg_map['RESULTADO_ACTUAL']}% en Resultado Actual, pero no tiene ningún informe cargado en 'INFORMES TOTALES'.",
                "Cargar el número de informe/resolución correspondiente en la hoja 'INFORMES TOTALES'.",
                'icg'
            )

        # E) En bloques multi-etapa: informe de etapa sin datos en esa etapa
        if es_multi_etapa:
            for et_check in ['FISCALIZACIÓN', 'INSTRUCCIÓN', 'DECISIÓN', 'RECONSIDERACIÓN', 'APELACIÓN']:
                if et_check in infs_por_etapa:
                    has_icg_et = et_check in icg_map
                    has_ici_et = any(not is_empty(r.get('ICI')) for r in metas_map.get(et_check, []))
                    if not has_icg_et and not has_ici_et:
                        sev = 'ALTO' if et_check in ('FISCALIZACIÓN', 'INSTRUCCIÓN', 'DECISIÓN') else 'MEDIO'
                        add_h(
                            sev,
                            f'Informe de {et_check.title()} sin datos en esa etapa',
                            cod, per, serv, anio, et_check,
                            f"Existe documento en INFORMES TOTALES ({'; '.join(infs_por_etapa[et_check])}), pero el cuadro de {et_check} está vacío (sin ICG ni ICI).",
                            f"Verificar el documento y llenar los ICI/ICG de la etapa {et_check} en la hoja de la EPS (si corresponde).",
                            'etapa'
                        )
                    elif not has_icg_et and has_ici_et:
                        add_h(
                            'ALTO',
                            f'Etapa {et_check.title()} con ICI pero sin ICG',
                            cod, per, serv, anio, et_check,
                            f"La etapa {et_check} tiene ICI individuales registrados pero le falta el cálculo/valor del ICG.",
                            f"Completar la fila del ICG en el cuadro de {et_check} de la hoja de la EPS.",
                            'icg'
                        )

        # F) Valores de ICG fuera de rango o no numéricos
        icg_no_num_visto = set()
        for et, val in icg_map.items():
            try:
                vf = float(val)
                if vf < 0 or vf > 100.05:
                    add_h(
                        'CRÍTICO',
                        'Valor de ICG fuera de rango (0-100%)',
                        cod, per, serv, anio, et,
                        f"El ICG en {et} es {vf}%, fuera del rango válido [0% - 100%].",
                        "Revisar la fórmula o el número digitado en la celda del ICG.",
                        'icg'
                    )
            except (ValueError, TypeError):
                if (cod, per, anio, str(val)) not in icg_no_num_visto:
                    icg_no_num_visto.add((cod, per, anio, str(val)))
                    add_h(
                        'CRÍTICO',
                        'Valor de ICG no numérico (formato texto)',
                        cod, per, serv, anio, et,
                        f"El ICG en {et} está escrito como texto '{val}' (con asterisco o coma), impidiendo cálculos y gráficos.",
                        "Reemplazar por un valor numérico puro (ej. 99.2%) y mover cualquier nota a un comentario.",
                        'icg'
                    )

        # G) Valores de ICI anómalos o fuera de rango en Resultado Actual
        ici_anomalos = []
        for r in metas_map.get('RESULTADO_ACTUAL', []):
            v = r.get('ICI')
            if is_empty(v):
                continue
            try:
                vf = float(v)
                if vf < 0 or vf > 100.05:
                    ici_anomalos.append(f"{r.get('Meta_Principal')}: {vf}%")
            except (ValueError, TypeError):
                sv = str(v).strip()
                if sv.upper() not in ('NA', 'N/A', 'ND', 'N/D', 'NO APLICA', 'P', 'C', 'SC', 'SE', 'SD'):
                    ici_anomalos.append(f"{r.get('Meta_Principal')}: '{sv}'")
        if ici_anomalos:
            add_h(
                'ALTO',
                'Valores de ICI anómalos o fuera de rango',
                cod, per, serv, anio, 'RESULTADO_ACTUAL',
                f"Se hallaron {len(ici_anomalos)} ICI fuera de [0%-100%] o con texto inválido: {'; '.join(ici_anomalos[:3])}",
                "Corregir los valores o fórmulas de ICI en la hoja de la EPS.",
                'ici'
            )

        # H) Rango de fechas faltante o incoherente
        if is_empty(rango):
            if not anio.startswith('PT'):
                add_h(
                    'MEDIO',
                    'Año regulatorio sin rango de fechas',
                    cod, per, serv, anio, 'GENERAL',
                    f"El {anio} del {per} no tiene definido el rango de fechas en 'INFORMES TOTALES'.",
                    "Ingresar el mes/año de inicio y fin del año regulatorio en 'INFORMES TOTALES'.",
                    'fecha'
                )
        else:
            m_r = re.match(r'^([A-Za-záéíóú]{3})\s+(\d{2,4})\s*-\s*([A-Za-záéíóú]{3})\s+(\d{2,4})$', rango)
            if m_r:
                y1 = int(m_r.group(2))
                y2 = int(m_r.group(4))
                if y1 < 100: y1 += 2000
                if y2 < 100: y2 += 2000
                if y2 < y1 or (y2 - y1) > 2:
                    add_h(
                        'ALTO',
                        'Rango de fechas incoherente en año regulatorio',
                        cod, per, serv, anio, 'GENERAL',
                        f"El {anio} registra el rango '{rango}' ({y1} a {y2}), que abarca más de 1 año regulatorio (posible copia del quinquenio completo).",
                        "Corregir la fecha de inicio y fin exacta de ese año regulatorio en 'INFORMES TOTALES'.",
                        'fecha'
                    )

    # 3. Enlaces faltantes en Estudios Tarifarios y RCD
    print("3. Verificando Estudios Tarifarios y Resoluciones de Consejo Directivo...")
    for e in et_list:
        if is_empty(e.get('link')):
            cod = norm_cod(e.get('cod_ep'))
            links_faltantes.append({
                'fuente': 'ESTUDIOS TARIFARIOS (ET)',
                'cod_eps': cod,
                'nombre_eps': e.get('ep') or eps_names.get(cod, cod),
                'periodo': str(e.get('periodo') or '-'),
                'servicio': str(e.get('tipo_servicio') or 'Agua y Saneamiento'),
                'anio': str(e.get('quinquenio') or '-'),
                'tipo_doc': str(e.get('tipo') or 'Estudio Tarifario'),
                'documento': f"ET {e.get('periodo', '')} ({e.get('quinquenio', '')})".strip()
            })

    for r in rcd_list:
        if is_empty(r.get('link')):
            cod = norm_cod(r.get('cod_ep'))
            doc_num = r.get('rcd') or r.get('resolucion') or r.get('numero') or 'RCD'
            links_faltantes.append({
                'fuente': 'RESOLUCIONES CD (RCD)',
                'cod_eps': cod,
                'nombre_eps': r.get('ep') or eps_names.get(cod, cod),
                'periodo': str(r.get('periodo') or '-'),
                'servicio': str(r.get('tipo_servicio') or 'Agua y Saneamiento'),
                'anio': str(r.get('fecha_publicacion') or '-'),
                'tipo_doc': str(r.get('tipo') or 'RCD'),
                'documento': str(doc_num)
            })

    # 4. Escanear errores de fórmula en el Excel Compilatorio
    print("4. Escaneando errores de fórmula Excel (#REF!, #DIV/0!, etc.)...")
    wb_src = openpyxl.load_workbook(EXCEL_ORIGEN, read_only=True, data_only=True)
    err_tokens = {'#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!'}
    for sh in wb_src.sheetnames:
        if not re.match(r'^\d{3}$', sh):
            continue
        ws = wb_src[sh]
        sh_errs = []
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                if isinstance(val, str) and val.strip().upper() in err_tokens:
                    sh_errs.append(f"{get_column_letter(c_idx)}{r_idx}={val.strip()}")
        if sh_errs:
            add_h(
                'CRÍTICO',
                'Error de fórmula en Excel (#REF!, #DIV/0!, etc.)',
                sh, 'General', 'Agua y Saneamiento', '-', 'EXCEL',
                f"La hoja '{sh}' contiene {len(sh_errs)} celdas con errores de fórmula Excel: {', '.join(sh_errs[:6])}" + ("..." if len(sh_errs) > 6 else ""),
                f"Abrir la hoja '{sh}' en el archivo Excel compilatorio y corregir las fórmulas rotas.",
                'excel'
            )
    wb_src.close()

    # Ordenar hallazgos por severidad y código EPS
    sev_order = {'CRÍTICO': 0, 'ALTO': 1, 'MEDIO': 2, 'BAJO': 3}
    hallazgos.sort(key=lambda x: (sev_order.get(x['severidad'], 9), x['cod_eps'], x['periodo'], x['anio']))
    for idx, h in enumerate(hallazgos, 1):
        h['id'] = f"AUD-{idx:03d}"

    # Estadísticas
    cnt_sev = Counter(h['severidad'] for h in hallazgos)
    cnt_cat = Counter((h['severidad'], h['categoria']) for h in hallazgos)
    links_por_eps = Counter(l['cod_eps'] for l in links_faltantes)
    eps_stats = {}
    all_cods = sorted(set(eps_names.keys()) | set(h['cod_eps'] for h in hallazgos) | set(links_por_eps.keys()))
    for cod in all_cods:
        if not re.match(r'^\d{3}$', cod):
            continue
        hs = [h for h in hallazgos if h['cod_eps'] == cod]
        eps_stats[cod] = {
            'cod_eps': cod,
            'nombre_eps': eps_names.get(cod, cod),
            'criticos': sum(1 for h in hs if h['severidad'] == 'CRÍTICO'),
            'altos': sum(1 for h in hs if h['severidad'] == 'ALTO'),
            'medios': sum(1 for h in hs if h['severidad'] == 'MEDIO'),
            'links_faltantes': links_por_eps.get(cod, 0)
        }

    fecha_gen = datetime.now().strftime("%d/%m/%Y %H:%M")

    # =========================================================
    # GENERAR REPORTE EXCEL PROFESIONAL
    # =========================================================
    print(f"5. Generando reporte Excel profesional: {EXCEL_REPORTE}...")
    wb = openpyxl.Workbook()

    # Estilos
    f_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    f_sub = Font(name="Calibri", size=10, italic=True, color="DCE6F1")
    f_sec = Font(name="Calibri", size=11, bold=True, color="003366")
    f_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_norm = Font(name="Calibri", size=10, color="1E293B")
    f_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")

    fill_navy = PatternFill("solid", fgColor="003366")
    fill_sub = PatternFill("solid", fgColor="1E4E79")
    fill_sec = PatternFill("solid", fgColor="E8F0FE")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")

    sev_fills = {
        'CRÍTICO': PatternFill("solid", fgColor="FEE2E2"),
        'ALTO': PatternFill("solid", fgColor="FFEDD5"),
        'MEDIO': PatternFill("solid", fgColor="FEF9C3"),
        'BAJO': PatternFill("solid", fgColor="E0F2FE"),
    }
    sev_fonts = {
        'CRÍTICO': Font(name="Calibri", size=10, bold=True, color="B91C1C"),
        'ALTO': Font(name="Calibri", size=10, bold=True, color="C2410C"),
        'MEDIO': Font(name="Calibri", size=10, bold=True, color="A16207"),
        'BAJO': Font(name="Calibri", size=10, bold=True, color="0369A1"),
    }

    thin = Side(border_style="thin", color="CBD5E1")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- HOJA 1: RESUMEN EJECUTIVO ---
    ws1 = wb.active
    ws1.title = "Resumen Ejecutivo"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:F1")
    ws1["A1"] = "AUDITORÍA DE CALIDAD DE LA BASE DE DATOS — METAS DE GESTIÓN (SUNASS)"
    ws1["A1"].font = f_title
    ws1["A1"].fill = fill_navy
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 30

    ws1.merge_cells("A2:F2")
    ws1["A2"] = f"Fecha de auditoría: {fecha_gen}  |  Archivo auditado: {EXCEL_ORIGEN}"
    ws1["A2"].font = f_sub
    ws1["A2"].fill = fill_sub
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 20

    # KPIs de Severidad
    ws1["A4"] = "1. RESUMEN GENERAL DE HALLAZGOS POR NIVEL DE PRIORIDAD"
    ws1["A4"].font = f_sec
    headers_kpi = ["Nivel de Prioridad", "Cantidad", "Impacto en el Aplicativo / Credibilidad", "Acción Recomendada"]
    for c_i, h_txt in enumerate(headers_kpi, 1):
        cell = ws1.cell(5, c_i, h_txt)
        cell.font = f_hdr
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center" if c_i == 2 else "left", vertical="center")
        cell.border = border_all
    ws1.row_dimensions[5].height = 22

    resumen_niveles = [
        ("CRÍTICO", cnt_sev['CRÍTICO'], "Datos contradictorios o faltantes que afectan directamente los indicadores ICG, gráficos o fórmulas.", "Corregir de forma inmediata (ver pestaña 'Críticos y Altos')."),
        ("ALTO", cnt_sev['ALTO'], "Inconsistencias entre informes y etapas, fechas mal digitadas o valores de ICI anómalos.", "Revisar y regularizar en el corto plazo."),
        ("MEDIO", cnt_sev['MEDIO'], "Información complementaria incompleta (años sin rango de fechas, etapas de recurso sin cuadro).", "Validar si corresponde completar o dejar en blanco."),
        ("BAJO (Enlaces)", len(links_faltantes), "Documentos registrados (Informes, ET o RCD) sin hipervínculo activo hacia SharePoint/OneDrive.", "Adjuntar los enlaces faltantes (ver pestaña 'Enlaces Faltantes')."),
    ]
    for idx_r, (niv, cant, imp, acc) in enumerate(resumen_niveles, 6):
        ws1.row_dimensions[idx_r].height = 24
        c1 = ws1.cell(idx_r, 1, niv)
        c2 = ws1.cell(idx_r, 2, cant)
        c3 = ws1.cell(idx_r, 3, imp)
        c4 = ws1.cell(idx_r, 4, acc)
        sev_k = niv.split()[0]
        c1.fill = sev_fills.get(sev_k, fill_zebra)
        c1.font = sev_fonts.get(sev_k, f_bold)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.font = f_bold
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.font = f_norm
        c3.alignment = Alignment(vertical="center", wrap_text=True)
        c4.font = f_norm
        c4.alignment = Alignment(vertical="center", wrap_text=True)
        for cx in (c1, c2, c3, c4):
            cx.border = border_all

    # Tabla por Categoría
    row_cat = 12
    ws1.cell(row_cat, 1, "2. DESGLOSE DE HALLAZGOS POR TIPO DE ERROR").font = f_sec
    for c_i, h_txt in enumerate(["Severidad", "Categoría de Error / Inconsistencia", "Cantidad", "Porcentaje"], 1):
        cell = ws1.cell(row_cat + 1, c_i, h_txt)
        cell.font = f_hdr
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center" if c_i in (1, 3, 4) else "left", vertical="center")
        cell.border = border_all
    ws1.row_dimensions[row_cat + 1].height = 22

    total_h = len(hallazgos) or 1
    r_cur = row_cat + 2
    for (sev, cat), cnt in sorted(cnt_cat.items(), key=lambda x: (sev_order.get(x[0][0], 9), -x[1])):
        ws1.row_dimensions[r_cur].height = 20
        c1 = ws1.cell(r_cur, 1, sev)
        c2 = ws1.cell(r_cur, 2, cat)
        c3 = ws1.cell(r_cur, 3, cnt)
        c4 = ws1.cell(r_cur, 4, f"{cnt / total_h * 100:.1f}%")
        c1.fill = sev_fills.get(sev, fill_zebra)
        c1.font = sev_fonts.get(sev, f_bold)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.font = f_norm
        c3.font = f_bold
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.font = f_norm
        c4.alignment = Alignment(horizontal="center", vertical="center")
        for cx in (c1, c2, c3, c4):
            cx.border = border_all
        r_cur += 1

    ws1.column_dimensions["A"].width = 20
    ws1.column_dimensions["B"].width = 48
    ws1.column_dimensions["C"].width = 48
    ws1.column_dimensions["D"].width = 46

    def llenar_tabla_hallazgos(ws, lista, titulo_hoja):
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells("A1:L1")
        ws["A1"] = titulo_hoja
        ws["A1"].font = f_title
        ws["A1"].fill = fill_navy
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        cols = [
            ("ID", 10), ("Severidad", 12), ("Cód. EPS", 10), ("Empresa Prestadora", 28),
            ("Periodo", 10), ("Tipo de Servicio", 18), ("Año", 10), ("Etapa", 18),
            ("Categoría de Error", 36), ("Detalle del Problema Detectado", 56),
            ("Acción Correctiva Sugerida", 44), ("Estado Corrección", 18)
        ]
        ws.row_dimensions[3].height = 24
        for c_i, (h_txt, w) in enumerate(cols, 1):
            cell = ws.cell(3, c_i, h_txt)
            cell.font = f_hdr
            cell.fill = fill_navy
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_all
            ws.column_dimensions[get_column_letter(c_i)].width = w

        for idx, h in enumerate(lista, 4):
            ws.row_dimensions[idx].height = 32
            vals = [
                h['id'], h['severidad'], h['cod_eps'], h['nombre_eps'],
                h['periodo'], h['servicio'], h['anio'], h['etapa'],
                h['categoria'], h['detalle'], h['accion'], "Pendiente"
            ]
            for c_i, v in enumerate(vals, 1):
                cell = ws.cell(idx, c_i, v)
                cell.border = border_all
                if c_i == 2:
                    cell.fill = sev_fills.get(v, fill_zebra)
                    cell.font = sev_fonts.get(v, f_bold)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i in (1, 3, 5, 7, 8, 12):
                    cell.font = f_bold if c_i in (1, 3) else f_norm
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if idx % 2 == 1:
                        cell.fill = fill_zebra
                else:
                    cell.font = f_norm
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    if idx % 2 == 1:
                        cell.fill = fill_zebra

        ws.auto_filter.ref = f"A3:L{max(3, 3 + len(lista))}"
        ws.freeze_panes = "A4"

    # --- HOJA 2: CRÍTICOS Y ALTOS ---
    ws2 = wb.create_sheet("Críticos y Altos (Prioridad)")
    crit_alt = [h for h in hallazgos if h['severidad'] in ('CRÍTICO', 'ALTO')]
    llenar_tabla_hallazgos(ws2, crit_alt, f"HALLAZGOS CRÍTICOS Y ALTOS A CORREGIR ({len(crit_alt)} casos)")

    # --- HOJA 3: RESUMEN POR EPS ---
    ws3 = wb.create_sheet("Resumen por EPS")
    ws3.views.sheetView[0].showGridLines = True
    ws3.merge_cells("A1:G1")
    ws3["A1"] = "RANKING DE OBSERVACIONES POR EMPRESA PRESTADORA (50 EPS)"
    ws3["A1"].font = f_title
    ws3["A1"].fill = fill_navy
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[1].height = 28

    cols_eps = [
        ("Cód. EPS", 11), ("Empresa Prestadora", 34), ("Críticos", 12),
        ("Altos", 12), ("Medios", 12), ("Total Datos", 14), ("Enlaces Faltantes", 18)
    ]
    ws3.row_dimensions[3].height = 24
    for c_i, (h_txt, w) in enumerate(cols_eps, 1):
        cell = ws3.cell(3, c_i, h_txt)
        cell.font = f_hdr
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
        ws3.column_dimensions[get_column_letter(c_i)].width = w

    eps_sorted = sorted(
        eps_stats.values(),
        key=lambda x: (-x['criticos'], -x['altos'], -x['medios'], -x['links_faltantes'], x['cod_eps'])
    )
    for idx, st in enumerate(eps_sorted, 4):
        ws3.row_dimensions[idx].height = 20
        tot_d = st['criticos'] + st['altos'] + st['medios']
        vals = [st['cod_eps'], st['nombre_eps'], st['criticos'], st['altos'], st['medios'], tot_d, st['links_faltantes']]
        for c_i, v in enumerate(vals, 1):
            cell = ws3.cell(idx, c_i, v)
            cell.border = border_all
            cell.alignment = Alignment(horizontal="left" if c_i == 2 else "center", vertical="center")
            if c_i == 3 and v > 0:
                cell.fill = sev_fills['CRÍTICO']
                cell.font = sev_fonts['CRÍTICO']
            elif c_i == 4 and v > 0:
                cell.fill = sev_fills['ALTO']
                cell.font = sev_fonts['ALTO']
            elif c_i == 5 and v > 0:
                cell.fill = sev_fills['MEDIO']
                cell.font = sev_fonts['MEDIO']
            else:
                cell.font = f_bold if c_i in (1, 6) else f_norm
                if idx % 2 == 1:
                    cell.fill = fill_zebra
    ws3.auto_filter.ref = f"A3:G{max(3, 3 + len(eps_sorted))}"
    ws3.freeze_panes = "A4"

    # --- HOJA 4: TODOS LOS HALLAZGOS ---
    ws4 = wb.create_sheet("Todos los Hallazgos")
    llenar_tabla_hallazgos(ws4, hallazgos, f"LISTADO COMPLETO DE HALLAZGOS DE DATOS ({len(hallazgos)} casos)")

    # --- HOJA 5: ENLACES FALTANTES ---
    ws5 = wb.create_sheet("Enlaces Faltantes (Links)")
    ws5.views.sheetView[0].showGridLines = True
    ws5.merge_cells("A1:H1")
    ws5["A1"] = f"DOCUMENTOS SIN HIPERVÍNCULO ACTIVO ({len(links_faltantes)} documentos)"
    ws5["A1"].font = f_title
    ws5["A1"].fill = fill_navy
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws5.row_dimensions[1].height = 28

    cols_lnk = [
        ("Fuente / Hoja", 24), ("Cód. EPS", 10), ("Empresa Prestadora", 30),
        ("Periodo", 11), ("Tipo de Servicio", 18), ("Año / Ref.", 14),
        ("Tipo de Documento", 24), ("N° Documento / Descripción", 42)
    ]
    ws5.row_dimensions[3].height = 24
    for c_i, (h_txt, w) in enumerate(cols_lnk, 1):
        cell = ws5.cell(3, c_i, h_txt)
        cell.font = f_hdr
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
        ws5.column_dimensions[get_column_letter(c_i)].width = w

    for idx, l in enumerate(links_faltantes, 4):
        ws5.row_dimensions[idx].height = 20
        vals = [l['fuente'], l['cod_eps'], l['nombre_eps'], l['periodo'], l['servicio'], l['anio'], l['tipo_doc'], l['documento']]
        for c_i, v in enumerate(vals, 1):
            cell = ws5.cell(idx, c_i, v)
            cell.border = border_all
            cell.font = f_norm
            cell.alignment = Alignment(horizontal="center" if c_i in (2, 4, 6) else "left", vertical="center")
            if idx % 2 == 1:
                cell.fill = fill_zebra
    ws5.auto_filter.ref = f"A3:H{max(3, 3 + len(links_faltantes))}"
    ws5.freeze_panes = "A4"

    wb.save(EXCEL_REPORTE)
    print(f"   -> Guardado: {EXCEL_REPORTE}")

    # =========================================================
    # GENERAR ARCHIVO JS PARA EL APLICATIVO WEB
    # =========================================================
    print(f"6. Generando archivo de alertas para el aplicativo web: {JS_SALIDA}...")
    resumen_js = {
        'fecha_auditoria': fecha_gen,
        'total_hallazgos': len(hallazgos),
        'criticos': cnt_sev['CRÍTICO'],
        'altos': cnt_sev['ALTO'],
        'medios': cnt_sev['MEDIO'],
        'links_faltantes': len(links_faltantes),
        'eps_con_observaciones': sum(1 for st in eps_stats.values() if (st['criticos'] + st['altos'] + st['medios']) > 0),
        'por_categoria': [
            {'severidad': sev, 'categoria': cat, 'cantidad': cnt}
            for (sev, cat), cnt in sorted(cnt_cat.items(), key=lambda x: (sev_order.get(x[0][0], 9), -x[1]))
        ],
        'resumen_eps': eps_sorted
    }

    with open(JS_SALIDA, "w", encoding="utf-8") as f:
        f.write("// Resultados de Auditoría de Calidad de Datos generados automáticamente por generar_auditoria.py\n")
        f.write("const auditoriaResumen = ")
        json.dump(resumen_js, f, ensure_ascii=False, separators=(',', ':'))
        f.write(";\n\n")
        f.write("const auditoriaHallazgos = ")
        json.dump(hallazgos, f, ensure_ascii=False, separators=(',', ':'))
        f.write(";\n")

    print(f"   -> Guardado: {JS_SALIDA}")
    print("\n=== RESUMEN FINAL ===")
    print(f"CRÍTICOS: {cnt_sev['CRÍTICO']}")
    print(f"ALTOS:    {cnt_sev['ALTO']}")
    print(f"MEDIOS:   {cnt_sev['MEDIO']}")
    print(f"ENLACES FALTANTES (Informes + ET + RCD): {len(links_faltantes)}")
    print(f"EPS con al menos 1 observación de datos: {resumen_js['eps_con_observaciones']} / {len(eps_stats)}")

if __name__ == "__main__":
    ejecutar_auditoria()
