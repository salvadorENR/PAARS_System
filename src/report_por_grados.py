# src/report_por_grados.py
import os
import pandas as pd
import re
import sys
import glob
import unicodedata

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def limpiar_textos_redaccion(texto):
    """Limpia ortografía, caracteres corruptos de Unicode y miles con signos de interrogación."""
    if pd.isna(texto): return ""
    t = str(texto)
    
    # 1. Correcciones ortográficas y de símbolos Unicode corruptos
    reemplazos = {
        'à': 'á', 'è': 'é', 'ì': 'í', 'ò': 'ó', 'ù': 'ú',
        'À': 'Á', 'È': 'É', 'Ì': 'Í', 'Ò': 'Ó', 'Ù': 'Ú',
        '\ufffd': ' ', 
        ' ,': ',', ' .': '.'
    }
    for mal, bien in reemplazos.items():
        t = t.replace(mal, bien)
        
    # 2. Filtro quirúrgico para los miles corruptos (ej. 1,?000,?000)
    t = re.sub(r',\?(\d+)', r',\1', t)
    t = re.sub(r'(?<=\d)\?(\d+)', r',\1', t)
        
    # 3. Eliminar espacios dobles
    t = re.sub(r'\s+', ' ', t)
    
    return t.strip()

def normalize_item_code(code_str):
    """Strip everything except uppercase letters and digits. e.g. 'Lec 20' -> 'LEC20'."""
    if pd.isna(code_str):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).upper())


def clasificar_estricto(pct):
    """Classification rules: 0-20, 21-40, 41-60, 61-80, 81-100."""
    if pd.isna(pct):
        return "Crítico", "#1A2744", "#FFFFFF"
    if pct <= 20:
        return "Crítico", "#1A2744", "#FFFFFF"
    elif pct <= 40:
        return "Bajo", "#1E4D8C", "#FFFFFF"
    elif pct <= 60:
        return "Medio", "#2979C4", "#FFFFFF"
    elif pct <= 80:
        return "Bueno", "#5BAAE8", "#000000"
    else:
        return "Excelente", "#B8DEFF", "#000000"


def extraer_numero_grado(grado_str):
    if pd.isna(grado_str):
        return 0
    match = re.search(r'\d+', str(grado_str))
    return int(match.group()) if match else 0


# ---------------------------------------------------------------------------
# FIX 1 – robust LEC fingerprint (ported from report_profesores.py)
# ---------------------------------------------------------------------------

def make_lec_fingerprint(texto):
    """
    Stable grouping key for a LEC indicator text.
    Priority: objective code (e.g. '4.3') > normalised text.
    Multiple LEC items that share the same x.y objective code are merged
    into one group regardless of minor wording differences.
    """
    if not texto or texto.startswith('Indicador no disponible'):
        return f'missing_{texto}'          # keep each "missing" row separate
    # Extract objective code: digit(s).digit(s), e.g. "4.3", "1.10"
    code_match = re.search(r'\b(\d+\.\d+)\b', texto)
    if code_match:
        return f"obj_{code_match.group(1)}"
    # Fallback: strip accents + non-alphanumeric + whitespace, cap length
    norm = unicodedata.normalize('NFD', texto.lower())
    norm = ''.join(c for c in norm if unicodedata.category(c) != 'Mn')
    norm = re.sub(r'[\W_]+', '', norm)
    return norm[:60]


# ---------------------------------------------------------------------------
# FIX 2 – anchor-item loading (ported from report_profesores.py)
# ---------------------------------------------------------------------------

def load_anchor_items():
    """
    Loads the anchor-item exclusion list for the current exam month.

    File: items_anclaje_progreso_mes_N.xlsx  (in config.PATH_METADATA)
    Sheets: 'Matematica' and 'Lectura' (all sheets are read automatically).
    Columns: '2deg', '3deg', ... '11deg'
    Rows: item codes to exclude for that grade (e.g. MAT211, LEC315).
    Empty cells are ignored.

    Returns dict { numeric_grade_str : set_of_normalised_codes }
    e.g. { '2': {'MAT211','MAT208','LEC315',...}, '3': {...}, ... }
    Returns empty dict if the file is not found -> no items excluded.
    """
    anchor_by_grade = {}

    match = re.match(r'^(\d+)_', os.path.basename(config.MONTH_FOLDER))
    mes_num = str(int(match.group(1))) if match else re.sub(r'\D', '', config.MONTH_FOLDER) or '1'

    expected_name = f"items_anclaje_progreso_mes_{mes_num}.xlsx"
    anchor_path   = os.path.join(config.PATH_METADATA, expected_name)

    if not os.path.exists(anchor_path):
        pattern    = os.path.join(config.PATH_METADATA, f"items_anclaje*mes*{mes_num}*.xlsx")
        candidates = [f for f in glob.glob(pattern)
                      if not os.path.basename(f).startswith('~$')]
        if candidates:
            anchor_path = candidates[0]
        else:
            print(f"  [i] No se encontró el archivo de ítems ancla ({expected_name}). "
                  f"No se excluirá ningún ítem ancla.")
            return anchor_by_grade

    try:
        xl_anchor = pd.ExcelFile(anchor_path)
        for sheet_name in xl_anchor.sheet_names:
            df_sheet = xl_anchor.parse(sheet_name, dtype=str)
            for grade_col in df_sheet.columns:
                grade_num_key = re.sub(r'\D', '', str(grade_col).strip())
                if not grade_num_key:
                    continue
                if grade_num_key not in anchor_by_grade:
                    anchor_by_grade[grade_num_key] = set()
                for val in df_sheet[grade_col].dropna():
                    code = normalize_item_code(str(val))
                    if code:
                        anchor_by_grade[grade_num_key].add(code)

        total = sum(len(v) for v in anchor_by_grade.values())
        print(f"  -> Ítems ancla cargados ({os.path.basename(anchor_path)}): "
              f"{total} ítem(s) en {len(anchor_by_grade)} grado(s).")
    except Exception as e:
        print(f"  [!] Error leyendo archivo de ítems ancla: {e}")

    return anchor_by_grade


# ---------------------------------------------------------------------------
# FIX 3 – robust metadata loader
# ---------------------------------------------------------------------------

def cargar_metadata_indicadores():
    """
    Reads the maestro/procesado Excel and returns a dict
    { normalised_item_code : { 'objetivo', 'clases', 'orden' } }.

    Fixes applied vs. original:
      - Zero-padded month matching ('mes03' as well as 'mes3').
      - Uses normalize_item_code() for consistent key building.
      - Also loads Clases_Sugeridas and Orden_Clase for badge rendering.
      - Falls back to the first available metadata file if no specific
        match is found, so the report degrades gracefully.
      - APLICA LIMPIEZA DE REDACCIÓN AL TEXTO DE LOS INDICADORES.
    """
    item_dict = {}
    archivos_meta = (glob.glob(os.path.join(config.PATH_METADATA, "*.xlsx")) +
                     glob.glob(os.path.join(config.PATH_METADATA, "*.csv")))
    archivos_meta = [f for f in archivos_meta
                     if not os.path.basename(f).startswith('~$')]

    match_mes  = re.search(r'^0?(\d+)_', config.MONTH_FOLDER)
    mes_num    = match_mes.group(1) if match_mes else "1"
    mes_padded = mes_num.zfill(2)

    exam_lower = config.EXAM_TYPE.lower()

    archivo_maestro = None
    for f in archivos_meta:
        nombre = os.path.basename(f).lower()
        # Must contain the exam type (e.g. 'progreso')
        if exam_lower not in nombre:
            continue
        # Must match the month (plain or zero-padded, with or without separator)
        month_patterns = [
            f"mes{mes_num}", f"mes{mes_padded}",
            f"mes {mes_num}", f"mes {mes_padded}",
            f"mes_{mes_num}", f"mes_{mes_padded}",
            f"_{mes_num}_",  f"_{mes_padded}_",
        ]
        if not any(p in nombre for p in month_patterns):
            continue
        # Prefer files that look like the processed master
        if 'procesado' in nombre or 'maestro' in nombre:
            archivo_maestro = f
            break
        # Accept any match as a fallback candidate
        if archivo_maestro is None:
            archivo_maestro = f

    # Last-resort: use the first available metadata file
    if not archivo_maestro and archivos_meta:
        archivo_maestro = archivos_meta[0]
        print(f"  [!] No se encontró maestro exacto para mes {mes_num}. "
              f"Usando: {os.path.basename(archivo_maestro)}")

    if not archivo_maestro:
        print("  [!] No hay archivos de metadatos en PATH_METADATA.")
        return item_dict

    print(f"  -> Usando maestro de indicadores: {os.path.basename(archivo_maestro)}")

    try:
        if archivo_maestro.endswith('.csv'):
            df_meta = pd.read_csv(archivo_maestro, encoding='utf-8-sig', encoding_errors='replace',
                                  sep=None, engine='python', dtype=str)
        else:
            df_meta = pd.read_excel(archivo_maestro, dtype=str)

        df_meta.columns = df_meta.columns.str.strip()

        col_item   = next((c for c in df_meta.columns
                           if 'item' in c.lower() or 'código' in c.lower()), None)
        col_ind    = next((c for c in df_meta.columns
                           if 'indicador' in c.lower()), None)
        col_clases = next((c for c in df_meta.columns
                           if 'clases' in c.lower() or 'clase' in c.lower()), None)
        col_orden  = next((c for c in df_meta.columns
                           if 'orden' in c.lower()), None)

        if not col_item or not col_ind:
            print(f"  [!] No se encontraron columnas de ítem/indicador en "
                  f"{os.path.basename(archivo_maestro)}. "
                  f"Columnas disponibles: {list(df_meta.columns)}")
            return item_dict

        for _, row in df_meta.dropna(subset=[col_item, col_ind]).iterrows():
            codigo = normalize_item_code(row[col_item])
            if not codigo:
                continue
            
            # --- LIMPIEZA DE REDACCIÓN INYECTADA AQUÍ ---
            objetivo_limpio = limpiar_textos_redaccion(row[col_ind])
            
            clases_val = ""
            if col_clases and pd.notna(row[col_clases]) and str(row[col_clases]).strip() not in ('', 'nan'):
                clases_val = str(row[col_clases]).strip()
            orden_val = 9999.0
            if col_orden and pd.notna(row[col_orden]) and str(row[col_orden]).strip() not in ('', 'nan'):
                try:
                    orden_val = float(row[col_orden])
                except ValueError:
                    pass
            item_dict[codigo] = {
                'objetivo': objetivo_limpio,
                'clases':   clases_val,
                'orden':    orden_val,
            }

        print(f"  -> {len(item_dict)} indicadores cargados.")
    except Exception as e:
        print(f"  [!] Error procesando metadatos: {e}")

    return item_dict


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------

def cargar_datos_directo_resultados():
    ruta    = os.path.join(config.PATH_INTERIM, "Resultados")
    archivos = glob.glob(os.path.join(ruta, "*.csv"))
    dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            df['Area temática'] = ('Matemática'
                                   if 'MAT' in os.path.basename(f).upper()
                                   else 'Lengua')
            dfs.append(df)
        except Exception as e:
            print(f" [!] Error leyendo {f}: {e}")
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# ---------------------------------------------------------------------------
# MAIN REPORT GENERATOR
# ---------------------------------------------------------------------------

def generar_reporte_por_grados(df_master=None):
    print("\n[*] Generando Reporte Nacional por Grados (Dashboard)...")

    df = cargar_datos_directo_resultados()
    if df.empty:
        print("  [!] No hay datos en la carpeta Resultados.")
        return

    # Remove annulled tests
    if 'anular_prueba' in df.columns:
        df = df[df['anular_prueba'].isna() |
                (df['anular_prueba'].astype(str).str.strip() == '')]

    col_doc   = next((c for c in df.columns
                      if 'documento' in str(c).lower() or 'nie' in str(c).lower()),
                     'Documento')
    col_grado = next((c for c in df.columns
                      if 'grado' in str(c).lower()), 'Grado')
    df['Grado_Num'] = df[col_grado].apply(extraer_numero_grado)

    # Title & date
    match_mes  = re.search(r'^0?(\d+)_', config.MONTH_FOLDER)
    mes_n      = match_mes.group(1) if match_mes else "X"
    titulo     = f"Resultados de la Prueba de {config.EXAM_TYPE.capitalize()} Mes {mes_n}"

    col_fecha   = next((c for c in df.columns if 'fecha' in str(c).lower()), None)
    fecha_texto = "No especificada"
    if col_fecha:
        fechas_validas = pd.to_datetime(
            df[col_fecha].replace('Sin Especificar', pd.NA),
            errors='coerce', dayfirst=True)
        if not fechas_validas.dropna().empty:
            meses_es  = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril',
                         5:'Mayo',  6:'Junio',   7:'Julio', 8:'Agosto',
                         9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
            mes_moda  = fechas_validas.dropna().dt.month.mode().iloc[0]
            anio_moda = fechas_validas.dropna().dt.year.mode().iloc[0]
            fecha_texto = f"{meses_es.get(int(mes_moda), '')} {int(anio_moda)}"

    # Load indicator metadata and anchor items ONCE
    mapa_indicadores = cargar_metadata_indicadores()
    anchor_by_grade  = load_anchor_items()

    # -----------------------------------------------------------------------
    # Per-subject HTML generation
    # -----------------------------------------------------------------------
    html_sections = {}

    for materia in ['Matemática', 'Lengua']:
        df_m = df[df['Area temática'] == materia].copy()
        if df_m.empty:
            continue

        prefijo = 'MAT' if materia == 'Matemática' else 'LEC'

        # Summary table
        conteo     = df_m.groupby('Grado_Num')[col_doc].nunique().reset_index()
        tabla_html = (f"<div class='card'><h2>👥 Resumen de Estudiantes - {materia}</h2>"
                      f"<table><thead><tr><th>Grado</th>"
                      f"<th>Estudiantes Evaluados</th></tr></thead><tbody>")
        for _, r in conteo.sort_values('Grado_Num').iterrows():
            if r['Grado_Num'] > 0:
                tabla_html += (f"<tr><td>{int(r['Grado_Num'])}° Grado</td>"
                               f"<td>{int(r[col_doc])}</td></tr>")
        tabla_html += "</tbody></table></div>"

        # Accordion per grade
        indicadores_html = ""
        grados = sorted([g for g in df_m['Grado_Num'].unique() if g > 0])

        for g in grados:
            df_g = df_m[df_m['Grado_Num'] == g].copy()

            # All item columns for this grade/subject
            item_cols_raw = [c for c in df_g.columns
                             if re.match(rf'^{prefijo}\d+$', str(c).upper())]

            # FIX 2 – remove anchor items
            grade_num_key    = str(g)
            anchor_for_grade = anchor_by_grade.get(grade_num_key, set())
            item_cols        = [c for c in item_cols_raw
                                  if normalize_item_code(c) not in anchor_for_grade]

            if not item_cols:
                continue

            # -------------------------------------------------------------------
            # FIX 3 – normalise item codes when looking up the indicator map
            # FIX 1 – use make_lec_fingerprint() for Lengua to merge duplicates
            # -------------------------------------------------------------------
            indicator_groups = {}

            for item in item_cols:
                # Normalised code for dict lookup (handles spaces, casing, etc.)
                codigo_puro = normalize_item_code(item)
                meta        = mapa_indicadores.get(codigo_puro, {})
                texto_ind   = meta.get('objetivo', f"Indicador no disponible ({item})")
                clases_ind  = meta.get('clases', '')
                orden_ind   = meta.get('orden', 9999.0)

                if materia == 'Lengua':
                    fingerprint = make_lec_fingerprint(texto_ind)
                else:
                    # For Matemática each item is its own indicator
                    fingerprint = codigo_puro

                if fingerprint not in indicator_groups:
                    indicator_groups[fingerprint] = {
                        'display_text': texto_ind,
                        'items':        [],
                        'clases':       set(),
                        'orden':        orden_ind,
                    }
                indicator_groups[fingerprint]['items'].append(item)
                if clases_ind:
                    indicator_groups[fingerprint]['clases'].add(clases_ind)
                # Keep the earliest orden for sorting LEC groups
                if orden_ind < indicator_groups[fingerprint]['orden']:
                    indicator_groups[fingerprint]['orden'] = orden_ind

            # Calculate % success per group
            lista_final = []
            for fp, data in indicator_groups.items():
                total_corr, total_pos = 0, 0
                for it in data['items']:
                    df_g[it] = pd.to_numeric(df_g[it], errors='coerce')
                    validos   = df_g[it].notna()
                    if validos.sum() > 0:
                        total_corr += (df_g.loc[validos, it] == 1).sum()
                        total_pos  += validos.sum()
                if total_pos > 0:
                    pct             = (total_corr / total_pos) * 100
                    cat, bg, text_color = clasificar_estricto(pct)
                    clases_str      = " | ".join(sorted(data['clases'])) if data['clases'] else ""
                    lista_final.append({
                        'texto':   data['display_text'],
                        'clases':  clases_str,
                        'pct':     pct,
                        'cat':     cat,
                        'bg':      bg,
                        'text_color': text_color,
                    })

            # Sort ascending (hardest first)
            lista_final.sort(key=lambda x: x['pct'])

            html_items = ""
            for ind in lista_final:
                badge = (f"<span class='badge-class'>{ind['clases']}</span> "
                         if ind['clases'] else "")
                html_items += f"""
                <div class="indicator-item" style="background-color: {ind['bg']}; color: {ind['text_color']}; border: none;">
                    <div class="pct-box" style="border-right: 1px solid rgba(150,150,150,0.3);">
                        <strong style="color: {ind['text_color']};">{ind['pct']:.1f}%</strong><br>
                        <small style="color: {ind['text_color']}; opacity: 0.9;">{ind['cat']}</small>
                    </div>
                    <div class="ind-text" style="color: {ind['text_color']};">{badge}{ind['texto']}</div>
                </div>"""

            indicadores_html += f"""
            <details class="grade-accordion">
                <summary>Grado {g}°: Mostrar/Ocultar Niveles de los Indicadores</summary>
                <div class="content">
                    {html_items if html_items else "<p>No hay indicadores evaluados en este grado.</p>"}
                </div>
            </details>"""

        html_sections[materia] = tabla_html + indicadores_html

    # -----------------------------------------------------------------------
    # Final HTML with tabs
    # -----------------------------------------------------------------------
    html_final = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>{titulo}</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f4f7f9; padding: 20px; }}
            .header {{ text-align: center; margin-bottom: 30px; background: white; padding: 20px; border-radius: 8px; border-bottom: 4px solid #2980b9; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            .header h1 {{ margin: 0; color: #2c3e50; font-size: 24px; }}
            .header h3 {{ margin: 5px 0 0 0; color: #7f8c8d; font-weight: normal; font-size: 14px; }}
            .legend {{ display: inline-block; background: #fff; padding: 10px 20px; border-radius: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-top: 15px; font-size: 13px; }}

            /* Tabs */
            .tabs {{ display: flex; gap: 10px; margin-bottom: 20px; justify-content: center; }}
            .tab-btn {{ padding: 12px 30px; cursor: pointer; background: #ddd; border: none; border-radius: 5px; font-weight: bold; font-size: 16px; transition: 0.3s; color: #333; }}
            .tab-btn.active {{ background: #2980b9; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .tab-btn:hover:not(.active) {{ background: #ccc; }}
            .tab-content {{ display: none; }}
            .tab-content.active {{ display: block; }}

            /* Accordion & Tables */
            .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; }}
            .card h2 {{ margin-top: 0; color: #2c3e50; font-size: 18px; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            .grade-accordion {{ margin-bottom: 15px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; background: white; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
            .grade-accordion summary {{ padding: 15px 20px; background: #f8f9fa; cursor: pointer; font-weight: bold; font-size: 16px; color: #2980b9; list-style: none; transition: background 0.2s; }}
            .grade-accordion summary:hover {{ background: #e9ecef; }}
            .grade-accordion summary::-webkit-details-marker {{ display: none; }}
            .grade-accordion .content {{ padding: 20px; border-top: 1px solid #eee; }}

            table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
            th, td {{ border: 1px solid #eee; padding: 12px; text-align: center; }}
            th {{ background: #2c3e50; color: white; }}

            /* Indicators */
            .indicator-item {{ display: flex; align-items: center; padding: 12px; margin-bottom: 10px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            .pct-box {{ width: 80px; text-align: center; padding-right: 15px; margin-right: 15px; flex-shrink: 0; }}
            .pct-box strong {{ font-size: 16px; color: inherit; }}
            .pct-box small {{ text-transform: uppercase; font-size: 10px; font-weight: bold; color: inherit; }}
            .ind-text {{ font-size: 14px; color: inherit; font-weight: 500; }}
            .badge-class {{ display: inline-block; background-color: #34495e; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; margin-right: 8px; font-weight: bold; letter-spacing: 0.5px; box-shadow: 0 1px 2px rgba(0,0,0,0.1); white-space: nowrap; vertical-align: middle; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{titulo}</h1>
            <h3>Fecha de Aplicación: {fecha_texto}</h3>
            <div class="legend"><strong>Cortes de Evaluación:</strong> Crítico (0-20) | Bajo (21-40) | Medio (41-60) | Bueno (61-80) | Excelente (81-100)</div>
        </div>

        <div class="tabs">
            <button class="tab-btn active" onclick="openTab(event, 'MAT')">📘 Matemática</button>
            <button class="tab-btn" onclick="openTab(event, 'LEC')">📕 Lengua</button>
        </div>

        <div id="MAT" class="tab-content active">{html_sections.get('Matemática', '<div class="card">No hay datos procesados para Matemática.</div>')}</div>
        <div id="LEC" class="tab-content">{html_sections.get('Lengua', '<div class="card">No hay datos procesados para Lengua.</div>')}</div>

        <script>
            function openTab(evt, tabName) {{
                var i, content, btn;
                content = document.getElementsByClassName("tab-content");
                for (i = 0; i < content.length; i++) content[i].style.display = "none";
                btn = document.getElementsByClassName("tab-btn");
                for (i = 0; i < btn.length; i++) btn[i].className = btn[i].className.replace(" active", "");
                document.getElementById(tabName).style.display = "block";
                evt.currentTarget.className += " active";
            }}
        </script>
    </body>
    </html>
    """

    out_dir  = os.path.join(config.PATH_REPORTS, "00_Reportes_Nacionales_Agregados")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"Reporte_Nacional_Grados_{config.MONTH_FOLDER}.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_final)
    print(f"  [OK] Reporte generado exitosamente en:\n       {out_path}")