# src/report_profesores.py
import os
import re
import pandas as pd
import numpy as np
import sys
import glob
import unicodedata

# --- 1. CONEXIÓN CON EL MAPA PAARS ---
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# --- 2. PLANTILLA HTML Y CSS ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Resultados - {grupo}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <style>
        * {{
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }}
        
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }}
        
        /* Controles y Botones */
        .controls {{ display: flex; justify-content: flex-end; gap: 10px; margin-bottom: 20px; }}
        .btn {{ padding: 10px 15px; border: 1px solid #bdc3c7; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold; transition: all 0.2s; }}
        .btn-default {{ background: #ffffff; color: #2c3e50; }}
        .btn-default:hover {{ background: #ecf0f1; }}
        .btn-primary {{ background: #2980b9; color: white; border-color: #2471a3; }}
        .btn-primary:hover {{ background: #1f618d; }}
        
        .header-card {{ background-color: #ffffff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; border-left: 5px solid #2980b9; }}
        .header-card h2 {{ margin: 0 0 15px 0; color: #2c3e50; font-size: 24px; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; }}
        .info-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; font-size: 14px; color: #444; }}
        .info-item strong {{ color: #2c3e50; display: block; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
        .info-item span {{ font-size: 15px; font-weight: 500; }}
        
        .subject-section {{ margin-bottom: 50px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
        .subject-title {{ font-size: 20px; color: #2980b9; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; margin-bottom: 15px; text-transform: uppercase; font-weight: bold; }}
        
        .table-container {{ max-width: 100%; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 15px; }}
        table {{ border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13px; }}
        th, td {{ padding: 10px 15px; text-align: center; border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; }}
        
        th {{ background-color: #34495e; color: white; font-weight: 600; position: sticky; top: 0; z-index: 10; border-bottom: 2px solid #2c3e50; }}
        
        th:nth-child(2), td:nth-child(2) {{ position: sticky; left: 0; background-color: #fdfdfd; z-index: 5; text-align: left; border-right: 2px solid #bdc3c7; font-weight: bold; box-shadow: 2px 0 5px rgba(0,0,0,0.05); vertical-align: middle; }}
        th:nth-child(2) {{ background-color: #2c3e50; color: white; z-index: 15; }}
        
        .total-cell {{ background-color: #ecf0f1; font-weight: bold; font-size: 14px; color: #333; }}
        
        tbody tr:hover td {{ background-color: rgba(0,0,0,0.03); }}
        tbody tr:hover td:nth-child(2) {{ filter: brightness(0.95); }}

        .indicator-section {{ padding: 20px; background-color: #f8fcfd; border-radius: 6px; border: 1px solid #e1ecef; border-left: 4px solid #1abc9c; margin-bottom: 25px; box-sizing: border-box; }}
        
        .indicator-list {{ display: flex; flex-direction: column; gap: 8px; font-size: 13px; color: #333; }}
        .indicator-item {{ display: flex; align-items: center; padding: 10px 12px; border-radius: 4px; border: 1px solid rgba(0,0,0,0.1); }}
        
        details.info-accordion {{ margin-bottom: 15px; text-align: left; display: block; }}
        details.info-accordion summary {{ cursor: pointer; font-size: 13px; font-weight: bold; color: #2980b9; padding: 8px 12px; background-color: #f4f6f9; border-radius: 4px; border: 1px solid #d0e1f9; display: inline-block; list-style: none; transition: 0.2s; }}
        details.info-accordion summary::-webkit-details-marker {{ display: none; }}
        details.info-accordion summary:hover {{ background-color: #eaf2f8; }}
        details.info-accordion p {{ padding: 15px; border-left: 3px solid #2980b9; background-color: #fafafa; margin-top: 8px; font-size: 13.5px; color: #444; border-radius: 0 4px 4px 0; box-shadow: inset 0 1px 3px rgba(0,0,0,0.03); line-height: 1.5; margin-bottom: 0; }}
        
        /* Paleta de Colores de Desempeño Unificada */
        .cat-critico   {{ background-color: #1A2744 !important; color: #FFFFFF !important; border: none !important; }}
        .cat-bajo      {{ background-color: #1E4D8C !important; color: #FFFFFF !important; border: none !important; }}
        .cat-medio     {{ background-color: #2979C4 !important; color: #FFFFFF !important; border: none !important; }}
        .cat-bueno     {{ background-color: #5BAAE8 !important; color: #000000 !important; border: none !important; }}
        .cat-excelente {{ background-color: #B8DEFF !important; color: #000000 !important; border: none !important; }}

        .indicator-pct {{ font-weight: bold; width: 65px; text-align: center; margin-right: 15px; padding-right: 15px; border-right: 1px solid currentColor; flex-shrink: 0; }}
        .pct-number {{ font-size: 16px; color: inherit; font-weight: bold; }}
        .pct-label {{ font-size: 10px; color: inherit; opacity: 0.8; text-transform: uppercase; letter-spacing: 0.5px; }}
        
        .item-label {{ color: inherit; display: inline-block; width: 55px; font-weight: bold; }}
        
        .badge-class {{ display: inline-block; background-color: #34495e; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; margin-right: 10px; vertical-align: baseline; font-weight: bold; letter-spacing: 0.5px; box-shadow: 0 1px 2px rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.2); }}
        .nivelacion-badge {{ font-weight: bold; margin-left: 5px; font-size: 12px; }}
        
        .footnote {{ font-size: 13px; color: #333; margin-bottom: 10px; padding: 10px 15px; border-radius: 4px; }}
        .footnote-warning {{ background-color: #fdf2f2; border-left: 4px solid #e74c3c; color: #c0392b; }}
        
        .badge-container {{ display: flex; flex-direction: column; align-items: center; justify-content: center; }}
        .donut-seal {{ width: 95px; height: 95px; border-radius: 50%; border: 14px solid var(--badge-color); display: flex; align-items: center; justify-content: center; font-size: 28px; font-weight: 900; color: #2c3e50; background: #fff; box-shadow: inset 0 2px 4px rgba(0,0,0,0.1); }}
        .donut-label {{ margin-top: 15px; font-size: 16px; font-weight: bold; color: var(--badge-color); text-transform: uppercase; text-align: center; background: rgba(0,0,0,0.03); padding: 6px 15px; border-radius: 15px; border: 1px solid var(--badge-color); }}
        
        @media print {{
            .no-print {{ display: none !important; }}
            body {{ background-color: white; padding: 0; }}
            .header-card, .subject-section {{ box-shadow: none; border: 1px solid #ddd; break-inside: avoid; }}
            .table-container {{ max-height: none !important; overflow: visible !important; page-break-inside: auto; }}
            tr {{ page-break-inside: avoid; }}
            .chart-wrapper-print {{ width: 300px !important; height: 300px !important; margin: 0 auto; }}
            canvas {{ max-width: 300px !important; max-height: 300px !important; }}
        }}
    </style>
</head>
<body>

    <div class="controls no-print">
        <button onclick="window.print()" class="btn btn-primary">Descargar como PDF</button>
    </div>

    <div class="header-card">
        <div>
            <h2>Resultados de la Prueba de Progreso - Mes {mes_prueba}, 2026</h2>
            <div class="info-grid">
                <div class="info-item"><strong>Departamento</strong><span>{departamento}</span></div>
                <div class="info-item"><strong>Código de Infraestructura</strong><span>{nro_centro}</span></div>
                <div class="info-item"><strong>Centro Escolar</strong><span>{centro}</span></div>
                <div class="info-item"><strong>Grado</strong><span>{grado}</span></div>
                <div class="info-item"><strong>Grupo</strong><span>{grupo}</span></div>
                <div class="info-item"><strong>Mes de Aplicación</strong><span>{mes_aplicacion}</span></div>
            </div>
        </div>
    </div>

    {tablas_html}

    <script>
        function toggleSort(btn, tbodyId) {{
            var tbody = document.getElementById(tbodyId);
            var rows = Array.from(tbody.querySelectorAll('tr'));
            
            var currentState = btn.getAttribute('data-sort-state') || 'rendimiento';
            
            if (currentState === 'rendimiento') {{
                rows.sort(function(a, b) {{
                    var nameA = a.getAttribute('data-nombre') || '';
                    var nameB = b.getAttribute('data-nombre') || '';
                    return nameA.localeCompare(nameB, 'es', {{sensitivity: 'base'}});
                }});
                btn.setAttribute('data-sort-state', 'alfabetico');
                btn.innerHTML = 'Ordenar los estudiantes de la tabla por Nivel en la prueba';
            }} else {{
                rows.sort(function(a, b) {{
                    var scoreA = parseFloat(a.getAttribute('data-puntaje'));
                    var scoreB = parseFloat(b.getAttribute('data-puntaje'));
                    return scoreA - scoreB;
                }});
                btn.setAttribute('data-sort-state', 'rendimiento');
                btn.innerHTML = 'Ordenar los estudiantes de la tabla por orden alfabético';
            }}
            
            rows.forEach(function(row) {{ tbody.appendChild(row); }});
        }}

        function toggleIndicatorSort(btn, listId, titleId) {{
            var list = document.getElementById(listId);
            var title = document.getElementById(titleId);
            var items = Array.from(list.querySelectorAll('.indicator-item'));
            
            var currentState = btn.getAttribute('data-sort-state') || 'pct';
            
            if (currentState === 'pct') {{
                items.sort(function(a, b) {{
                    var classA = parseFloat(a.getAttribute('data-class'));
                    var classB = parseFloat(b.getAttribute('data-class'));
                    if (classA === classB) {{
                        return parseFloat(a.getAttribute('data-pct')) - parseFloat(b.getAttribute('data-pct'));
                    }}
                    return classA - classB;
                }});
                btn.setAttribute('data-sort-state', 'class');
                btn.innerHTML = 'Ordenar por porcentaje de alcance';
                title.innerHTML = 'Porcentaje de alcance del indicador de logro (Por orden cronológico de clases)';
            }} else {{
                items.sort(function(a, b) {{
                    var pctA = parseFloat(a.getAttribute('data-pct'));
                    var pctB = parseFloat(b.getAttribute('data-pct'));
                    return pctA - pctB;
                }});
                btn.setAttribute('data-sort-state', 'pct');
                btn.innerHTML = 'Ordenar por orden cronológico de clases';
                title.innerHTML = 'Porcentaje de alcance del indicador de logro (De menor a mayor alcance)';
            }}
            
            items.forEach(function(item) {{ list.appendChild(item); }});
        }}

        Chart.register(ChartDataLabels);
        document.addEventListener('DOMContentLoaded', function() {{
            {chart_scripts}
        }});
    </script>
</body>
</html>
"""

# --- 3. FUNCIONES AUXILIARES ---

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
    # Si hay una coma seguida de '?' y números, elimina el '?': "1,?000" -> "1,000"
    t = re.sub(r',\?(\d+)', r',\1', t)
    # Si hay un '?' atrapado directamente entre números, lo vuelve coma: "1?000" -> "1,000"
    t = re.sub(r'(?<=\d)\?(\d+)', r',\1', t)
        
    # 3. Eliminar espacios dobles
    t = re.sub(r'\s+', ' ', t)
    
    return t.strip()


def clasificar_puntaje(val):
    if pd.isna(val): return "Crítico"
    val = float(val)
    if val <= 35: return "Crítico"
    elif val <= 45: return "Bajo"
    elif val <= 55: return "Medio"
    elif val <= 65: return "Bueno"
    else: return "Excelente"

def normalize_item_code(code_str):
    """Limpia el codigo quitando espacios, guiones, etc., y forzando mayusculas. Ej: 'Lec 20' -> 'LEC20'"""
    if pd.isna(code_str): return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).upper())

def load_item_indicators():
    """Lee dinamicamente el archivo de items asegurando que el Mes coincida exactamente."""
    item_base_dict = {}
    
    match = re.match(r'^(\d+)_([A-Z]+)', os.path.basename(config.MONTH_FOLDER))
    if match:
        exam_num = str(int(match.group(1))) 
        exam_type = match.group(2).capitalize()
    else:
        exam_num = re.sub(r'\D', '', config.MONTH_FOLDER)
        if not exam_num: exam_num = '1'
        exam_type = 'Progreso' if 'PROGRESO' in config.MONTH_FOLDER.upper() else 'Resultados'

    if exam_type == 'Resultado': exam_type = 'Resultados'
    
    todos_archivos = glob.glob(os.path.join(config.PATH_METADATA, "*.xlsx"))
    archivos_procesados = []
    
    for f in todos_archivos:
        nombre_archivo = os.path.basename(f)
        if nombre_archivo.startswith('~$'):
            continue
        if 'procesado' in nombre_archivo.lower() and exam_type.lower() in nombre_archivo.lower():
            if re.search(rf'Mes\s*0?{exam_num}(?:\D|$)', nombre_archivo, re.IGNORECASE) or \
               re.search(rf'_{exam_num}_', nombre_archivo):
                archivos_procesados.append(f)

    if not archivos_procesados:
        print(f"  [!] No se encontro el Excel maestro de items procesado para {exam_type} {exam_num}.")
        return item_base_dict

    ITEMS_FILE = archivos_procesados[0]
    print(f"  -> Usando maestro de items integrado: {os.path.basename(ITEMS_FILE)}")

    try:
        df_items = pd.read_excel(ITEMS_FILE, dtype=str)
        columnas_requeridas = ['ItemCodigo', 'indicador_logro', 'Grado_Aplicado',
                               'Clases_Sugeridas', 'Orden_Clase', 'Nivelacion']
        for col in columnas_requeridas:
            if col not in df_items.columns:
                df_items[col] = ''
                
        for _, row in df_items.dropna(subset=['ItemCodigo', 'indicador_logro']).iterrows():
            codigo_puro = normalize_item_code(row['ItemCodigo'])
            if not codigo_puro: continue
            orden_val = 9999.0
            if pd.notna(row['Orden_Clase']) and str(row['Orden_Clase']).strip() != '':
                try:
                    orden_val = float(row['Orden_Clase'])
                except:
                    pass
            
            # --- APLICAMOS LA LIMPIEZA DE REDACCIÓN AL INDICADOR ---
            objetivo_limpio = limpiar_textos_redaccion(row['indicador_logro'])
            
            item_base_dict[codigo_puro] = {
                'objetivo':   objetivo_limpio,
                'clases':     str(row['Clases_Sugeridas']).strip()
                              if pd.notna(row['Clases_Sugeridas'])
                              and str(row['Clases_Sugeridas']).strip() != 'nan' else "",
                'orden':      orden_val,
                'nivelacion': str(row['Nivelacion']).strip()
                              if pd.notna(row['Nivelacion'])
                              and str(row['Nivelacion']).strip() != 'nan' else ""
            }
    except Exception as e:
        print(f"  [!] Error procesando items maestros: {e}")
    return item_base_dict


def load_anchor_items():
    METADATA_DIR = config.PATH_METADATA
    anchor_by_grade = {}

    match = re.match(r'^(\d+)_', os.path.basename(config.MONTH_FOLDER))
    mes_num = str(int(match.group(1))) if match else re.sub(r'\D', '', config.MONTH_FOLDER) or '1'

    expected_name = f"items_anclaje_progreso_mes_{mes_num}.xlsx"
    anchor_path   = os.path.join(METADATA_DIR, expected_name)

    if not os.path.exists(anchor_path):
        pattern    = os.path.join(METADATA_DIR, f"items_anclaje*mes*{mes_num}*.xlsx")
        candidates = [f for f in glob.glob(pattern)
                      if not os.path.basename(f).startswith('~$')]
        if candidates:
            anchor_path = candidates[0]
        else:
            print(f"  [i] No se encontro el archivo de items ancla ({expected_name}). "
                  f"No se excluira ningun item ancla.")
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
        print(f"  -> Items ancla cargados ({os.path.basename(anchor_path)}): "
              f"{total} item(s) en {len(anchor_by_grade)} grado(s) "
              f"[hojas: {xl_anchor.sheet_names}].")
        for g, codes in sorted(anchor_by_grade.items(),
                               key=lambda x: int(x[0]) if x[0].isdigit() else 99):
            mat_codes = sorted(c for c in codes if c.startswith('MAT'))
            lec_codes = sorted(c for c in codes if c.startswith('LEC'))
            print(f"       Grado {g}: MAT={mat_codes}  LEC={lec_codes}")

    except Exception as e:
        print(f"  [!] Error leyendo archivo de items ancla: {e}")

    return anchor_by_grade

def build_master_geiser_dataframe():
    geiser_csv_dir_sub = os.path.join(config.PATH_INTERIM, "Resultados")
    GEISER_CSV_DIR = (geiser_csv_dir_sub
                      if os.path.exists(geiser_csv_dir_sub)
                      else config.PATH_INTERIM)

    if not os.path.exists(GEISER_CSV_DIR):
        print(f"  [!] Directorio no encontrado: {GEISER_CSV_DIR}")
        return pd.DataFrame()

    all_csvs = [f for f in os.listdir(GEISER_CSV_DIR)
                if f.lower().endswith('.csv') and 'legend' not in f.lower()]

    print(f"  [INFO] Leyendo CSVs desde: {GEISER_CSV_DIR}")
    print(f"  [INFO] Archivos CSV encontrados: {len(all_csvs)}")

    res_dfs = []

    for f in sorted(all_csvs):
        if '_compl' in f.lower():
            continue

        path = os.path.join(GEISER_CSV_DIR, f)
        try:
            df = pd.read_csv(path, dtype=str, encoding='utf-8-sig')
            df.columns = df.columns.str.strip()
            df = df.loc[:, ~df.columns.duplicated()].copy()

            if 'anular_prueba' not in df.columns:
                col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
                if col_anular:
                    df = df.rename(columns={col_anular: 'anular_prueba'})

            if 'theta.global (escala 0-100)' not in df.columns:
                col_theta = next((c for c in df.columns
                                  if 'theta' in c.lower() and 'global' in c.lower()), None)
                if col_theta:
                    df = df.rename(columns={col_theta: 'theta.global (escala 0-100)'})

            df = df.loc[:, ~df.columns.duplicated()].copy()

            if 'MAT-' in f.upper():
                df['Area temática'] = 'Matemática'
            elif 'LEC-' in f.upper():
                df['Area temática'] = 'Lectura'
            elif 'Area temática' not in df.columns:
                df['Area temática'] = 'Desconocida'

            res_dfs.append(df)
            print(f"  [OK] Leído: {f}  ({len(df)} filas)")

        except Exception as e:
            print(f"  [!] Error leyendo {f}: {e}")

    if not res_dfs:
        print("  [!] No se encontraron archivos de resultados en el directorio.")
        print("       Verifique que los CSVs existen y que no tienen '_compl' en el nombre.")
        return pd.DataFrame()

    try:
        master_df = pd.concat(res_dfs, ignore_index=True)
        master_df = master_df.loc[:, ~master_df.columns.duplicated()].copy()

        if 'Documento' in master_df.columns:
            master_df['Documento'] = (master_df['Documento'].astype(str)
                                      .str.strip()
                                      .str.replace(r'\.0$', '', regex=True))

        if 'Area temática' in master_df.columns:
            master_df['Area temática'] = master_df['Area temática'].astype(str).str.strip()

        print(f"  [OK] Master dataframe: {len(master_df)} filas, "
              f"{len(master_df.columns)} columnas.")
        return master_df

    except Exception as e:
        print(f"  [!] Error concatenando dataframes: {e}")
        return pd.DataFrame()


def process_section_reports():
    mes_folder  = os.path.dirname(config.PATH_INTERIM)
    REPORTS_DIR = os.path.join(mes_folder, "Final_Reports", "Reportes_Por_Secciones")
    os.makedirs(REPORTS_DIR, exist_ok=True)

    item_base_dict = load_item_indicators()
    anchor_by_grade = load_anchor_items()

    meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo',
                6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre',
                10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
                
    solid_colors = {"Crítico": "#1A2744", "Bajo": "#1E4D8C", "Medio": "#2979C4",
                    "Bueno": "#5BAAE8", "Excelente": "#B8DEFF"}

    print("Cargando y unificando datos de Analisis Psicometrico...")
    master_df = build_master_geiser_dataframe()
    if master_df.empty: return

    if 'Grupo' in master_df.columns:
        master_df['Grupo_Letra'] = (master_df['Grupo'].astype(str)
                                    .apply(lambda x: x.split('(')[0].strip()))
    else:
        master_df['Grupo_Letra'] = ''

    if 'Nombre' not in master_df.columns:   master_df['Nombre'] = ''
    if 'Apellido' not in master_df.columns: master_df['Apellido'] = ''
    
    # --- APLICAMOS LA LIMPIEZA DE REDACCIÓN A LOS NOMBRES ---
    master_df['Nombre']   = master_df['Nombre'].fillna('').apply(limpiar_textos_redaccion).str.title()
    master_df['Apellido'] = master_df['Apellido'].fillna('').apply(limpiar_textos_redaccion).str.title()

    def format_name(row):
        ap  = row['Apellido']
        nom = row['Nombre']
        if ap and nom: return f"{ap}, {nom}"
        return ap or nom
    master_df['Apellidos y Nombres'] = master_df.apply(format_name, axis=1)

    global_items_map = {}
    for (grado, subject), g_data in master_df.groupby(['Grado', 'Area temática']):
        prefix     = 'MAT' if subject == 'Matemática' else 'LEC'
        item_regex = re.compile(rf'^{prefix}\d+$')
        valid = [c for c in master_df.columns
                 if item_regex.match(c) and g_data[c].notna().any()]
        global_items_map[(grado, subject)] = sorted(
            valid, key=lambda x: int(re.sub(r'\D', '', x)))

    grouped = master_df.groupby(
        ['Nro de centro', 'Centro', 'Grado', 'Grupo', 'Departamento'])
    reportes_generados = 0

    print(f"Generando reportes de secciones en: {REPORTS_DIR} ...")
    for (nro_centro, centro, grado, grupo_original, departamento), group_data in grouped:
        centro_nombre_str = str(centro).strip()
        safe_centro  = re.sub(r'[\\/*?:"<>|]', "", centro_nombre_str)
        school_folder = os.path.join(REPORTS_DIR, f"{nro_centro} - {safe_centro}")
        os.makedirs(school_folder, exist_ok=True)
        safe_grupo_file = (str(grupo_original)
                           .replace(" ", "").replace("(", "_").replace(")", ""))
        tablas_html_combinadas, chart_scripts = "", []

        mes_prueba = "2"
        pruebas_unicas = group_data['Prueba'].dropna().unique()
        for p in pruebas_unicas:
            match = re.search(r'Mes\s*(\d+)', str(p), re.IGNORECASE)
            if match:
                mes_prueba = match.group(1)
                break

        fechas_validas = pd.to_datetime(
            group_data['Fecha-Hora de Inicio'].replace('Sin Especificar', pd.NA),
            errors='coerce', dayfirst=True, format='mixed')
        if not fechas_validas.dropna().empty:
            mes_num      = fechas_validas.dropna().dt.month.mode().iloc[0]
            mes_aplicacion = meses_es.get(int(mes_num), "No especificado")
        else:
            mes_aplicacion = "No especificado"

        for subject in ['Matemática', 'Lectura']:
            subject_data = group_data[group_data['Area temática'] == subject].copy()
            if subject_data.empty: continue

            item_cols_raw = global_items_map.get((grado, subject), [])
            if not item_cols_raw: continue

            grade_num        = re.sub(r'\D', '', str(grado))
            anchor_for_grade = anchor_by_grade.get(grade_num, set())

            item_cols = [col for col in item_cols_raw
                         if normalize_item_code(col) not in anchor_for_grade]
            if not item_cols: continue

            item_display_num = {col: idx + 1 for idx, col in enumerate(item_cols)}

            if 'theta.global (escala 0-100)' not in subject_data.columns:
                subject_data['theta.global (escala 0-100)'] = pd.NA
            if 'anular_prueba' not in subject_data.columns:
                subject_data['anular_prueba'] = pd.NA

            matrix = subject_data[
                ['Documento', 'Apellidos y Nombres',
                 'theta.global (escala 0-100)', 'anular_prueba'] + item_cols
            ].copy()
            matrix['Puntaje_Num'] = pd.to_numeric(
                matrix['theta.global (escala 0-100)']
                .astype(str).str.replace('"', '').str.replace(',', '.'),
                errors='coerce')

            def tiene_tiempo_inusual(val):
                if pd.isna(val) or str(val).strip().lower() in ['nan','none','','null','false']:
                    return False
                val_str = str(val).strip().lower()
                val_str = ''.join(
                    c for c in unicodedata.normalize('NFD', val_str)
                    if unicodedata.category(c) != 'Mn')
                return (("duracion <5 min" in val_str)
                        or ("tiempo extremo de demora" in val_str)
                        or ("5 min" in val_str))

            matrix['tiempo_inusual'] = matrix['anular_prueba'].apply(tiene_tiempo_inusual)
            matrix['Puntaje_Sort']   = matrix['Puntaje_Num'].fillna(9999)
            matrix = matrix.sort_values(by='Puntaje_Sort', ascending=True)

            indicadores_info = []

            if subject == 'Lectura':
                def make_fingerprint(texto):
                    if not texto or texto == 'Indicador no disponible.':
                        return 'missing'
                    code_match = re.search(r'\b(\d+\.\d+)\b', texto)
                    if code_match:
                        return f"obj_{code_match.group(1)}"
                    import unicodedata as _ud
                    norm = _ud.normalize('NFD', texto.lower())
                    norm = ''.join(c for c in norm if _ud.category(c) != 'Mn')
                    norm = re.sub(r'[\W_]+', '', norm)
                    return norm[:60]

                indicator_groups = {}
                for i, item in enumerate(item_cols):
                    norm_item = normalize_item_code(item)
                    base_info = item_base_dict.get(norm_item, {})
                    objetivo  = base_info.get('objetivo', 'Indicador no disponible.')
                    fingerprint = make_fingerprint(objetivo)

                    if fingerprint not in indicator_groups:
                        indicator_groups[fingerprint] = {
                            'display_text': objetivo,
                            'items':        [],
                            'clases':       set(),
                            'f_day':        9999.0,
                            'nivelacion':   set(),
                            'item_indices': []
                        }
                    indicator_groups[fingerprint]['items'].append(item)
                    indicator_groups[fingerprint]['item_indices'].append(item_display_num[item])
                    clase = base_info.get('clases', '')
                    if clase: indicator_groups[fingerprint]['clases'].add(clase)
                    orden = base_info.get('orden', 9999.0)
                    if orden < indicator_groups[fingerprint]['f_day']:
                        indicator_groups[fingerprint]['f_day'] = orden
                    nivelacion = base_info.get('nivelacion', '')
                    if nivelacion: indicator_groups[fingerprint]['nivelacion'].add(nivelacion)

                for fingerprint, data in indicator_groups.items():
                    total_correct = sum(
                        pd.to_numeric(matrix[it], errors='coerce').fillna(0).sum()
                        for it in data['items'])
                    total_possible = len(matrix) * len(data['items'])
                    pct = (total_correct / total_possible * 100) if total_possible > 0 else 0
                    clases_str = " | ".join(sorted(data['clases']))
                    nivel_str  = " | ".join(sorted(data['nivelacion']))
                    nivelacion_html = (f" <strong style='white-space: nowrap;'>"
                                       f"{nivel_str}</strong>") if nivel_str else ""
                    ind_html = (f"<span class='badge-class'>{clases_str}</span> "
                                f"{data['display_text']}{nivelacion_html}"
                                if clases_str
                                else f"{data['display_text']}{nivelacion_html}")
                    indicadores_info.append({
                        'disp_name':      "",
                        'pct':            pct,
                        'indicador_base': ind_html,
                        'f_day':          data['f_day']
                    })

            else:
                for i, item in enumerate(item_cols):
                    pct = ((pd.to_numeric(matrix[item], errors='coerce').fillna(0).sum()
                            / len(matrix)) * 100) if len(matrix) > 0 else 0
                    norm_item       = normalize_item_code(item)
                    base_info       = item_base_dict.get(norm_item, {})
                    objetivo_base   = base_info.get('objetivo', 'Indicador no disponible.')
                    dia_label       = base_info.get('clases', '')
                    first_class_num = base_info.get('orden', 9999.0)
                    nivelacion_texto = base_info.get('nivelacion', '')
                    disp_name = f"MAT {item_display_num[item]}"
                    nivelacion_html = (f" <strong style='white-space: nowrap;'>"
                                       f"{nivelacion_texto}</strong>") if nivelacion_texto else ""
                    ind_html = (f"<span class='badge-class'>{dia_label}</span> "
                                f"{objetivo_base}{nivelacion_html}"
                                if dia_label
                                else f"{objetivo_base}{nivelacion_html}")
                    indicadores_info.append({
                        'disp_name':     disp_name,
                        'pct':           pct,
                        'indicador_base': ind_html,
                        'f_day':         first_class_num
                    })

            indicadores_info.sort(key=lambda x: x['pct'])
            tbody_id = f"tbody_{nro_centro}_{subject.replace(' ', '')}"
            indicators_html = f"""
            <details open class="info-accordion no-print">
                <summary>Descripcion de los Indicadores</summary>
                <p>Muestra los indicadores evaluados, ordenados de mayor dificultad (arriba) a mayor logro (abajo). El recuadro oscuro marca el dia de planificacion de {grado}.</p>
            </details>
            <div class="indicator-section">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; border-bottom:1px solid #eee; padding-bottom:10px;">
                    <h3 id="title-{tbody_id}" style="margin: 0; font-size: 16px; color: #2c3e50;">Porcentaje de alcance del indicador de logro (De menor a mayor alcance)</h3>
                    <button onclick="toggleIndicatorSort(this, 'list-{tbody_id}', 'title-{tbody_id}')" class="btn btn-default no-print" data-sort-state="pct">Ordenar por orden de clases</button>
                </div>
                <div class="indicator-list" id="list-{tbody_id}">
            """
            for ind in indicadores_info:
                if   ind['pct'] <= 20: c, txt = "cat-critico",   "Crítico"
                elif ind['pct'] <= 40: c, txt = "cat-bajo",      "Bajo"
                elif ind['pct'] <= 60: c, txt = "cat-medio",     "Medio"
                elif ind['pct'] <= 80: c, txt = "cat-bueno",     "Bueno"
                else:                  c, txt = "cat-excelente", "Excelente"
                
                disp_html = (f"<span class='item-label'>{ind['disp_name']}:</span> "
                             if ind['disp_name'] else "")
                indicators_html += f"""
                <div class="indicator-item {c}" data-pct="{ind['pct']}" data-class="{ind['f_day']}">
                    <div class="indicator-pct">
                        <div class="pct-number">{ind['pct']:.1f}%</div>
                        <div class="pct-label">{txt}</div>
                    </div>
                    <div>{disp_html}{ind['indicador_base']}</div>
                </div>
                """
            indicators_html += '</div></div>'

            matrix_html = f"""
            <div style="display: flex; justify-content: flex-start; margin-bottom: 10px;" class="no-print">
                <button onclick="toggleSort(this, '{tbody_id}')" class="btn btn-default" data-sort-state="rendimiento" style="font-size: 13px;">
                    Ordenar los estudiantes de la tabla por orden alfabetico
                </button>
            </div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>NIE</th>
                            <th>Apellidos y Nombres</th>
                            <th>Puntaje (0-100)</th>
                        </tr>
                    </thead>
                    <tbody id="{tbody_id}">
            """
            for _, row in matrix.iterrows():
                p               = row['Puntaje_Num']
                tiene_asterisco = bool(row['tiempo_inusual'])
                nombre_visible  = (f"{row['Apellidos y Nombres']} *"
                                   if tiene_asterisco
                                   else str(row['Apellidos y Nombres']))
                bg, lvl = "", ""
                if pd.notna(p):
                    if   p <= 35: bg, lvl = "cat-critico",   "Crítico"
                    elif p <= 45: bg, lvl = "cat-bajo",      "Bajo"
                    elif p <= 55: bg, lvl = "cat-medio",     "Medio"
                    elif p <= 65: bg, lvl = "cat-bueno",     "Bueno"
                    else:         bg, lvl = "cat-excelente", "Excelente"
                    
                matrix_html += (f'<tr data-nombre="{row["Apellidos y Nombres"]}" '
                                 f'data-puntaje="{row["Puntaje_Sort"]}">')
                if lvl:
                    matrix_html += (f"<td>{row['Documento']}</td>"
                                    f"<td class='{bg}'><strong>{nombre_visible}</strong>"
                                    f"<div style='font-size:10px; opacity: 0.8;'>NIVEL: {lvl}</div>"
                                    f"</td><td class='total-cell'>{p:.1f}</td></tr>")
                else:
                    matrix_html += (f"<td>{row['Documento']}</td>"
                                    f"<td><strong>{nombre_visible}</strong></td>"
                                    f"<td class='total-cell'>-</td></tr>")
            matrix_html += """
                    </tbody>
                </table>
            </div>
            <div class='footnote footnote-warning'>
                <strong>* Resultados con observacion:</strong> El estudiante marcado con (*) registro un tiempo de ejecucion menor a 5 minutos o presento un tiempo de demora inusual.
            </div>
            """

            cat_c = {'Crítico': 0, 'Bajo': 0, 'Medio': 0, 'Bueno': 0, 'Excelente': 0}
            for v in matrix['Puntaje_Num'].dropna():
                cat_c[clasificar_puntaje(v)] += 1
            tot      = len(matrix['Puntaje_Num'].dropna())
            d_data   = [(cat_c[k] / tot * 100) if tot else 0
                        for k in ['Crítico', 'Bajo', 'Medio', 'Bueno', 'Excelente']]
            d_data_str = ",".join(f"{x:.1f}" for x in d_data)
            donut_id = f"donut_{subject.replace(' ', '')}_{nro_centro}"
            g_mean   = matrix['Puntaje_Num'].dropna().mean()
            g_cat    = clasificar_puntaje(g_mean) if pd.notna(g_mean) else "Sin Datos"
            b_color  = solid_colors.get(g_cat, "#bdc3c7")

            graficos_footer = f"""
            <div class="indicator-section" style="display:flex; flex-wrap:wrap; justify-content:space-around; align-items:center; page-break-inside:avoid; margin-bottom: 20px;">
                <div style="flex:2; min-width:350px; display:flex; flex-direction:column; align-items:center;">
                    <h3 style="text-align:center; margin-bottom:5px; border:none; padding-bottom:0;">Distribucion de Estudiantes por Nivel de Alcance</h3>
                    <p style="font-size: 12px; color: #7f8c8d; margin-top: 0; text-align: center;">Grafica para los {tot} estudiantes del grupo</p>
                    <div class="chart-wrapper-print" style="position:relative; height:300px; width:100%; display:flex; justify-content:center; margin-top:15px;">
                        <canvas id="{donut_id}"></canvas>
                    </div>
                </div>
                <div style="flex:1; min-width:250px; display:flex; flex-direction:column; align-items:center; border-left:2px dashed #eee; padding-left:20px;">
                    <h3 style="text-align: center; color: #2c3e50; margin-bottom: 20px; border: none;">Promedio de la Seccion</h3>
                    <div class="badge-container" style="--badge-color:{b_color};">
                        <div class="donut-seal">{f'{g_mean:.1f}' if pd.notna(g_mean) else '-'}</div>
                        <div class="donut-label">{g_cat}</div>
                    </div>
                </div>
            </div>
            """
            script = f"""
                (function() {{
                    var ctx = document.getElementById('{donut_id}').getContext('2d');
                    new Chart(ctx, {{
                        type: 'doughnut',
                        data: {{
                            labels: ['Crítico (0-35)', 'Bajo (36-45)', 'Medio (46-55)', 'Bueno (56-65)', 'Excelente (66-100)'],
                            datasets: [{{
                                data: [{d_data_str}],
                                backgroundColor: ['#1A2744', '#1E4D8C', '#2979C4', '#5BAAE8', '#B8DEFF'],
                                borderColor: ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
                                borderWidth: 2,
                                datalabels: {{
                                    color: function(context) {{
                                        return context.dataIndex <= 2 ? '#ffffff' : '#000000';
                                    }},
                                    font: {{ weight: 'bold', size: 14 }},
                                    formatter: function(value) {{ return value > 0 ? value.toFixed(1) + '%' : ''; }}
                                }}
                            }}]
                        }},
                        options: {{ animation: false, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
                    }});
                }})();
            """
            chart_scripts.append(script)
            tablas_html_combinadas += (
                f'<div class="subject-section">'
                f'<div class="subject-title">Asignatura: <strong>{subject}</strong></div>'
                f'{indicators_html}{matrix_html}{graficos_footer}</div>')

        if tablas_html_combinadas:
            safe_grado = str(grado).replace(" ", "").replace("o", "")
            filename   = f"Reporte_{nro_centro}_Grado{safe_grado}_Grupo{safe_grupo_file}.html"
            filepath   = os.path.abspath(os.path.join(school_folder, filename))
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(HTML_TEMPLATE.format(
                    mes_prueba     = mes_prueba,
                    mes_aplicacion = mes_aplicacion,
                    departamento   = departamento if pd.notnull(departamento) else "Sin Especificar",
                    nro_centro     = nro_centro,
                    centro         = centro_nombre_str,
                    grado          = str(grado).replace(" ", "").replace("o", ""),
                    grupo          = grupo_original,
                    tablas_html    = tablas_html_combinadas,
                    chart_scripts  = "\n".join(chart_scripts)
                ))
            reportes_generados += 1

    sep = "=" * 60
    print(f"\n{sep}")
    print(f"  [OK] PROCESO TERMINADO")
    print(sep)
    print(f"  Reportes generados : {reportes_generados}")
    print(f"  Directorio de salida:")
    print(f"    {REPORTS_DIR}")
    print(f"  Los reportes estan organizados por escuela dentro")
    print(f"  de ese directorio (una subcarpeta por centro escolar).")
    print(sep)

if __name__ == "__main__":
    process_section_reports()