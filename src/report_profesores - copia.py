# src/report_profesores.py
import os
import re
import pandas as pd
import numpy as np
import pathlib
import sys
import glob
import unicodedata
import difflib 

# --- 1. CONEXIÓN CON EL MAPA PAARS ---
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# Usamos las rutas dinámicas del sistema
GEISER_CSV_DIR = config.PATH_INTERIM
REPORTS_DIR = os.path.join(config.PATH_REPORTS, "Reportes_Por_Secciones")
MAPEO_DIR = config.PATH_METADATA

os.makedirs(REPORTS_DIR, exist_ok=True)

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
        
        .pdf-link-container {{ text-align: center; margin-bottom: 25px; }}
        .pdf-link-container a {{ color: #2980b9; font-weight: bold; text-decoration: underline; font-size: 14px; }}

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
        
        .total-cell {{ background-color: #ecf0f1; font-weight: bold; font-size: 14px; }}
        
        tbody tr:hover td {{ background-color: #f1f8ff; }}
        tbody tr:hover td:nth-child(2) {{ filter: brightness(0.95); }}

        .indicator-section {{ padding: 20px; background-color: #f8fcfd; border-radius: 6px; border: 1px solid #e1ecef; border-left: 4px solid #1abc9c; margin-bottom: 25px; box-sizing: border-box; }}
        
        .indicator-list {{ display: flex; flex-direction: column; gap: 8px; font-size: 13px; color: #333; }}
        .indicator-item {{ display: flex; align-items: center; padding: 10px 12px; border-radius: 4px; border: 1px solid rgba(0,0,0,0.1); }}
        
        details.info-accordion {{ margin-bottom: 15px; text-align: left; display: block; }}
        details.info-accordion summary {{ cursor: pointer; font-size: 13px; font-weight: bold; color: #2980b9; padding: 8px 12px; background-color: #f4f6f9; border-radius: 4px; border: 1px solid #d0e1f9; display: inline-block; list-style: none; transition: 0.2s; }}
        details.info-accordion summary::-webkit-details-marker {{ display: none; }}
        details.info-accordion summary:hover {{ background-color: #eaf2f8; }}
        details.info-accordion p {{ padding: 15px; border-left: 3px solid #2980b9; background-color: #fafafa; margin-top: 8px; font-size: 13.5px; color: #444; border-radius: 0 4px 4px 0; box-shadow: inset 0 1px 3px rgba(0,0,0,0.03); line-height: 1.5; margin-bottom: 0; }}
        
        .color-red {{ background-color: #fef2f2; border-left: 8px solid #991b1b; }}       
        .color-orange {{ background-color: #fff7ed; border-left: 8px solid #ff8c2E; }}     
        .color-yellow {{ background-color: #fefce8; border-left: 8px solid #facc15; }}     
        .color-lightgreen {{ background-color: #f7fee7; border-left: 8px solid #84cc16; }} 
        .color-darkgreen {{ background-color: #ecfdf5; border-left: 8px solid #065f46; }}  
        
        .bg-red {{ background-color: #fef2f2 !important; border-left: 4px solid #991b1b !important; }}       
        .bg-orange {{ background-color: #fff7ed !important; border-left: 4px solid #ff8c2E !important; }}     
        .bg-yellow {{ background-color: #fefce8 !important; border-left: 4px solid #facc15 !important; }}     
        .bg-lightgreen {{ background-color: #f7fee7 !important; border-left: 4px solid #84cc16 !important; }} 
        .bg-darkgreen {{ background-color: #ecfdf5 !important; border-left: 4px solid #065f46 !important; }}  

        .indicator-pct {{ font-weight: bold; width: 65px; text-align: center; margin-right: 15px; padding-right: 15px; border-right: 1px solid #ccc; flex-shrink: 0; }}
        .pct-number {{ font-size: 16px; color: #2c3e50; }}
        .pct-label {{ font-size: 10px; color: #555; text-transform: uppercase; letter-spacing: 0.5px; }}
        .indicator-item strong {{ color: #2c3e50; display: inline-block; width: 55px; }}
        
        .badge-class {{ display: inline-block; background-color: #34495e; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; margin-right: 10px; vertical-align: baseline; font-weight: bold; letter-spacing: 0.5px; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }}
        
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
    
    <div class="pdf-link-container no-print">
        <a href="{file_uri}" target="_blank">Link para visualizar el html (Guardar para el PDF)</a>
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
def clasificar_puntaje(val):
    if pd.isna(val): return "Crítico"
    val = float(val)
    if val <= 35: return "Crítico"
    elif val <= 45: return "Bajo"
    elif val <= 55: return "Medio"
    elif val <= 65: return "Bueno"
    else: return "Excelente"

def limpiar_texto(texto):
    """Elimina TODO el formato (espacios, tildes, signos) para hacer un match perfecto."""
    if not isinstance(texto, str) or pd.isna(texto): return ""
    texto = str(texto).strip()
    texto = re.sub(r'^\d+([.-]\d+)*[.)\-]?\s*', '', texto)
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-z0-9]', '', texto.lower())
    return texto

def get_grado_key(texto):
    """Traductor unificado robusto para grados."""
    texto = str(texto).lower()
    num = None
    
    if '1er' in texto and ('año' in texto or 'bach' in texto): num = 10
    elif '2do' in texto and ('año' in texto or 'bach' in texto): num = 11
    elif '10' in texto or 'decim' in texto or 'décim' in texto: num = 10
    elif '11' in texto or 'undec' in texto or 'undéc' in texto or 'once' in texto: num = 11
    elif '9' in texto or 'noven' in texto or 'nueve' in texto: num = 9
    elif '8' in texto or 'octav' in texto or 'ocho' in texto: num = 8
    elif '7' in texto or 'septim' in texto or 'séptim' in texto or 'siete' in texto: num = 7
    elif '6' in texto or 'sext' in texto or 'seis' in texto: num = 6
    elif '5' in texto or 'quint' in texto or 'cinco' in texto: num = 5
    elif '4' in texto or 'cuart' in texto or 'cuatro' in texto: num = 4
    elif '3' in texto or 'tercer' in texto or 'tres' in texto: num = 3
    elif '2' in texto or 'segund' in texto or 'dos' in texto: num = 2
    elif '1' in texto or 'primer' in texto or 'uno' in texto: num = 1
    
    if num is None:
        match = re.search(r'(\d+)', texto)
        if match: num = int(match.group(1))
        else: return ""
        
    return f"{num} grado"

def parse_day_for_sorting(d):
    match = re.search(r'(\d+(\.\d+)?)', str(d))
    return float(match.group(1)) if match else 9999.0

def format_multiple_days(days_set):
    if not days_set: 
        return "", 9999.0
    
    sorted_days = sorted(list(days_set), key=parse_day_for_sorting)
    first_day_val = parse_day_for_sorting(sorted_days[0])
    
    formatted_days = []
    for d in sorted_days:
        try:
            num = float(d)
            formatted_days.append(str(int(num)) if num.is_integer() else str(num))
        except:
            formatted_days.append(str(d).strip())
            
    if len(formatted_days) == 1:
        return f"Clase {formatted_days[0]}", first_day_val
    elif len(formatted_days) == 2:
        return f"Clases {formatted_days[0]} y {formatted_days[1]}", first_day_val
    else:
        comas = ", ".join(formatted_days[:-1])
        return f"Clases {comas} y {formatted_days[-1]}", first_day_val

def buscar_dia_borroso(materia, grado, txt_objetivo, dias_ref_dict, threshold=0.70):
    candidatos = [k[2] for k in dias_ref_dict.keys() if k[0] == materia and k[1] == grado]
    if not candidatos: return None
    
    if txt_objetivo in candidatos: 
        return dias_ref_dict[(materia, grado, txt_objetivo)]
        
    matches = difflib.get_close_matches(txt_objetivo, candidatos, n=1, cutoff=threshold)
    if matches: 
        return dias_ref_dict[(materia, grado, matches[0])]
        
    return None

def load_item_indicators():
    ITEMS_FILE = os.path.join(MAPEO_DIR, "Mapeo_Items_Grados_Completo_Pregreso_Mes_1_2026.xlsx")
    item_base_dict = {}
    dias_referencia = {}
    if not os.path.exists(ITEMS_FILE): return item_base_dict, dias_referencia

    print("  -> Leyendo Currículos y Mapeando clases...")
    
    # MATEMÁTICA
    math_file = os.path.join(MAPEO_DIR, "Curriculum_Math_Didactic_Sequence.xlsx")
    if os.path.exists(math_file):
        try:
            xls_math = pd.read_excel(math_file, sheet_name=None, dtype=str)
            for sheet_name, df_cur in xls_math.items():
                grado_key = get_grado_key(sheet_name)
                df_cur.columns = df_cur.columns.str.strip()
                
                # Buscador inteligente de columnas
                ind_col = next((c for c in df_cur.columns if 'indicador' in c.lower()), None)
                dia_col = next((c for c in df_cur.columns if 'dia' in c.lower() or 'día' in c.lower()), None)
                
                if ind_col and dia_col:
                    for _, row in df_cur.iterrows():
                        ind_bundle = str(row[ind_col])
                        dia_val = str(row[dia_col]).strip()
                        if ind_bundle.lower() != 'nan' and dia_val.lower() != 'nan':
                            parts = [p.strip() for p in re.sub(r'[\n;]', '|', ind_bundle).split('|') if p.strip()]
                            for p in parts:
                                txt_limpio = limpiar_texto(p)
                                if txt_limpio:
                                    key = ("MAT", grado_key, txt_limpio)
                                    if key not in dias_referencia: dias_referencia[key] = set()
                                    dias_referencia[key].add(dia_val)
        except Exception: pass

    # LENGUAJE
    lang_file = os.path.join(MAPEO_DIR, "Curriculum_Language_Didactic.xlsx")
    if os.path.exists(lang_file):
        try:
            xls_lang = pd.read_excel(lang_file, sheet_name=None, dtype=str)
            for sheet_name, df_cur in xls_lang.items():
                grado_key = get_grado_key(sheet_name)
                df_cur.columns = df_cur.columns.str.strip()
                
                ind_col = next((c for c in df_cur.columns if 'indicador' in c.lower()), None)
                dia_col = next((c for c in df_cur.columns if 'dia' in c.lower() or 'día' in c.lower()), None)
                
                if ind_col and dia_col:
                    for _, row in df_cur.iterrows():
                        ind_bundle = str(row[ind_col])
                        dia_val = str(row[dia_col]).strip()
                        if ind_bundle.lower() != 'nan' and dia_val.lower() != 'nan':
                            parts = [p.strip() for p in re.sub(r'[\n;]', '|', ind_bundle).split('|') if p.strip()]
                            for p in parts:
                                txt_limpio = limpiar_texto(p)
                                if txt_limpio:
                                    key = ("LEC", grado_key, txt_limpio)
                                    if key not in dias_referencia: dias_referencia[key] = set()
                                    dias_referencia[key].add(dia_val)
        except Exception: pass

    # Mapeo ítems
    try:
        df_items = pd.read_excel(ITEMS_FILE, dtype=str)
        for _, row in df_items.dropna(subset=['ItemCodigo', 'Objetivo', 'Grado']).iterrows():
            item_base_dict[row['ItemCodigo'].strip()] = {'objetivo': row['Objetivo'].strip(), 'grado_orig': row['Grado'].strip()}
    except Exception: pass
    return item_base_dict, dias_referencia

def build_master_geiser_dataframe():
    compl_dfs, res_dfs = [], []
    if not os.path.exists(GEISER_CSV_DIR): return pd.DataFrame()
    for f in os.listdir(GEISER_CSV_DIR):
        if not f.lower().endswith('.csv') or 'legend' in f.lower(): continue
        path = os.path.join(GEISER_CSV_DIR, f)
        try:
            df = pd.read_csv(path, dtype=str, encoding='utf-8-sig')
            df.columns = df.columns.str.strip()
            if 'MAT-' in f.upper(): df['Area temática'] = 'Matemática'
            elif 'LEC-' in f.upper(): df['Area temática'] = 'Lengua'
            elif 'Area temática' not in df.columns: df['Area temática'] = 'Desconocida'
            if '_compl' in f.lower(): compl_dfs.append(df)
            elif 'resultados' in f.lower(): res_dfs.append(df)
        except: pass
    if not compl_dfs: return pd.DataFrame()
    compl_df = pd.concat(compl_dfs, ignore_index=True).drop_duplicates(subset=['Documento', 'Area temática'])
    if res_dfs:
        res_df = pd.concat(res_dfs, ignore_index=True).drop_duplicates(subset=['Documento', 'Area temática'])
        master_df = pd.merge(compl_df, res_df[['Documento', 'Area temática', 'theta.global (escala 0-100)', 'anular_prueba']], on=['Documento', 'Area temática'], how='left')
    else:
        master_df = compl_df
        master_df['theta.global (escala 0-100)'] = pd.NA
        master_df['anular_prueba'] = pd.NA
    return master_df

def process_section_reports():
    item_base_dict, dias_referencia = load_item_indicators()
    meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
    solid_colors = {"Crítico": "#991b1b", "Bajo": "#ff8c2e", "Medio": "#facc15", "Bueno": "#84cc16", "Excelente": "#065f46"}

    master_df = build_master_geiser_dataframe()
    if master_df.empty: return
    
    if 'Grupo' in master_df.columns:
        master_df['Grupo_Letra'] = master_df['Grupo'].astype(str).apply(lambda x: x.split('(')[0].strip())
    else:
        master_df['Grupo_Letra'] = ''

    if 'Nombre' not in master_df.columns:
        master_df['Nombre'] = ''
    master_df['Nombre'] = master_df['Nombre'].fillna('').astype(str).str.strip().str.title()
    
    if 'Apellido' not in master_df.columns:
        master_df['Apellido'] = ''
    master_df['Apellido'] = master_df['Apellido'].fillna('').astype(str).str.strip().str.title()
    
    def format_name(row):
        ap = row['Apellido']
        nom = row['Nombre']
        if ap and nom: return f"{ap}, {nom}"
        return ap or nom
        
    master_df['Apellidos y Nombres'] = master_df.apply(format_name, axis=1)
    
    global_items_map = {}
    for (grado, subject), g_data in master_df.groupby(['Grado', 'Area temática']):
        prefix = 'MAT' if subject == 'Matemática' else 'LEC'
        item_regex = re.compile(rf'^{prefix}\d+$')
        valid = [c for c in master_df.columns if item_regex.match(c) and g_data[c].notna().any()]
        global_items_map[(grado, subject)] = sorted(valid, key=lambda x: int(re.sub(r'\D', '', x)))

    grouped = master_df.groupby(['Nro de centro', 'Centro', 'Grado', 'Grupo', 'Departamento'])
    reportes_generados = 0

    for (nro_centro, centro, grado, grupo_original, departamento), group_data in grouped:
        centro_nombre_str = str(centro).strip()
        safe_centro = re.sub(r'[\\/*?:"<>|]', "", centro_nombre_str)
        school_folder = os.path.join(REPORTS_DIR, f"{nro_centro} - {safe_centro}")
        os.makedirs(school_folder, exist_ok=True)
        safe_grupo_file = str(grupo_original).replace(" ", "").replace("(", "_").replace(")", "")
        tablas_html_combinadas, chart_scripts = "", []
        grado_eval_key = get_grado_key(grado)

        mes_prueba = "2" 
        pruebas_unicas = group_data['Prueba'].dropna().unique()
        for p in pruebas_unicas:
            match = re.search(r'Mes\s*(\d+)', str(p), re.IGNORECASE)
            if match:
                mes_prueba = match.group(1)
                break
                
        fechas_validas = pd.to_datetime(group_data['Fecha-Hora de Inicio'].replace('Sin Especificar', pd.NA), errors='coerce', dayfirst=True, format='mixed')
        if not fechas_validas.dropna().empty:
            mes_num = fechas_validas.dropna().dt.month.mode().iloc[0]
            mes_aplicacion = meses_es.get(int(mes_num), "No especificado")
        else:
            mes_aplicacion = "No especificado"

        for subject in ['Matemática', 'Lengua']:
            subject_data = group_data[group_data['Area temática'] == subject].copy()
            if subject_data.empty: continue
            item_cols = global_items_map.get((grado, subject), [])
            if not item_cols: continue
            
            if 'theta.global (escala 0-100)' not in subject_data.columns: 
                subject_data['theta.global (escala 0-100)'] = pd.NA
            if 'anular_prueba' not in subject_data.columns: 
                subject_data['anular_prueba'] = pd.NA
            
            matrix = subject_data[['Documento', 'Apellidos y Nombres', 'theta.global (escala 0-100)', 'anular_prueba'] + item_cols].copy()
            matrix['Puntaje_Num'] = pd.to_numeric(matrix['theta.global (escala 0-100)'].astype(str).str.replace('"', '').str.replace(',', '.'), errors='coerce')
            
            # --- SOLUCIÓN EXACTA DEL ASTERISCO ---
            matrix['tiempo_inusual'] = matrix['anular_prueba'].fillna('').astype(str).str.contains("duración <5 min o tiempo extremo de demora", case=False, na=False)
            
            matrix['Puntaje_Sort'] = matrix['Puntaje_Num'].fillna(9999) 
            matrix = matrix.sort_values(by='Puntaje_Sort', ascending=True)

            indicadores_info = []
            materia_prefix = "MAT" if subject == 'Matemática' else "LEC"
            for i, item in enumerate(item_cols):
                pct = (pd.to_numeric(matrix[item], errors='coerce').fillna(0).sum() / len(matrix)) * 100 if len(matrix) > 0 else 0
                obj = item_base_dict.get(item, {}).get('objetivo', 'Indicador no disponible.')
                g_orig = item_base_dict.get(item, {}).get('grado_orig', '')
                
                txt_obj_limpio = limpiar_texto(obj)
                dias_set = buscar_dia_borroso(materia_prefix, grado_eval_key, txt_obj_limpio, dias_referencia) or set()
                
                dia_label, f_day = format_multiple_days(dias_set)
                disp_name = f"{'MAT' if subject == 'Matemática' else 'LEN'}{i+1}"
                ind_html = f"<span class='badge-class'>{dia_label}</span> {obj}" + (f" ({g_orig})" if g_orig else "") if dia_label else f"{obj}" + (f" ({g_orig})" if g_orig else "")
                indicadores_info.append({'disp_name': disp_name, 'pct': pct, 'indicador_base': ind_html, 'f_day': f_day})

            indicadores_info.sort(key=lambda x: x['pct'])
            tbody_id = f"tbody_{nro_centro}_{subject.replace(' ', '')}"
            indicators_html = f"""
            <details open class="info-accordion no-print">
                <summary>📖 Descripción de los Indicadores</summary>
                <p>Muestra los indicadores evaluados, ordenados de mayor dificultad (arriba) a mayor logro (abajo). El recuadro oscuro marca el día de planificación de {grado}.</p>
            </details>
            <div class="indicator-section">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; border-bottom:1px solid #eee; padding-bottom:10px;">
                    <h3 id="title-{tbody_id}" style="margin: 0; font-size: 16px; color: #2c3e50;">Porcentaje de alcance del indicador de logro (De menor a mayor alcance)</h3>
                    <button onclick="toggleIndicatorSort(this, 'list-{tbody_id}', 'title-{tbody_id}')" class="btn btn-default no-print" data-sort-state="pct">Ordenar por orden de clases</button>
                </div>
                <div class="indicator-list" id="list-{tbody_id}">
            """
            for ind in indicadores_info:
                if ind['pct'] <= 20: c, txt = "color-red", "Crítico"
                elif ind['pct'] <= 40: c, txt = "color-orange", "Bajo"
                elif ind['pct'] <= 60: c, txt = "color-yellow", "Medio"
                elif ind['pct'] <= 80: c, txt = "color-lightgreen", "Bueno"
                else: c, txt = "color-darkgreen", "Excelente"
                indicators_html += f"""
                <div class="indicator-item {c}" data-pct="{ind['pct']}" data-class="{ind['f_day']}">
                    <div class="indicator-pct">
                        <div class="pct-number">{ind['pct']:.1f}%</div>
                        <div class="pct-label">{txt}</div>
                    </div>
                    <div><strong>{ind['disp_name']}:</strong> {ind['indicador_base']}</div>
                </div>
                """
            indicators_html += '</div></div>'

            matrix_html = f"""
            <div style="display: flex; justify-content: flex-start; margin-bottom: 10px;" class="no-print">
                <button onclick="toggleSort(this, '{tbody_id}')" class="btn btn-default" data-sort-state="rendimiento" style="font-size: 13px;">
                    Ordenar los estudiantes de la tabla por orden alfabético
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
                p = row['Puntaje_Num']
                
                tiene_asterisco = bool(row['tiempo_inusual'])
                nombre_visible = f"{row['Apellidos y Nombres']} *" if tiene_asterisco else str(row['Apellidos y Nombres'])
                
                bg, lvl = "", ""
                
                if pd.notna(p):
                    if p <= 35: bg, lvl = "bg-red", "Crítico"
                    elif p <= 45: bg, lvl = "bg-orange", "Bajo"
                    elif p <= 55: bg, lvl = "bg-yellow", "Medio"
                    elif p <= 65: bg, lvl = "bg-lightgreen", "Bueno"
                    else: bg, lvl = "bg-darkgreen", "Excelente"
                
                matrix_html += f'<tr data-nombre="{row["Apellidos y Nombres"]}" data-puntaje="{row["Puntaje_Sort"]}">'
                
                if lvl:
                    matrix_html += f"<td>{row['Documento']}</td><td class='{bg}'><strong>{nombre_visible}</strong><div style='font-size:10px; color:#555;'>NIVEL: {lvl}</div></td><td class='total-cell'>{p:.1f}</td></tr>"
                else:
                    matrix_html += f"<td>{row['Documento']}</td><td><strong>{nombre_visible}</strong></td><td class='total-cell'>-</td></tr>"
                    
            matrix_html += """
                    </tbody>
                </table>
            </div>
            <div class='footnote footnote-warning'>
                <strong>* Resultados con observación:</strong> El estudiante marcado con (*) registró un tiempo de ejecución menor a 5 minutos o presentó un tiempo de demora inusual.
            </div>
            """

            cat_c = {'Crítico':0, 'Bajo':0, 'Medio':0, 'Bueno':0, 'Excelente':0}
            for v in matrix['Puntaje_Num'].dropna():
                cat_c[clasificar_puntaje(v)] += 1
            
            tot = len(matrix['Puntaje_Num'].dropna())
            d_data = [ (cat_c[k]/tot*100) if tot else 0 for k in ['Crítico', 'Bajo', 'Medio', 'Bueno', 'Excelente'] ]
            d_data_str = ",".join(f"{x:.1f}" for x in d_data)
            
            donut_id = f"donut_{subject.replace(' ','')}_{nro_centro}"
            g_mean = matrix['Puntaje_Num'].dropna().mean()
            g_cat = clasificar_puntaje(g_mean) if pd.notna(g_mean) else "Sin Datos"
            b_color = solid_colors.get(g_cat, "#bdc3c7")
            
            graficos_footer = f"""
            <div class="indicator-section" style="display:flex; flex-wrap:wrap; justify-content:space-around; align-items:center; page-break-inside:avoid; margin-bottom: 20px;">
                <div style="flex:2; min-width:350px; display:flex; flex-direction:column; align-items:center;">
                    <h3 style="text-align:center; margin-bottom:5px; border:none; padding-bottom:0;">Distribución de Estudiantes por Nivel de Alcance</h3>
                    <p style="font-size: 12px; color: #7f8c8d; margin-top: 0; text-align: center;">Gráfica para los {tot} estudiantes del grupo</p>
                    <div class="chart-wrapper-print" style="position:relative; height:300px; width:100%; display:flex; justify-content:center; margin-top:15px;">
                        <canvas id="{donut_id}"></canvas>
                    </div>
                </div>
                
                <div style="flex:1; min-width:250px; display:flex; flex-direction:column; align-items:center; border-left:2px dashed #eee; padding-left:20px;">
                    <h3 style="text-align: center; color: #2c3e50; margin-bottom: 20px; border: none;">Promedio de la Sección</h3>
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
                                backgroundColor: ['rgba(153, 27, 27, 0.8)', 'rgba(255, 140, 46, 0.8)', 'rgba(250, 204, 21, 0.8)', 'rgba(132, 204, 22, 0.8)', 'rgba(6, 95, 70, 0.8)'],
                                borderColor: ['#991b1b', '#ff8c2e', '#facc15', '#84cc16', '#065f46'],
                                borderWidth: 1,
                                datalabels: {{
                                    color: '#ffffff',
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
            tablas_html_combinadas += f'<div class="subject-section"><div class="subject-title">Asignatura: <strong>{subject}</strong></div>{indicators_html}{matrix_html}{graficos_footer}</div>'

        if tablas_html_combinadas:
            safe_grado = str(grado).replace(" ", "").replace("º", "")
            filename = f"Reporte_{nro_centro}_Grado{safe_grado}_Grupo{safe_grupo_file}.html"
            filepath = os.path.abspath(os.path.join(school_folder, filename))
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(HTML_TEMPLATE.format(
                    mes_prueba=mes_prueba, 
                    mes_aplicacion=mes_aplicacion, 
                    departamento=departamento if pd.notnull(departamento) else "Sin Especificar", 
                    nro_centro=nro_centro, 
                    centro=centro_nombre_str, 
                    grado=str(grado).replace(" ", "").replace("º", ""), 
                    grupo=grupo_original, 
                    file_uri=pathlib.Path(filepath).as_uri(), 
                    tablas_html=tablas_html_combinadas, 
                    chart_scripts="\n".join(chart_scripts)
                ))
            reportes_generados += 1
            
    print(f"\n¡Proceso Terminado! Se generaron {reportes_generados} reportes listos para revisar e imprimir.")

if __name__ == "__main__":
    process_section_reports()