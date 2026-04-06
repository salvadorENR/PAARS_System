# src/report_profesores.py
import os
import re
import pandas as pd
import numpy as np
import pathlib
import sys

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
        
        .table-container {{ max-width: 100%; max-height: 55vh; overflow: auto; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 15px; }}
        table {{ border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; font-size: 12px; }}
        th, td {{ padding: 8px 12px; text-align: center; border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; white-space: nowrap; }}
        
        th {{ background-color: #34495e; color: white; font-weight: 600; position: sticky; top: 0; z-index: 10; border-bottom: 2px solid #2c3e50; }}
        
        .item-header {{ cursor: help; text-decoration: underline dotted; }}
        .item-header:hover {{ background-color: #2980b9; }}
        
        th:nth-child(2), td:nth-child(2) {{ position: sticky; left: 0; background-color: #fdfdfd; z-index: 5; text-align: left; border-right: 2px solid #bdc3c7; font-weight: bold; box-shadow: 2px 0 5px rgba(0,0,0,0.05); vertical-align: middle; }}
        th:nth-child(2) {{ background-color: #2c3e50; color: white; z-index: 15; }}
        
        .correct {{ color: #27ae60; font-weight: bold; font-size: 14px; }}
        .incorrect {{ color: #e74c3c; font-weight: bold; font-size: 14px; }}
        .total-cell {{ background-color: #ecf0f1; font-weight: bold; font-size: 13px; }}
        
        tbody tr:hover td {{ background-color: #f1f8ff; }}
        tbody tr:hover td:nth-child(2) {{ filter: brightness(0.95); }}

        .indicator-section {{ padding: 20px; background-color: #f8fcfd; border-radius: 6px; border: 1px solid #e1ecef; border-left: 4px solid #1abc9c; margin-bottom: 25px; box-sizing: border-box; }}
        .indicator-section h3 {{ margin-top: 0; font-size: 16px; color: #2c3e50; margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 8px; }}
        
        .indicator-list {{ display: flex; flex-direction: column; gap: 8px; font-size: 13px; color: #333; }}
        .indicator-item {{ display: flex; align-items: center; padding: 10px 12px; border-radius: 4px; border: 1px solid rgba(0,0,0,0.1); }}
        
        .color-red {{ background-color: #fdedec; border-left: 8px solid #c0392b; }}       
        .color-orange {{ background-color: #fbeee6; border-left: 8px solid #d35400; }}     
        .color-yellow {{ background-color: #fef9e7; border-left: 8px solid #f1c40f; }}     
        .color-lightgreen {{ background-color: #e8f8f5; border-left: 8px solid #48c9b0; }} 
        .color-darkgreen {{ background-color: #d4efdf; border-left: 8px solid #196f3d; }}  
        
        .bg-red {{ background-color: #fdedec !important; border-left: 4px solid #c0392b !important; }}       
        .bg-orange {{ background-color: #fbeee6 !important; border-left: 4px solid #d35400 !important; }}     
        .bg-yellow {{ background-color: #fef9e7 !important; border-left: 4px solid #f1c40f !important; }}     
        .bg-lightgreen {{ background-color: #e8f8f5 !important; border-left: 4px solid #48c9b0 !important; }} 
        .bg-darkgreen {{ background-color: #d4efdf !important; border-left: 4px solid #196f3d !important; }}  

        .indicator-pct {{ font-weight: bold; width: 65px; text-align: center; margin-right: 15px; padding-right: 15px; border-right: 1px solid #ccc; }}
        .pct-number {{ font-size: 16px; color: #2c3e50; }}
        .pct-label {{ font-size: 10px; color: #555; text-transform: uppercase; letter-spacing: 0.5px; }}
        .indicator-item strong {{ color: #2c3e50; display: inline-block; width: 55px; }}
        
        .footnote {{ font-size: 13px; color: #333; margin-bottom: 10px; padding: 10px 15px; border-radius: 4px; }}
        .footnote-warning {{ background-color: #fdf2f2; border-left: 4px solid #e74c3c; color: #c0392b; }}
        .footnote-info {{ background-color: #fef9e7; border-left: 4px solid #f39c12; color: #d35400; margin-bottom: 30px; }}
        
        /* Estilos del Sello de Calidad */
        .badge-container {{ display: flex; flex-direction: column; align-items: center; justify-content: center; }}
        .donut-seal {{ width: 95px; height: 95px; border-radius: 50%; border: 14px solid var(--badge-color); display: flex; align-items: center; justify-content: center; font-size: 28px; font-weight: 900; color: #2c3e50; background: #fff; box-shadow: inset 0 2px 4px rgba(0,0,0,0.1); }}
        .donut-label {{ margin-top: 15px; font-size: 16px; font-weight: bold; color: var(--badge-color); text-transform: uppercase; text-align: center; background: rgba(0,0,0,0.03); padding: 6px 15px; border-radius: 15px; border: 1px solid var(--badge-color); }}
        
        @media print {{
            .no-print {{ display: none !important; }}
            body {{ background-color: white; padding: 0; }}
            .header-card, .subject-section {{ box-shadow: none; border: 1px solid #ddd; break-inside: avoid; }}
            .table-container {{ max-height: none !important; overflow: visible !important; page-break-inside: auto; }}
            tr {{ page-break-inside: avoid; }}
            canvas {{ max-height: 350px !important; width: 100% !important; }}
        }}
    </style>
</head>
<body>

    <div class="controls no-print">
        <button onclick="toggleItems()" class="btn btn-default" id="toggleBtn">Ocultar columnas de los ítems</button>
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
        // Función para ocultar y mostrar ítems
        function toggleItems() {{
            var cols = document.querySelectorAll('.item-col');
            var btn = document.getElementById('toggleBtn');
            var isHidden = cols.length > 0 && cols[0].style.display === 'none';
            
            cols.forEach(function(col) {{ 
                col.style.display = isHidden ? '' : 'none'; 
            }});
            
            btn.innerHTML = isHidden ? 'Ocultar columnas de los ítems' : 'Mostrar columnas de los ítems';
        }}

        // Función dinámica de ordenamiento (Alfabético vs Desempeño)
        function toggleSort(btn, tbodyId) {{
            var tbody = document.getElementById(tbodyId);
            var rows = Array.from(tbody.querySelectorAll('tr'));
            
            var currentState = btn.getAttribute('data-sort-state') || 'rendimiento';
            
            if (currentState === 'rendimiento') {{
                // Cambiar a Orden Alfabético
                rows.sort(function(a, b) {{
                    var nameA = a.getAttribute('data-nombre') || '';
                    var nameB = b.getAttribute('data-nombre') || '';
                    return nameA.localeCompare(nameB, 'es', {{sensitivity: 'base'}});
                }});
                btn.setAttribute('data-sort-state', 'alfabetico');
                btn.innerHTML = 'Ordenar los estudiantes de la tabla por desempeño en la prueba';
            }} else {{
                // Cambiar a Orden por Desempeño
                rows.sort(function(a, b) {{
                    var scoreA = parseFloat(a.getAttribute('data-puntaje'));
                    var scoreB = parseFloat(b.getAttribute('data-puntaje'));
                    return scoreA - scoreB;
                }});
                btn.setAttribute('data-sort-state', 'rendimiento');
                btn.innerHTML = 'Ordenar los estudiantes de la tabla por orden alfabético';
            }}
            
            // Reinsertar las filas ordenadas
            rows.forEach(function(row) {{ tbody.appendChild(row); }});
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
    if pd.isna(val): return "Crítica"
    val = float(val)
    if val <= 35: return "Crítica"
    elif val <= 45: return "Baja"
    elif val <= 55: return "Media"
    elif val <= 65: return "Buena"
    else: return "Excelente"

def load_item_indicators():
    ITEMS_FILE = os.path.join(MAPEO_DIR, "Mapeo_Items_Grados_Completo_Pregreso_Mes_1_2026.xlsx")
    if not os.path.exists(ITEMS_FILE): return {}
    try:
        df_items = pd.read_excel(ITEMS_FILE, dtype=str)
        df_items = df_items.dropna(subset=['ItemCodigo', 'Objetivo', 'Grado'])
        item_dict = {}
        for _, row in df_items.iterrows():
            item_dict[row['ItemCodigo'].strip()] = f"{row['Objetivo'].strip()} ({row['Grado'].strip()})"
        return item_dict
    except:
        return {}

# --- 4. CARGA EXCLUSIVA DE ARCHIVOS GEISER ---
def build_master_geiser_dataframe():
    compl_dfs = []
    res_dfs = []
    
    if not os.path.exists(GEISER_CSV_DIR):
        print(f"ADVERTENCIA: No se encontró la carpeta {GEISER_CSV_DIR}")
        return pd.DataFrame()
        
    for f in os.listdir(GEISER_CSV_DIR):
        f_lower = f.lower()
        if not f_lower.endswith('.csv') or 'legend' in f_lower: continue
        
        path = os.path.join(GEISER_CSV_DIR, f)
        try:
            df = pd.read_csv(path, dtype=str, encoding='utf-8-sig')
            df.columns = df.columns.str.strip()
            
            if 'MAT-' in f.upper(): df['Area temática'] = 'Matemática'
            elif 'LEC-' in f.upper(): df['Area temática'] = 'Lengua'
            elif 'Area temática' not in df.columns: df['Area temática'] = 'Desconocida'
            
            if '_compl' in f_lower: compl_dfs.append(df)
            elif 'resultados' in f_lower: res_dfs.append(df)
        except Exception as e:
            print(f"Error al leer {f}: {e}")
            
    if not compl_dfs: return pd.DataFrame()
        
    compl_df = pd.concat(compl_dfs, ignore_index=True)
    compl_df['Documento'] = compl_df['Documento'].astype(str).str.strip()
    compl_df['Area temática'] = compl_df['Area temática'].replace('Lectura', 'Lengua')
    compl_df = compl_df.drop_duplicates(subset=['Documento', 'Area temática'])
    
    demographic_cols = ['Departamento', 'Nro de centro', 'Centro', 'Grado', 'Grupo', 'Prueba', 'Fecha-Hora de Inicio']
    for col in demographic_cols:
        if col not in compl_df.columns: compl_df[col] = 'Sin Especificar'
        else: compl_df[col] = compl_df[col].fillna('Sin Especificar')
        
    if res_dfs:
        res_df = pd.concat(res_dfs, ignore_index=True)
        res_df['Documento'] = res_df['Documento'].astype(str).str.strip()
        res_df['Area temática'] = res_df['Area temática'].replace('Lectura', 'Lengua')
        res_df = res_df.drop_duplicates(subset=['Documento', 'Area temática'])
        
        cols_to_merge = ['Documento', 'Area temática']
        if 'theta.global (escala 0-100)' in res_df.columns: cols_to_merge.append('theta.global (escala 0-100)')
        if 'anular_prueba' in res_df.columns: cols_to_merge.append('anular_prueba')
        
        if 'anular_prueba' in compl_df.columns: compl_df = compl_df.drop(columns=['anular_prueba'])
        if 'theta.global (escala 0-100)' in compl_df.columns: compl_df = compl_df.drop(columns=['theta.global (escala 0-100)'])
            
        res_subset = res_df[cols_to_merge]
        master_df = pd.merge(compl_df, res_subset, on=['Documento', 'Area temática'], how='left')
    else:
        master_df = compl_df
        master_df['theta.global (escala 0-100)'] = pd.NA
        master_df['anular_prueba'] = pd.NA
        
    return master_df

# --- 5. PROCESAMIENTO PRINCIPAL ---
def process_section_reports():
    item_indicators = load_item_indicators()
    meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
    
    solid_colors = {
        "Crítica": "#c0392b", "Baja": "#d35400", "Media": "#f1c40f", 
        "Buena": "#1abc9c", "Excelente": "#196f3d"
    }

    print("Cargando y unificando datos de Analisis Psicometrico...")
    master_df = build_master_geiser_dataframe()
    if master_df.empty: return
    
    master_df['Grupo_Letra'] = master_df['Grupo'].astype(str).apply(lambda x: x.split('(')[0].strip())
    
    master_df['Nombre'] = master_df.get('Nombre', pd.Series(['']*len(master_df))).fillna('').astype(str).str.strip().str.title()
    master_df['Apellido'] = master_df.get('Apellido', pd.Series(['']*len(master_df))).fillna('').astype(str).str.strip().str.title()
    
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
        potential_cols = [col for col in master_df.columns if item_regex.match(col)]
        valid = [col for col in potential_cols if g_data[col].notna().any()]
        valid_sorted = sorted(valid, key=lambda x: int(re.sub(r'\D', '', x)))
        global_items_map[(grado, subject)] = valid_sorted

    grouped = master_df.groupby(['Nro de centro', 'Centro', 'Grado', 'Grupo', 'Grupo_Letra', 'Departamento'])
    reportes_generados = 0
    print(f"Generando reportes de secciones en: {REPORTS_DIR} ...")

    for (nro_centro, centro, grado, grupo_original, grupo_letra, departamento), group_data in grouped:
        centro_nombre_str = str(centro).strip()
        safe_centro_name = re.sub(r'[\\/*?:"<>|]', "", centro_nombre_str)
        school_folder_path = os.path.join(REPORTS_DIR, f"{nro_centro} - {safe_centro_name}")
        os.makedirs(school_folder_path, exist_ok=True)
        
        safe_grupo_file = str(grupo_original).replace(" ", "").replace("(", "_").replace(")", "")
        tablas_html_combinadas = ""
        chart_scripts = []
        subjects_order = ['Matemática', 'Lengua']
        
        mes_prueba = "2" 
        pruebas_unicas = group_data['Prueba'].dropna().unique()
        for p in pruebas_unicas:
            match = re.search(r'Mes\s*(\d+)', str(p), re.IGNORECASE)
            if match:
                mes_prueba = match.group(1)
                break
                
        # AQUÍ ESTÁ EL CAMBIO PARA QUITAR LA ADVERTENCIA: format='mixed'
        fechas_validas = pd.to_datetime(group_data['Fecha-Hora de Inicio'].replace('Sin Especificar', pd.NA), errors='coerce', dayfirst=True, format='mixed')
        
        if not fechas_validas.dropna().empty:
            mes_num = fechas_validas.dropna().dt.month.mode().iloc[0]
            mes_aplicacion = meses_es.get(int(mes_num), "No especificado")
        else:
            mes_aplicacion = "No especificado"

        for subject in subjects_order:
            subject_data = group_data[group_data['Area temática'] == subject].copy()
            if subject_data.empty: continue 
            
            item_cols = global_items_map.get((grado, subject), [])
            if not item_cols: continue
            
            for col in item_cols:
                if col not in subject_data.columns: subject_data[col] = 0
                subject_data[col] = pd.to_numeric(subject_data[col], errors='coerce').fillna(0)
            
            if 'theta.global (escala 0-100)' not in subject_data.columns: subject_data['theta.global (escala 0-100)'] = pd.NA
            if 'anular_prueba' not in subject_data.columns: subject_data['anular_prueba'] = pd.NA
                
            matrix_cols = ['Documento', 'Apellidos y Nombres', 'theta.global (escala 0-100)', 'anular_prueba'] + item_cols
            matrix = subject_data[matrix_cols].copy()
            
            total_items = len(item_cols)
            matrix['Suma_Correctos'] = matrix[item_cols].sum(axis=1)
            matrix['Porcentaje Numérico'] = (matrix['Suma_Correctos'] / total_items) * 100 if total_items > 0 else 0
            matrix['Porcentaje de ítems correctos'] = matrix['Porcentaje Numérico'].apply(lambda x: f"{x:.1f}%")
            matrix.drop(columns=['Suma_Correctos'], inplace=True)
            
            matrix['Puntaje_Num'] = pd.to_numeric(matrix['theta.global (escala 0-100)'], errors='coerce')
            matrix['anular_str'] = matrix['anular_prueba'].fillna('').astype(str)
            
            def check_invalid_time(row):
                return "duración <5 min o tiempo extremo de demora" in row['anular_str']
                
            matrix['tiempo_inusual'] = matrix.apply(check_invalid_time, axis=1)
            
            # ORDENAMIENTO POR DEFECTO: DESEMPEÑO EN LA PRUEBA
            matrix['Puntaje_Sort'] = matrix['Puntaje_Num']
            matrix.loc[matrix['tiempo_inusual'], 'Puntaje_Sort'] = 9999
            matrix = matrix.sort_values(by='Puntaje_Sort', ascending=True)
            
            total_grupo = len(matrix)
            
            valid_matrix = matrix[~matrix['tiempo_inusual']]
            group_mean_val = valid_matrix['Puntaje_Num'].mean()
            if pd.notna(group_mean_val):
                group_cat = clasificar_puntaje(group_mean_val)
                group_mean_str = f"{group_mean_val:.1f}"
                badge_color = solid_colors.get(group_cat, "#bdc3c7")
            else:
                group_cat = "Sin Datos"
                group_mean_str = "-"
                badge_color = "#bdc3c7"
            
            display_prefix = 'MAT' if subject == 'Matemática' else 'LEN'
            display_names = {col: f"{display_prefix}{i+1}" for i, col in enumerate(item_cols)}
            
            indicadores_info = []
            for item in item_cols:
                if total_grupo > 0:
                    pct = (matrix[item].sum() / total_grupo) * 100
                else: pct = 0
                indicadores_info.append({
                    'disp_name': display_names[item], 'pct': pct,
                    'indicador_base': item_indicators.get(item, "Indicador no disponible.")
                })

            subject_html = f'<div class="subject-section"><div class="subject-title">Asignatura: <strong>{subject}</strong></div>'

            indicadores_info.sort(key=lambda x: x['pct'])
            indicators_html = '<div class="indicator-section"><h3>Nivel de alcance del indicador de logro (De menor a mayor alcance)</h3><div class="indicator-list">'
            for ind in indicadores_info:
                pct = ind['pct']
                if pct <= 20.0: c, txt = "color-red", "Crítico"
                elif pct <= 40.0: c, txt = "color-orange", "Bajo"
                elif pct <= 60.0: c, txt = "color-yellow", "Medio"
                elif pct <= 80.0: c, txt = "color-lightgreen", "Bueno"
                else: c, txt = "color-darkgreen", "Excelente"
                indicators_html += f'<div class="indicator-item {c}"><div class="indicator-pct"><div class="pct-number">{pct:.1f}%</div><div class="pct-label">{txt}</div></div><div><strong>{ind["disp_name"]}:</strong> {ind["indicador_base"]}</div></div>'
            indicators_html += '</div></div>' 
            
            # --- CONSTRUCCIÓN DE LA TABLA ---
            tbody_id = f"tbody_{nro_centro}_{subject}" 
            
            matrix_html = f"""
            <div style="display: flex; justify-content: flex-start; margin-bottom: 10px;" class="no-print">
                <button onclick="toggleSort(this, '{tbody_id}')" class="btn btn-default" data-sort-state="rendimiento" style="font-size: 13px;">
                    Ordenar los estudiantes de la tabla por orden alfabético
                </button>
            </div>
            <div class="table-container"><table><thead><tr>
            """
            
            headers_principales = ['Documento', 'Apellidos y Nombres', 'Puntaje (0-100)', 'Porcentaje de ítems correctos']
            headers_completos = headers_principales + item_cols
            
            for h in headers_completos:
                css_class = 'item-col' if h in item_cols else ''
                if h in item_cols:
                    indicador_limpio = f"{display_names[h]}: {item_indicators.get(h, '')}".replace('"', '&quot;').replace("'", "&#39;")
                    matrix_html += f'<th title="{indicador_limpio}" class="item-header {css_class}">{display_names[h]}</th>'
                else: 
                    matrix_html += f'<th class="{css_class}">{h}</th>'
                    
            matrix_html += f"</tr></thead><tbody id='{tbody_id}'>"
            
            for _, row in matrix.iterrows():
                puntaje_num = row['Puntaje_Num']
                tiempo_inusual = row['tiempo_inusual']
                puntaje_sort = row['Puntaje_Sort']
                
                nombre_str = row['Apellidos y Nombres']
                nombre_puro = nombre_str 
                
                if tiempo_inusual: nombre_str += " *"
                    
                bg_class = ""
                nivel_texto = ""
                
                if pd.notna(puntaje_num):
                    if puntaje_num <= 35: bg_class, nivel_texto = "bg-red", "Crítico"
                    elif puntaje_num <= 45: bg_class, nivel_texto = "bg-orange", "Bajo"
                    elif puntaje_num <= 55: bg_class, nivel_texto = "bg-yellow", "Medio"
                    elif puntaje_num <= 65: bg_class, nivel_texto = "bg-lightgreen", "Bueno"
                    else: bg_class, nivel_texto = "bg-darkgreen", "Excelente"

                matrix_html += f'<tr data-nombre="{nombre_puro}" data-puntaje="{puntaje_sort}">'
                
                for col_name in headers_completos:
                    css_class = 'item-col' if col_name in item_cols else ''
                    
                    if col_name == 'Apellidos y Nombres':
                        if nivel_texto:
                            matrix_html += f'<td class="{bg_class}"><div style="font-weight:bold;">{nombre_str}</div><div style="font-size:10px; color:#555; text-transform:uppercase;">Nivel: {nivel_texto}</div></td>'
                        else:
                            matrix_html += f'<td><div style="font-weight:bold;">{nombre_str}</div></td>'
                    elif col_name == 'Documento':
                        matrix_html += f"<td>{row['Documento']}</td>"
                    elif col_name == 'Puntaje (0-100)':
                        val_str = f"{puntaje_num:.1f}" if pd.notna(puntaje_num) else "-"
                        matrix_html += f'<td class="total-cell">{val_str}</td>'
                    elif col_name == 'Porcentaje de ítems correctos':
                        matrix_html += f'<td class="total-cell">{row["Porcentaje de ítems correctos"]}</td>'
                    elif col_name in item_cols:
                        val = row[col_name]
                        if float(val) == 1.0: matrix_html += f'<td class="{css_class}"><span class="correct">&#10003;</span></td>'
                        else: matrix_html += f'<td class="{css_class}"><span class="incorrect">&#10007;</span></td>'
                        
                matrix_html += "</tr>"
                
            matrix_html += "</tbody></table></div>"
            
            # --- LEYENDAS ---
            matrix_html += f"""
            <div class="footnote footnote-warning">
                <strong>* Resultados en revisión:</strong> El estudiante registró un tiempo de ejecución menor a 5 minutos o presentó un tiempo de demora inusual.
            </div>
            <div class="footnote footnote-info">
                <strong>** Nota sobre el Puntaje (0-100):</strong> En algunos casos, estudiantes con el mismo porcentaje de respuestas correctas pueden tener diferente puntaje de 0-100. Esto se debe a que el puntaje final considera el nivel de dificultad específico de los ítems respondidos por cada estudiante.
            </div>
            """
            
            # --- PANEL INFERIOR: DONA Y SELLO DE CALIDAD ---
            cat_counts = {'Crítico': 0, 'Bajo': 0, 'Medio': 0, 'Bueno': 0, 'Excelente': 0}
            for p in matrix['Puntaje_Num'].dropna():
                if p <= 35: cat_counts['Crítico'] += 1
                elif p <= 45: cat_counts['Bajo'] += 1
                elif p <= 55: cat_counts['Medio'] += 1
                elif p <= 65: cat_counts['Bueno'] += 1
                else: cat_counts['Excelente'] += 1
            
            tot_val = len(matrix['Puntaje_Num'].dropna())
            d_data = [
                (cat_counts['Crítico']/tot_val*100) if tot_val else 0,
                (cat_counts['Bajo']/tot_val*100) if tot_val else 0,
                (cat_counts['Medio']/tot_val*100) if tot_val else 0,
                (cat_counts['Bueno']/tot_val*100) if tot_val else 0,
                (cat_counts['Excelente']/tot_val*100) if tot_val else 0
            ]
            
            donut_id = f"donut_{subject}_{safe_grupo_file}_{nro_centro}"
            donut_data_str = ", ".join(f"{x:.1f}" for x in d_data)
            
            graficos_footer_html = f"""
            <div class="indicator-section" style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: space-around; align-items: center; page-break-inside: avoid; margin-bottom: 20px;">
                <div style="flex: 2; min-width: 350px; display: flex; flex-direction: column; align-items: center;">
                    <h3 style="text-align: center; margin-bottom: 5px; border: none; padding-bottom: 0;">Distribución de Estudiantes por Nivel de Alcance General</h3>
                    <p style="font-size: 12px; color: #7f8c8d; margin-top: 0; text-align: center;">Gráfica para los {total_grupo} estudiantes del grupo</p>
                    <div style="position: relative; height:300px; width:100%; display: flex; justify-content: center; margin-top: 15px;">
                        <canvas id="{donut_id}"></canvas>
                    </div>
                </div>
                
                <div style="flex: 1; min-width: 250px; display: flex; flex-direction: column; align-items: center; justify-content: center; border-left: 2px dashed #ecf0f1; padding-left: 20px;">
                    <h3 style="text-align: center; color: #2c3e50; margin-bottom: 20px; border: none;">Promedio de la Sección</h3>
                    <div class="badge-container" style="--badge-color: {badge_color};">
                        <div class="donut-seal">{group_mean_str}</div>
                        <div class="donut-label">{group_cat}</div>
                    </div>
                </div>
            </div>
            """
            
            script = f"""
                (function() {{
                    var ctx_{donut_id} = document.getElementById('{donut_id}').getContext('2d');
                    new Chart(ctx_{donut_id}, {{
                        type: 'doughnut',
                        data: {{
                            labels: ['Crítico (0-35)', 'Bajo (36-45)', 'Medio (46-55)', 'Bueno (56-65)', 'Excelente (66-100)'],
                            datasets: [{{
                                data: [{donut_data_str}],
                                backgroundColor: ['rgba(192, 57, 43, 0.7)', 'rgba(211, 84, 0, 0.7)', 'rgba(241, 196, 15, 0.7)', 'rgba(72, 201, 176, 0.7)', 'rgba(25, 111, 61, 0.7)'],
                                borderColor: ['#c0392b', '#d35400', '#f1c40f', '#1abc9c', '#196f3d'],
                                borderWidth: 1,
                                datalabels: {{
                                    color: '#2c3e50',
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
            
            tablas_html_combinadas += subject_html + indicators_html + matrix_html + graficos_footer_html + '</div>'

        if tablas_html_combinadas:
            safe_grado = str(grado).replace(" ", "").replace("º", "")
            filename = f"Reporte_{nro_centro}_Grado{safe_grado}_Grupo{safe_grupo_file}.html"
            filepath = os.path.abspath(os.path.join(school_folder_path, filename))
            
            file_uri = pathlib.Path(filepath).as_uri()
            depto_seguro = departamento if pd.notnull(departamento) else "Sin Especificar"
            
            all_scripts = "\n".join(chart_scripts)
            
            header_info = {
                'mes_prueba': mes_prueba,
                'mes_aplicacion': mes_aplicacion,
                'departamento': depto_seguro,
                'nro_centro': nro_centro,
                'centro': centro_nombre_str,
                'grado': str(grado).replace(" ", "").replace("º", ""),
                'grupo': grupo_original,
                'file_uri': file_uri,
                'tablas_html': tablas_html_combinadas,
                'chart_scripts': all_scripts
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(HTML_TEMPLATE.format(**header_info))
                
            reportes_generados += 1

    print(f"\n¡Proceso Terminado! Se generaron {reportes_generados} reportes listos para revisar e imprimir.")

if __name__ == "__main__":
    process_section_reports()