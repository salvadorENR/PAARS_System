# src/report_directores.py
import os
import pandas as pd
import re
import numpy as np
import pathlib
import sys

# --- 1. CONEXIÓN CON EL MAPA PAARS ---
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# Usamos las rutas dinámicas del sistema
PREV_MONTH_DIR = config.PATH_PREV_INTERIM
CURR_MONTH_DIR = config.PATH_INTERIM

# Se guardará exactamente al lado de Reportes_Por_Secciones, dentro de Final_Reports
REPORTS_DIR = os.path.join(config.PATH_REPORTS, "Reportes_Por_Escuela")

os.makedirs(REPORTS_DIR, exist_ok=True)

# --- 2. PLANTILLA HTML Y CSS ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resultados Prueba de Progreso - {centro}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <style>
        * {{
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }}
        
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }}
        
        .controls {{ display: flex; justify-content: flex-end; gap: 10px; margin-bottom: 20px; }}
        .btn {{ padding: 10px 15px; border: 1px solid #bdc3c7; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold; transition: all 0.2s; }}
        .btn-primary {{ background: #2980b9; color: white; border-color: #2471a3; }}
        .btn-primary:hover {{ background: #1f618d; }}
        
        .pdf-link-container {{ text-align: center; margin-bottom: 25px; }}
        .pdf-link-container a {{ color: #2980b9; font-weight: bold; text-decoration: underline; font-size: 14px; }}

        .header-card {{ background-color: #ffffff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; border-left: 5px solid #2980b9; }}
        .header-card h2 {{ margin: 0 0 15px 0; color: #2c3e50; font-size: 24px; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; }}
        .info-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; font-size: 14px; color: #444; }}
        .info-item strong {{ color: #2c3e50; display: block; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
        .info-item span {{ font-size: 15px; font-weight: 500; }}
        
        .footer-card {{ background-color: #ffffff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 30px; display: flex; flex-direction: column; align-items: center; border-top: 5px solid #2c3e50; page-break-inside: avoid; }}
        .footer-card h3 {{ margin: 0 0 15px 0; color: #2c3e50; font-size: 18px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center; }}
        .badge-container {{ display: flex; flex-direction: column; align-items: center; justify-content: center; }}
        .donut-seal {{ width: 85px; height: 85px; border-radius: 50%; border: 14px solid var(--badge-color); display: flex; align-items: center; justify-content: center; font-size: 24px; font-weight: 900; color: #2c3e50; background: #fff; box-shadow: inset 0 2px 4px rgba(0,0,0,0.1); }}
        .donut-label {{ margin-top: 15px; font-size: 16px; font-weight: bold; color: var(--badge-color); text-transform: uppercase; text-align: center; background: rgba(0,0,0,0.03); padding: 6px 15px; border-radius: 15px; border: 1px solid var(--badge-color); }}
        
        .subject-section {{ margin-bottom: 50px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
        .subject-title {{ font-size: 20px; color: #2980b9; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; margin-bottom: 15px; text-transform: uppercase; font-weight: bold; }}
        
        .chart-container {{ position: relative; height: 250px; width: 100%; margin-bottom: 30px; }}
        
        .group-row {{ display: flex; align-items: stretch; margin-bottom: 8px; page-break-inside: avoid; }}
        
        .group-data {{ display: flex; width: 40%; justify-content: space-between; align-items: center; padding: 12px 15px; border-radius: 4px; }}
        .group-faja {{ width: 57%; display: flex; align-items: center; padding: 12px 15px; border-radius: 4px; background-color: #ffffff; border: 1px solid #ecf0f1; box-sizing: border-box; }}
        .spacer {{ width: 3%; }}
        
        .data-col {{ flex: 1; font-size: 14px; }}
        .text-left {{ text-align: left; }}
        .text-center {{ text-align: center; }}
        
        .header-dark {{ background-color: #34495e; color: white; font-weight: bold; border: none; padding: 10px 15px; }}
        
        .desc-text {{ font-size: 16px; color: #2c3e50; margin-bottom: 5px; font-weight: bold; }}
        .sub-text {{ font-size: 13px; color: #7f8c8d; margin-bottom: 15px; margin-top: 0; }}
        
        @media print {{
            .no-print {{ display: none !important; }}
            body {{ background-color: white; padding: 0; }}
            .header-card, .subject-section, .footer-card {{ box-shadow: none; border: 1px solid #ddd; break-inside: avoid; margin-bottom: 20px; }}
            .chart-container {{ page-break-inside: avoid; height: 250px !important; display: block !important; width: 100% !important; }}
            canvas {{ height: 250px !important; width: 100% !important; }}
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
                <div class="info-item"><strong>Mes de Aplicación</strong><span>{mes_aplicacion}</span></div>
            </div>
        </div>
    </div>
    
    <div class="pdf-link-container">
        <a href="{file_uri}" target="_blank">Link para visualizar el html (Versión Interactiva)</a>
    </div>

    {tablas_html}

    <div class="footer-card">
        <h3>Media de los estudiantes de toda la escuela</h3>
        <div class="badge-container" style="--badge-color: {badge_color};">
            <div class="donut-seal">{global_mean}</div>
            <div class="donut-label">{global_cat}</div>
        </div>
    </div>

    <script>
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

def calcular_pcts_100(df_estudiantes, niveles):
    tot = len(df_estudiantes)
    if tot == 0: return [0.0, 0.0, 0.0, 0.0, 0.0], 0
    pcts = []
    for nivel in niveles:
        cnt = len(df_estudiantes[df_estudiantes['Clasificacion'] == nivel])
        pcts.append((cnt / tot) * 100)
    
    suma_actual = sum(pcts)
    if abs(suma_actual - 100) > 0.1: pcts[-1] += (100 - suma_actual)
    return pcts, tot

def generar_barra_css(pcts):
    colores = ["#c0392b", "#d35400", "#f1c40f", "#1abc9c", "#196f3d"]
    text_colors = ["white", "white", "#333", "white", "white"]
    html = '<div style="display: flex; height: 28px; width: 100%; border-radius: 4px; overflow: hidden; background-color: #ffffff; border: 1px solid #bdc3c7;">'
    for i, p in enumerate(pcts):
        if p > 0:
            text = f"{p:.1f}%" if p >= 5 else ""
            html += f'<div style="width: {p}%; background-color: {colores[i]}; color: {text_colors[i]}; font-size: 11px; display: flex; align-items: center; justify-content: center; font-weight: bold;" title="{p:.1f}%">{text}</div>'
    html += '</div>'
    return html

def read_geiser_folder(folder_path):
    all_dfs = []
    if not os.path.exists(folder_path): 
        print(f"  [!] No se encontraron datos en: {folder_path}")
        return pd.DataFrame()
    for f in os.listdir(folder_path):
        f_lower = f.lower()
        if f_lower.endswith('.csv') and 'resultados' in f_lower and '_compl' not in f_lower and 'legend' not in f_lower:
            try:
                df = pd.read_csv(os.path.join(folder_path, f), dtype=str, encoding='utf-8-sig')
                if 'MAT-' in f.upper(): df['Area temática'] = 'Matemática'
                elif 'LEC-' in f.upper(): df['Area temática'] = 'Lengua'
                elif 'Area temática' not in df.columns: df['Area temática'] = 'Desconocida'
                
                columnas_vitales = ['Documento', 'Grado', 'Grupo', 'theta.global (escala 0-100)']
                if set(columnas_vitales).issubset(df.columns):
                    extras = ['Departamento', 'Nro de centro', 'Centro', 'Prueba', 'Fecha-Hora de Inicio']
                    for col in extras:
                        if col not in df.columns: df[col] = 'Sin Especificar'
                    cols_a_guardar = columnas_vitales + extras + ['Area temática']
                    all_dfs.append(df[cols_a_guardar])
            except Exception as e:
                pass
                
    if all_dfs:
        res = pd.concat(all_dfs, ignore_index=True)
        res['Documento'] = res['Documento'].astype(str).str.strip()
        res['Nro de centro'] = res['Nro de centro'].astype(str).str.strip()
        res['Grupo_Completo'] = res['Grupo'].astype(str).str.strip()
        res['theta.global (escala 0-100)'] = pd.to_numeric(res['theta.global (escala 0-100)'], errors='coerce')
        res = res.dropna(subset=['theta.global (escala 0-100)'])
        res = res.drop_duplicates(subset=['Documento', 'Area temática'])
        res['Clasificacion'] = res['theta.global (escala 0-100)'].apply(clasificar_puntaje)
        return res
    return pd.DataFrame()

# --- 4. PROCESAMIENTO PRINCIPAL ---
def process_comparative_reports():
    print("1. Cargando CSVs de Analisis Psicometrico del Mes Anterior...")
    df_prev = read_geiser_folder(PREV_MONTH_DIR)
    print("2. Cargando CSVs de Analisis Psicometrico del Mes Actual...")
    df_curr = read_geiser_folder(CURR_MONTH_DIR)
    
    if df_curr.empty:
        print("ERROR: No hay datos válidos del mes actual para procesar.")
        return

    print(f"3. Generando reportes por centro escolar en: {REPORTS_DIR}...")
    grouped_curr = df_curr.groupby(['Nro de centro', 'Centro', 'Departamento'])

    niveles = ["Crítica", "Baja", "Media", "Buena", "Excelente"]
    colores_hex = ["#c0392b", "#d35400", "#f1c40f", "#1abc9c", "#196f3d"]
    
    bg_colors = {
        "Crítica": "rgba(192, 57, 43, 0.15)", "Baja": "rgba(211, 84, 0, 0.15)", 
        "Media": "rgba(241, 196, 15, 0.15)", "Buena": "rgba(26, 188, 156, 0.15)", 
        "Excelente": "rgba(25, 111, 61, 0.15)"
    }
    solid_colors = {
        "Crítica": "#c0392b", "Baja": "#d35400", "Media": "#f1c40f", 
        "Buena": "#1abc9c", "Excelente": "#196f3d"
    }
    
    meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}

    reportes_generados = 0

    for (centro_id, centro_nombre, departamento), group_data_curr in grouped_curr:
        tablas_html_combinadas = ""
        chart_scripts = []
        subjects_order = ['Matemática', 'Lengua']
        
        mes_prueba = "2" 
        pruebas_unicas = group_data_curr['Prueba'].dropna().unique()
        for p in pruebas_unicas:
            match = re.search(r'Mes\s*(\d+)', str(p), re.IGNORECASE)
            if match:
                mes_prueba = match.group(1)
                break
                
        # SILENCIADOR DE FECHAS: format='mixed'
        fechas_validas = pd.to_datetime(group_data_curr['Fecha-Hora de Inicio'].replace('Sin Especificar', pd.NA), errors='coerce', dayfirst=True, format='mixed')
        
        if not fechas_validas.dropna().empty:
            mes_num = int(fechas_validas.dropna().dt.month.mode().iloc[0])
            mes_aplicacion = meses_es.get(mes_num, "Mes Actual")
            prev_mes_num = 12 if mes_num == 1 else mes_num - 1
            mes_anterior_str = meses_es.get(prev_mes_num, "Mes Anterior")
        else:
            mes_aplicacion = "Mes Actual"
            mes_anterior_str = "Mes Anterior"

        escuela_mean_val = group_data_curr['theta.global (escala 0-100)'].mean()
        if pd.notnull(escuela_mean_val):
            global_cat = clasificar_puntaje(escuela_mean_val)
            global_mean_str = f"{escuela_mean_val:.1f}"
            badge_color = solid_colors.get(global_cat, "#bdc3c7")
        else:
            global_cat = "Sin Datos"
            global_mean_str = "-"
            badge_color = "#bdc3c7"

        for subject in subjects_order:
            curr_subject_data = group_data_curr[group_data_curr['Area temática'] == subject].copy()
            if curr_subject_data.empty: continue
            
            prev_subject_data = pd.DataFrame()
            if not df_prev.empty:
                prev_subject_data = df_prev[(df_prev['Nro de centro'] == centro_id) & (df_prev['Area temática'] == subject)].copy()

            pcts_prev, tot_prev = calcular_pcts_100(prev_subject_data, niveles)
            pcts_curr, tot_curr = calcular_pcts_100(curr_subject_data, niveles)

            chart_id = f"chart_{centro_id}_{subject}".replace(" ", "_").replace("-", "_").replace(".", "_")
            
            html_table = f"""
            <div class="subject-section">
                <div class="subject-title">Resultados de {subject}</div>
                <p class="desc-text">Comparativa Institucional Global - {subject}</p>
                <div class="chart-container">
                    <canvas id="{chart_id}"></canvas>
                </div>
                <p class="desc-text" style="margin-top: 30px; margin-bottom: 15px;">Detalle del Mes Actual por Grupo Escolar</p>
                
                <div class="group-row" style="margin-bottom: 10px;">
                    <div class="group-data header-dark" style="border-radius: 4px;">
                        <div class="data-col text-left">Grado - Grupo</div>
                        <div class="data-col text-center">Clasificación</div>
                        <div class="data-col text-center">Promedio</div>
                    </div>
                    <div class="spacer"></div>
                    <div class="group-faja header-dark" style="justify-content: center; border-radius: 4px; border: none;">
                        Distribución de Estudiantes
                    </div>
                </div>
            """
            
            grupos_unicos = curr_subject_data[['Grado', 'Grupo_Completo']].drop_duplicates().copy()
            grupos_unicos['Grado_Num'] = grupos_unicos['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
            grupos_unicos = grupos_unicos.sort_values(by=['Grado_Num', 'Grupo_Completo'])
            
            for _, row_grupo in grupos_unicos.iterrows():
                grado_actual = row_grupo['Grado']
                grupo_id_actual = row_grupo['Grupo_Completo']
                grado_num = row_grupo['Grado_Num'] 
                
                estudiantes_grupo = curr_subject_data[
                    (curr_subject_data['Grado'] == grado_actual) & 
                    (curr_subject_data['Grupo_Completo'] == grupo_id_actual)
                ]
                
                total_grupo = len(estudiantes_grupo)
                if total_grupo == 0: continue
                
                scores = estudiantes_grupo['theta.global (escala 0-100)']
                mean_score = scores.mean()
                mean_str = f"{mean_score:.1f}"
                clasif_mean = clasificar_puntaje(mean_score)
                
                bg_c = bg_colors.get(clasif_mean, "#ecf0f1")
                solid_c = solid_colors.get(clasif_mean, "#bdc3c7")
                
                pcts_row, _ = calcular_pcts_100(estudiantes_grupo, niveles)
                barra_css_html = generar_barra_css(pcts_row)
                
                grupo_mostrar = str(grupo_id_actual).strip()
                
                html_table += f"""
                <div class="group-row">
                    <div class="group-data" style="background-color: {bg_c}; border-left: 6px solid {solid_c};">
                        <div class="data-col text-left"><strong>{grado_num}° - {grupo_mostrar}</strong></div>
                        <div class="data-col text-center" style="color: black; font-weight: bold; text-transform: uppercase;">{clasif_mean}</div>
                        <div class="data-col text-center"><strong>{mean_str}</strong></div>
                    </div>
                    <div class="spacer"></div>
                    <div class="group-faja">
                        {barra_css_html}
                    </div>
                </div>
                """
            
            html_table += "</div>"
            tablas_html_combinadas += html_table
            
            script = f"""
                (function() {{
                    var ctx = document.getElementById('{chart_id}').getContext('2d');
                    if (!ctx) return;
                    new Chart(ctx, {{
                        type: 'bar',
                        data: {{
                            labels: ['{mes_anterior_str} (N = {tot_prev})', '{mes_aplicacion} (N = {tot_curr})'],
                            datasets: [
                                {{ label: 'Crítica (0-35)', data: [{pcts_prev[0]:.1f}, {pcts_curr[0]:.1f}], backgroundColor: '{colores_hex[0]}' }},
                                {{ label: 'Baja (>35-45)', data: [{pcts_prev[1]:.1f}, {pcts_curr[1]:.1f}], backgroundColor: '{colores_hex[1]}' }},
                                {{ label: 'Media (>45-55)', data: [{pcts_prev[2]:.1f}, {pcts_curr[2]:.1f}], backgroundColor: '{colores_hex[2]}' }},
                                {{ label: 'Buena (>55-65)', data: [{pcts_prev[3]:.1f}, {pcts_curr[3]:.1f}], backgroundColor: '{colores_hex[3]}' }},
                                {{ label: 'Excelente (>65-100)', data: [{pcts_prev[4]:.1f}, {pcts_curr[4]:.1f}], backgroundColor: '{colores_hex[4]}' }}
                            ]
                        }},
                        options: {{
                            animation: false,
                            responsive: true, maintainAspectRatio: false, indexAxis: 'y',
                            scales: {{
                                x: {{ stacked: true, min: 0, max: 100, title: {{ display: true, text: 'Porcentaje de Estudiantes (%)' }}, ticks: {{ callback: function(v) {{ return v + '%'; }}, stepSize: 10 }} }},
                                y: {{ stacked: true, title: {{ display: true, text: 'Período de Evaluación' }} }}
                            }},
                            plugins: {{
                                tooltip: {{ callbacks: {{ label: function(c) {{ return c.dataset.label + ': ' + c.raw.toFixed(1) + '%'; }} }} }},
                                datalabels: {{
                                    display: true,
                                    color: function(c) {{ return c.dataset.label.includes('Media') ? '#000000' : '#ffffff'; }},
                                    font: {{ weight: 'bold', size: 12 }},
                                    formatter: function(v) {{ return v > 3 ? v.toFixed(1) + '%' : ''; }},
                                    anchor: 'center', align: 'center', offset: 4
                                }}
                            }}
                        }}
                    }});
                }})();
            """
            chart_scripts.append(script)
        
        if tablas_html_combinadas:
            centro_nombre_str = str(centro_nombre).strip()
            centro_nombre_seguro = "".join([c for c in centro_nombre_str if c not in r'\/:*?"<>|'])
            nombre_carpeta = f"{centro_id} - {centro_nombre_seguro}"
            school_folder = os.path.join(REPORTS_DIR, nombre_carpeta)
            os.makedirs(school_folder, exist_ok=True)
            filename = f"Reporte_Comparativo_{centro_id}.html"
            filepath = os.path.abspath(os.path.join(school_folder, filename))
            
            depto_seguro = departamento if pd.notnull(departamento) else "Sin Especificar"
            all_scripts = "\n".join(chart_scripts)
            
            file_uri = pathlib.Path(filepath).as_uri()
            
            header_info = {
                'mes_prueba': mes_prueba,
                'mes_aplicacion': mes_aplicacion,
                'departamento': depto_seguro,
                'nro_centro': centro_id,
                'centro': centro_nombre_str,
                'global_mean': global_mean_str,
                'global_cat': global_cat,
                'badge_color': badge_color,
                'file_uri': file_uri,
                'tablas_html': tablas_html_combinadas,
                'chart_scripts': all_scripts
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(HTML_TEMPLATE.format(**header_info))
            
            reportes_generados += 1

    print(f"\n¡Listo! Se generaron {reportes_generados} reportes comparativos en:\n{REPORTS_DIR}")

if __name__ == "__main__":
    process_comparative_reports()