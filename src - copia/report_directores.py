# src/report_directores.py
import os
import pandas as pd
import json
import config

HTML_BASE = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resultados Prueba de Progreso - {centro}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0/dist/chartjs-plugin-datalabels.min.js"></script>
    <style>
        * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }}
        
        .controls {{ display: flex; justify-content: flex-end; gap: 10px; margin-bottom: 20px; }}
        .btn {{ padding: 10px 15px; border: 1px solid #bdc3c7; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold; transition: all 0.2s; }}
        .btn-primary {{ background: #2980b9; color: white; border-color: #2471a3; }}
        .btn-primary:hover {{ background: #1f618d; }}
        
        .header-card {{ background-color: #ffffff; padding: 25px 30px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); margin-bottom: 25px; border-left: 6px solid #2980b9; }}
        .header-card h2 {{ margin: 0 0 15px 0; color: #2c3e50; font-size: 22px; text-transform: uppercase; font-weight: 700; border-bottom: 1px solid #ecf0f1; padding-bottom: 12px; }}
        .info-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; align-items: start; }}
        .info-item {{ display: flex; flex-direction: column; }}
        .info-item strong {{ color: #2c3e50; font-size: 11px; text-transform: uppercase; font-weight: 700; margin-bottom: 4px; letter-spacing: 0.5px; }}
        .info-item span {{ font-size: 14px; color: #333; font-weight: 500; }}
        
        .footer-card {{ background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 30px; display: flex; flex-direction: column; align-items: center; border-top: 5px solid #2c3e50; page-break-inside: avoid; }}
        .footer-card h3 {{ margin: 0 0 15px 0; color: #2c3e50; font-size: 18px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center; }}
        
        .subject-section {{ margin-bottom: 50px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
        .subject-title {{ font-size: 20px; color: #2980b9; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; margin-bottom: 15px; text-transform: uppercase; font-weight: bold; }}
        
        .chart-container {{ position: relative; height: 250px; width: 100%; margin-bottom: 30px; margin-top: 15px; }}
        
        /* --- LAYOUT GRÁFICA + TABLA HISTÓRICA MODIFICADA PARA NO TENER SCROLL --- */
        .history-wrapper {{ display: flex; gap: 20px; align-items: flex-start; width: 100%; max-width: 1000px; margin: auto; margin-top: 15px; }}
        .history-chart-container {{ flex: 1.6; position: relative; min-height: 350px; }}
        .history-table-container {{ flex: 1; background: #fff; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); border: 1px solid #ecf0f1; }}
        .history-table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
        .history-table th, .history-table td {{ border: 1px solid #ecf0f1; padding: 10px 6px; text-align: center; }}
        .history-table th {{ background-color: #34495e; color: white; text-transform: uppercase; font-size: 11px; }}
        .history-table .mat-col {{ background-color: rgba(41, 128, 185, 0.05); }}
        .history-table .lec-col {{ background-color: rgba(192, 57, 43, 0.05); }}
        .history-table td strong {{ color: #2c3e50; }}
        
        .group-row {{ display: flex; align-items: stretch; margin-bottom: 8px; page-break-inside: avoid; }}
        .group-data {{ display: flex; width: 35%; justify-content: space-between; align-items: center; padding: 12px 15px; border-radius: 4px; }}
        .group-faja {{ width: 62%; display: flex; align-items: center; padding: 12px 15px; border-radius: 4px; background-color: #ffffff; border: 1px solid #ecf0f1; box-sizing: border-box; }}
        .spacer {{ width: 3%; }}
        
        .data-col {{ flex: 1; font-size: 14px; }}
        .text-left {{ text-align: left; }}
        .text-center {{ text-align: center; }}
        .text-right {{ text-align: right; }}
        
        .header-dark {{ background-color: #34495e; color: white; font-weight: bold; border: none; padding: 10px 15px; margin-top: 15px; }}
        .desc-text {{ font-size: 16px; color: #2c3e50; margin-bottom: 15px; font-weight: bold; margin-top: 25px; }}
        
        details.info-accordion {{ margin-bottom: 15px; text-align: left; display: block; }}
        details.info-accordion summary {{ cursor: pointer; font-size: 13px; font-weight: bold; color: #2980b9; padding: 8px 12px; background-color: #f4f6f9; border-radius: 4px; border: 1px solid #d0e1f9; display: inline-block; list-style: none; transition: 0.2s; }}
        details.info-accordion summary::-webkit-details-marker {{ display: none; }}
        details.info-accordion summary:hover {{ background-color: #eaf2f8; }}
        details.info-accordion p {{ padding: 15px; border-left: 3px solid #2980b9; background-color: #fafafa; margin-top: 8px; font-size: 13.5px; color: #444; border-radius: 0 4px 4px 0; box-shadow: inset 0 1px 3px rgba(0,0,0,0.03); line-height: 1.5; margin-bottom: 0; }}
        
        @media print {{
            .no-print {{ display: none !important; }}
            body {{ background-color: white; padding: 0; }}
            .header-card, .subject-section, .footer-card {{ box-shadow: none; border: 1px solid #ddd; break-inside: avoid; margin-bottom: 20px; }}
            .history-wrapper {{ display: flex; gap: 15px; flex-wrap: nowrap; }}
            .chart-container, .history-chart-container {{ page-break-inside: avoid; display: block !important; width: 100% !important; }}
            canvas {{ max-height: 350px !important; width: 100% !important; }}
        }}
    </style>
</head>
<body>
    <div class="controls no-print">
        <button onclick="window.print()" class="btn btn-primary">Descargar como PDF</button>
    </div>

    <div class="header-card">
        <h2>RESULTADOS DE LA PRUEBA DE PROGRESO - {str_mes}{mes_actual_mayus} 2026</h2>
        <div class="info-grid">
            <div class="info-item"><strong>DEPARTAMENTO</strong><span>{departamento}</span></div>
            <div class="info-item"><strong>CÓDIGO DE INFRAESTRUCTURA</strong><span>{codigo_infra}</span></div>
            <div class="info-item"><strong>CENTRO ESCOLAR</strong><span>{centro}</span></div>
            <div class="info-item"><strong>MES DE APLICACIÓN</strong><span>{mes_actual}</span></div>
        </div>
    </div>

    {contenido}

    <div class="footer-card">
        <h3>Evolución de los niveles de toda la escuela</h3>
        
        <details open class="info-accordion no-print" style="width: 100%; max-width: 1000px;">
            <summary>📖 Ocultar / Mostrar descripción de la gráfica histórica</summary>
            <p>Esta sección muestra la evolución mensual del desempeño en Matemática y Lengua. La gráfica permite visualizar la tendencia general, y la tabla presenta los puntajes promedio y el nivel alcanzado en cada mes.</p>
        </details>

        <div class="history-wrapper">
            <div class="history-chart-container">
                <canvas id="historical_line_chart"></canvas>
            </div>
            {tabla_historial}
        </div>
    </div>

    <script>
        Chart.register(ChartDataLabels);
        document.addEventListener('DOMContentLoaded', function() {{
            {scripts_graficos}
        }});
    </script>
</body>
</html>
"""

# Configuración de Niveles (Colores actualizados por Pruebas de Accesibilidad)
NIVELES_CONFIG = [
    {'limite': 35, 'nombre': 'Crítico',   'color': '#991b1b', 'bg_color': '#fef2f2', 'texto': 'white'},
    {'limite': 45, 'nombre': 'Bajo',      'color': '#ff8c2e', 'bg_color': '#fff7ed', 'texto': 'white'},
    {'limite': 55, 'nombre': 'Medio',     'color': '#facc15', 'bg_color': '#fefce8', 'texto': '#333'},
    {'limite': 65, 'nombre': 'Bueno',     'color': '#84cc16', 'bg_color': '#f7fee7', 'texto': '#333'},
    {'limite': 100,'nombre': 'Excelente', 'color': '#065f46', 'bg_color': '#ecfdf5', 'texto': 'white'}
]

# Mapa para Meses de Progreso
MESES_PROGRESO = {
    'MARZO': 1, 'ABRIL': 2, 'MAYO': 3, 'JUNIO': 4, 
    'JULIO': 5, 'AGOSTO': 6, 'SEPTIEMBRE': 7, 'OCTUBRE': 8, 'NOVIEMBRE': 9
}

def obtener_clasificacion(score):
    if pd.isna(score): return None
    for n in NIVELES_CONFIG:
        if score <= n['limite']: return n
    return NIVELES_CONFIG[-1]

def generar_barra_html(distribucion):
    html = '<div style="display: flex; height: 28px; width: 100%; border-radius: 4px; overflow: hidden; background-color: #ffffff; border: 1px solid #bdc3c7;">'
    for nivel in NIVELES_CONFIG:
        nombre = nivel['nombre']
        pct = distribucion.get(nombre, 0)
        if pct > 0:
            texto = f"{pct:.1f}%" if pct >= 5 else "" 
            html += f'<div style="width: {pct}%; background-color: {nivel["color"]}; color: {nivel["texto"]}; font-size: 11px; display: flex; align-items: center; justify-content: center; font-weight: bold;" title="{pct:.1f}%">{texto}</div>'
    html += '</div>'
    return html

def cargar_thetas(ruta_csvs, df_actual, mes_nombre):
    archivos = [f for f in os.listdir(ruta_csvs) if f.endswith('.csv') and not f.startswith('legend')]
    lista_dfs = []
    
    for archivo in archivos:
        try:
            df = pd.read_csv(os.path.join(ruta_csvs, archivo), dtype=str, encoding_errors='ignore')
            materia = 'Matemática' if 'MAT' in archivo.upper() else 'Lengua'
            df['Materia'] = materia
            
            col_doc = next((c for c in df.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), None)
            if col_doc and 'theta.global (escala 0-100)' in df.columns:
                df = df[[col_doc, 'theta.global (escala 0-100)', 'Materia']].copy()
                df.rename(columns={col_doc: 'Documento'}, inplace=True)
                df['Documento'] = df['Documento'].str.strip().str.replace('"', '', regex=False)
                df['theta.global (escala 0-100)'] = pd.to_numeric(df['theta.global (escala 0-100)'].str.replace(',', '.'), errors='coerce')
                lista_dfs.append(df)
        except: continue
        
    if not lista_dfs: return pd.DataFrame()
    df_thetas = pd.concat(lista_dfs, ignore_index=True)
    df_thetas = df_thetas.dropna(subset=['theta.global (escala 0-100)'])
    
    df_merged = pd.merge(
        df_actual[['Documento', 'Centro', 'Materia']],
        df_thetas,
        on=['Documento', 'Materia'],
        how='inner'
    )
    
    df_merged['Categoria'] = df_merged['theta.global (escala 0-100)'].apply(lambda x: obtener_clasificacion(x)['nombre'] if pd.notna(x) else 'N/A')
    return df_merged

def construir_historial(df_actual):
    print("  [Directores] Rastreador Histórico: Analizando meses anteriores...")
    historial = []
    mes_actual_nombre = config.MONTH_FOLDER.split('_')[-1]
    
    try:
        base_dir = os.path.dirname(os.path.dirname(config.PATH_INTERIM))
        current_folder = os.path.basename(os.path.dirname(config.PATH_INTERIM))
        
        if os.path.exists(base_dir):
            carpetas = sorted([d for d in os.listdir(base_dir) if "PROGRESO" in d.upper()])
            for carpeta in carpetas:
                mes_nombre = carpeta.split('_')[-1]
                if carpeta < current_folder:
                    ruta_csvs = os.path.join(base_dir, carpeta, "Interim_CSVs")
                    if os.path.exists(ruta_csvs):
                        df_hist = cargar_thetas(ruta_csvs, df_actual, mes_nombre)
                        if not df_hist.empty:
                            historial.append({'mes': mes_nombre, 'df': df_hist})
                elif carpeta == current_folder:
                    historial.append({'mes': mes_actual_nombre, 'df': df_actual})
                    break
    except Exception as e:
        print(f"  [!] Error rastreando historial: {e}")
        
    if not historial or historial[-1]['mes'] != mes_actual_nombre:
        historial.append({'mes': mes_actual_nombre, 'df': df_actual})
        
    return historial

def process_comparative_reports(df_master):
    print("  [Directores] Preparando base de datos...")
    df_actual = df_master.copy()
    
    col_doc_act = next((c for c in df_actual.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), None)
    if col_doc_act: df_actual.rename(columns={col_doc_act: 'Documento'}, inplace=True)
    df_actual['Documento'] = df_actual['Documento'].astype(str).str.strip().str.replace('"', '', regex=False)
    
    if 'Area temática' in df_actual.columns: df_actual.rename(columns={'Area temática': 'Materia'}, inplace=True)
    if 'Materia' in df_actual.columns: df_actual['Materia'] = df_actual['Materia'].apply(lambda x: 'Matemática' if 'MAT' in str(x).upper() else 'Lengua')
        
    col_grupo = [c for c in df_actual.columns if 'rup' in str(c).lower() or 'ecc' in str(c).lower()]
    if col_grupo: df_actual.rename(columns={col_grupo[0]: 'Grupo'}, inplace=True)
    else: df_actual['Grupo'] = "A"
        
    col_theta_actual = df_actual.get('theta.global (escala 0-100)', pd.Series(dtype=str))
    if isinstance(col_theta_actual, pd.Series):
        df_actual['theta.global (escala 0-100)'] = col_theta_actual.astype(str).str.replace('"', '').str.replace(',', '.')
    
    df_actual['theta.global (escala 0-100)'] = pd.to_numeric(df_actual['theta.global (escala 0-100)'], errors='coerce')
    df_actual = df_actual.dropna(subset=['theta.global (escala 0-100)'])
    
    df_actual['Grado_Num'] = df_actual['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
    df_actual['Categoria'] = df_actual['theta.global (escala 0-100)'].apply(lambda x: obtener_clasificacion(x)['nombre'] if pd.notna(x) else 'N/A')
    
    col_dep = next((c for c in df_actual.columns if 'depar' in str(c).lower()), None)
    col_cod = next((c for c in df_actual.columns if 'cód' in str(c).lower() or 'cod' in str(c).lower() or 'nro' in str(c).lower() or 'infra' in str(c).lower()), None)

    historial = construir_historial(df_actual)
    tiene_anterior = len(historial) >= 2

    output_dir = os.path.join(config.PATH_REPORTS, "Reportes_Por_Escuela")
    os.makedirs(output_dir, exist_ok=True)
    
    escuelas = [e for e in df_actual['Centro'].dropna().unique() if str(e).strip().upper() != "DESCONOCIDO"]
    mes_nombre = config.MONTH_FOLDER.split('_')[-1]
    
    mes_numero = MESES_PROGRESO.get(mes_nombre.upper(), 0)
    str_mes = f"MES {mes_numero}, " if mes_numero > 0 else ""
    
    for escuela in escuelas:
        df_esc_act = df_actual[df_actual['Centro'] == escuela]
        
        dep_val = str(df_esc_act[col_dep].dropna().iloc[0]) if col_dep and not df_esc_act[col_dep].dropna().empty else "N/D"
        cod_val = str(df_esc_act[col_cod].dropna().iloc[0]) if col_cod and not df_esc_act[col_cod].dropna().empty else "N/D"

        # CÁLCULOS HISTÓRICOS PARA LA LÍNEA Y LA TABLA
        labels_hist, data_mat_hist, data_lec_hist = [], [], []
        levels_mat_hist, levels_lec_hist = [], []
        history_rows = ""
        
        for item in historial:
            mes_h = item['mes'].upper()
            labels_hist.append(f"'{mes_h}'")
            df_h = item['df'][item['df']['Centro'] == escuela]
            
            m_mat_h = df_h[df_h['Materia'] == 'Matemática']['theta.global (escala 0-100)'].mean()
            m_lec_h = df_h[df_h['Materia'] == 'Lengua']['theta.global (escala 0-100)'].mean()
            
            data_mat_hist.append(f"{m_mat_h:.1f}" if pd.notna(m_mat_h) else "null")
            data_lec_hist.append(f"{m_lec_h:.1f}" if pd.notna(m_lec_h) else "null")
            
            c_mat = obtener_clasificacion(m_mat_h)['nombre'] if pd.notna(m_mat_h) else '-'
            c_lec = obtener_clasificacion(m_lec_h)['nombre'] if pd.notna(m_lec_h) else '-'
            
            levels_mat_hist.append(f"'{c_mat}'")
            levels_lec_hist.append(f"'{c_lec}'")
            
            val_mat_str = f"{m_mat_h:.1f}" if pd.notna(m_mat_h) else "-"
            val_lec_str = f"{m_lec_h:.1f}" if pd.notna(m_lec_h) else "-"
            
            history_rows += f"<tr><td><strong>{mes_h}</strong></td><td class='mat-col'>{val_mat_str}</td><td class='mat-col'>{c_mat}</td><td class='lec-col'>{val_lec_str}</td><td class='lec-col'>{c_lec}</td></tr>"

        history_table = f"""
        <div class="history-table-container">
            <table class="history-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="background-color: #2c3e50;">Mes</th>
                        <th colspan="2" style="background-color: #2980b9;">Matemática</th>
                        <th colspan="2" style="background-color: #c0392b;">Lengua</th>
                    </tr>
                    <tr>
                        <th style="background-color: #34495e;">Media</th><th style="background-color: #34495e;">Nivel</th>
                        <th style="background-color: #34495e;">Media</th><th style="background-color: #34495e;">Nivel</th>
                    </tr>
                </thead>
                <tbody>
                    {history_rows}
                </tbody>
            </table>
        </div>
        """

        js_labels = "[" + ", ".join(labels_hist) + "]"
        js_data_mat = "[" + ", ".join(data_mat_hist) + "]"
        js_data_lec = "[" + ", ".join(data_lec_hist) + "]"

        contenido_html = ""
        scripts_graficos = f"""
        (function() {{
            var ctx = document.getElementById('historical_line_chart').getContext('2d');
            if (!ctx) return;
            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: {js_labels},
                    datasets: [
                        {{ 
                            label: 'Media de Matemática', 
                            data: {js_data_mat}, 
                            borderColor: '#2980b9', 
                            backgroundColor: '#2980b9', 
                            pointBackgroundColor: '#2980b9',
                            pointBorderColor: '#fff',
                            pointBorderWidth: 2,
                            tension: 0.3, pointRadius: 6, pointHoverRadius: 8 
                        }},
                        {{ 
                            label: 'Media de Lengua', 
                            data: {js_data_lec}, 
                            borderColor: '#c0392b', 
                            backgroundColor: '#c0392b', 
                            pointBackgroundColor: '#c0392b',
                            pointBorderColor: '#fff',
                            pointBorderWidth: 2,
                            tension: 0.3, pointRadius: 6, pointHoverRadius: 8 
                        }}
                    ]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    layout: {{ padding: {{ top: 10 }} }},
                    scales: {{ y: {{ min: 0, max: 100, title: {{ display: true, text: 'Nivel Promedio' }} }} }},
                    plugins: {{
                        datalabels: {{ display: false }}
                    }}
                }}
            }});
        }})();
        """
        
        # BUCLE DE MATERIAS
        for materia in ['Matemática', 'Lengua']:
            df_mat_act = df_esc_act[df_esc_act['Materia'] == materia]
            if df_mat_act.empty: continue
            
            n_act = len(df_mat_act)
            datos_act = (df_mat_act['Categoria'].value_counts(normalize=True) * 100).to_dict() if n_act > 0 else {}
            
            if tiene_anterior:
                df_mat_ant = historial[-2]['df'][(historial[-2]['df']['Centro'] == escuela) & (historial[-2]['df']['Materia'] == materia)]
                n_ant = len(df_mat_ant)
                datos_ant = (df_mat_ant['Categoria'].value_counts(normalize=True) * 100).to_dict() if n_ant > 0 else {}
                label_ant = f"'{historial[-2]['mes']} (N = {n_ant})'"
            else:
                datos_ant = {}
                label_ant = "'Sin datos previos'"
                
            array_data = []
            for n in NIVELES_CONFIG:
                val_ant = datos_ant.get(n['nombre'], 0)
                val_act = datos_act.get(n['nombre'], 0)
                array_data.append(f"{{ label: '{n['nombre']}', data: [{val_ant:.1f}, {val_act:.1f}], backgroundColor: '{n['color']}' }}")
                
            js_datasets = ",\n".join(array_data)
            label_act = f"'{mes_nombre} (N = {n_act})'"
            id_canvas = f"chart_{materia[:3]}"
            materia_lower = materia.lower()
            
            scripts_graficos += f"""
            (function() {{
                var ctx = document.getElementById('{id_canvas}').getContext('2d');
                if (!ctx) return;
                new Chart(ctx, {{
                    type: 'bar',
                    data: {{ labels: [{label_ant}, {label_act}], datasets: [ {js_datasets} ] }},
                    options: {{
                        animation: false, responsive: true, maintainAspectRatio: false, indexAxis: 'y',
                        scales: {{ x: {{ stacked: true, min: 0, max: 100, title: {{ display: true, text: 'Porcentaje de Estudiantes (%)' }}, ticks: {{ callback: function(v) {{ return v + '%'; }}, stepSize: 10 }} }}, y: {{ stacked: true }} }},
                        plugins: {{ tooltip: {{ callbacks: {{ label: function(c) {{ return c.dataset.label + ': ' + c.raw.toFixed(1) + '%'; }} }} }}, datalabels: {{ display: true, color: function(c) {{ return c.dataset.label.includes('Media') ? '#000000' : '#ffffff'; }}, font: {{ weight: 'bold', size: 12 }}, formatter: function(v) {{ return v > 3 ? v.toFixed(1) + '%' : ''; }}, anchor: 'center', align: 'center', offset: 4 }} }}
                    }}
                }});
            }})();
            """
            
            contenido_html += f"""
            <div class="subject-section">
                <div class="subject-title">Resultados de {materia}</div>
                <p class="desc-text">Comparativa mensual de los niveles - {materia}</p>
                
                <details open class="info-accordion no-print">
                    <summary>📖 Ocultar / Mostrar descripción de la gráfica comparativa mensual</summary>
                    <p>Esta gráfica muestra la comparación de la distribución de los niveles de todos los estudiantes que participaron en las pruebas de {materia_lower} de este mes, respecto al mes anterior. Las barras apiladas representan el porcentaje de estudiantes en cada nivel (Crítico: De 0 a 35, Bajo: De más de 35 a 45, Medio: De más de 45 a 55, Bueno: De más de 55 a 65, Excelente: De más de 65 a 100).</p>
                </details>

                <div class="chart-container"><canvas id="{id_canvas}"></canvas></div>
                
                <p class="desc-text">Nivel de las secciones en {mes_nombre}</p>
                
                <details open class="info-accordion no-print">
                    <summary>📖 Ocultar / Mostrar descripción de la tabla de secciones</summary>
                    <p>La etiqueta de la izquierda "Nivel" por cada Grado-Sección, indica el nivel alcanzado por cada grupo de estudiantes. La gráfica de la derecha desglosa específicamente el porcentaje de alumnos que se encuentran en cada nivel durante el mes {mes_nombre}.</p>
                </details>
                
                <div class="group-row" style="margin-bottom: 10px;">
                    <div class="group-data header-dark" style="border-radius: 4px;"><div class="data-col text-left">Grado - Sección</div><div class="data-col text-right">Nivel</div></div>
                    <div class="spacer"></div>
                    <div class="group-faja header-dark" style="justify-content: center; border-radius: 4px; border: none;">Niveles de los estudiantes ({mes_nombre})</div>
                </div>
            """
            
            grados_num = sorted(df_mat_act['Grado_Num'].dropna().unique())
            
            for grado_n in grados_num:
                df_grado_act = df_mat_act[df_mat_act['Grado_Num'] == grado_n]
                grupos = sorted(df_grado_act['Grupo'].dropna().unique())
                
                for grupo in grupos:
                    df_sec = df_grado_act[df_grado_act['Grupo'] == grupo]
                    if df_sec.empty: continue
                    
                    media_sec = df_sec['theta.global (escala 0-100)'].mean()
                    clas_sec = obtener_clasificacion(media_sec) or NIVELES_CONFIG[0]
                    dist = (df_sec['Categoria'].value_counts(normalize=True) * 100).to_dict()
                    faja_html = generar_barra_html(dist)
                    bg_color = clas_sec['bg_color']
                    border_color = clas_sec['color']
                    letra_grupo = str(grupo).split('(')[0].strip() if '(' in str(grupo) else str(grupo).strip()
                    
                    contenido_html += f"""
                    <div class="group-row">
                        <div class="group-data" style="background-color: {bg_color}; border-left: 6px solid {border_color};"><div class="data-col text-left"><strong>{grado_n}.° - {letra_grupo}</strong></div><div class="data-col text-right" style="color: black; font-weight: bold; text-transform: uppercase;">{clas_sec['nombre']}</div></div>
                        <div class="spacer"></div>
                        <div class="group-faja">{faja_html}</div>
                    </div>
                    """
            
            contenido_html += f"""
            </div>
            """

        base_filename = f"{cod_val} - {escuela}"
        safe_filename = "".join(x for x in base_filename if x.isalnum() or x in " ._-") + ".html"

        html_final = HTML_BASE.format(
            centro=escuela, mes_actual=mes_nombre, mes_actual_mayus=mes_nombre.upper(),
            str_mes=str_mes, departamento=dep_val, codigo_infra=cod_val,
            contenido=contenido_html, scripts_graficos=scripts_graficos,
            tabla_historial=history_table
        )
        
        with open(os.path.join(output_dir, safe_filename), 'w', encoding='utf-8') as f:
            f.write(html_final)
            
    print(f"  [OK] ¡Reportes interactivos HTML generados en: {output_dir}")