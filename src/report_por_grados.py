# src/report_por_grados.py
import os
import pandas as pd
import re
import sys
import glob

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def clasificar_estricto(pct):
    """Reglas 0-20, 21-40, 41-60, 61-80, 81-100."""
    if pd.isna(pct): return "Crítico", "#fef2f2", "#991b1b"
    if pct <= 20: return "Crítico", "#fef2f2", "#991b1b"
    elif pct <= 40: return "Bajo", "#fff7ed", "#ff8c2e"
    elif pct <= 60: return "Medio", "#fefce8", "#facc15"
    elif pct <= 80: return "Bueno", "#f7fee7", "#84cc16"
    else: return "Excelente", "#ecfdf5", "#065f46"

def extraer_numero_grado(grado_str):
    if pd.isna(grado_str): return 0
    match = re.search(r'\d+', str(grado_str))
    return int(match.group()) if match else 0

def cargar_datos_directo_resultados():
    ruta = os.path.join(config.PATH_INTERIM, "Resultados")
    archivos = glob.glob(os.path.join(ruta, "*.csv"))
    dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            df['Area temática'] = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
            dfs.append(df)
        except Exception as e:
            print(f" [!] Error leyendo {f}: {e}")
            pass
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def cargar_metadata_indicadores():
    item_dict = {}
    archivos_meta = glob.glob(os.path.join(config.PATH_METADATA, "*.xlsx")) + glob.glob(os.path.join(config.PATH_METADATA, "*.csv"))
    match_mes = re.search(r'^0?(\d+)_', config.MONTH_FOLDER)
    mes_num = match_mes.group(1) if match_mes else "1"
    
    archivo_maestro = None
    for f in archivos_meta:
        nombre = os.path.basename(f).lower()
        if config.EXAM_TYPE.lower() in nombre and (f"mes{mes_num}" in nombre or f"_{mes_num}_" in nombre or f"mes {mes_num}" in nombre):
            if 'procesado' in nombre or 'maestro' in nombre:
                archivo_maestro = f
                break
    if not archivo_maestro and archivos_meta: archivo_maestro = archivos_meta[0]
    if not archivo_maestro: return item_dict

    try:
        if archivo_maestro.endswith('.csv'): df_meta = pd.read_csv(archivo_maestro, encoding='utf-8-sig', sep=None, engine='python')
        else: df_meta = pd.read_excel(archivo_maestro)
            
        col_item = next((c for c in df_meta.columns if 'item' in c.lower() or 'código' in c.lower()), None)
        col_ind = next((c for c in df_meta.columns if 'indicador' in c.lower()), None)
        if col_item and col_ind:
            for _, row in df_meta.dropna(subset=[col_item, col_ind]).iterrows():
                codigo = re.sub(r'[^A-Z0-9]', '', str(row[col_item]).upper())
                item_dict[codigo] = str(row[col_ind]).strip()
    except: pass
    return item_dict

def generar_reporte_por_grados(df_master=None):
    print("\n[*] Generando Reporte Nacional por Grados (Dashboard)...")
    print("    (Calculando huellas dactilares y porcentajes...)")
    df = cargar_datos_directo_resultados()
    if df.empty:
        print("  [!] No hay datos en la carpeta Resultados.")
        return
    
    # Limpieza
    if 'anular_prueba' in df.columns:
        df = df[df['anular_prueba'].isna() | (df['anular_prueba'].astype(str).str.strip() == '')]
    
    col_doc = next((c for c in df.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), 'Documento')
    col_grado = next((c for c in df.columns if 'grado' in str(c).lower()), 'Grado')
    df['Grado_Num'] = df[col_grado].apply(extraer_numero_grado)
    
    # Título y Fecha
    match_mes = re.search(r'^0?(\d+)_', config.MONTH_FOLDER)
    mes_n = match_mes.group(1) if match_mes else "X"
    titulo = f"Resultados de la Prueba de {config.EXAM_TYPE.capitalize()} Mes {mes_n}"
    
    col_fecha = next((c for c in df.columns if 'fecha' in str(c).lower()), None)
    fecha_texto = "No especificada"
    if col_fecha:
        fechas_validas = pd.to_datetime(df[col_fecha].replace('Sin Especificar', pd.NA), errors='coerce', dayfirst=True)
        if not fechas_validas.dropna().empty:
            meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
            mes_moda = fechas_validas.dropna().dt.month.mode().iloc[0]
            anio_moda = fechas_validas.dropna().dt.year.mode().iloc[0]
            fecha_texto = f"{meses_es.get(int(mes_moda), '')} {int(anio_moda)}"
    
    mapa_indicadores = cargar_metadata_indicadores()

    # --- PROCESAMIENTO POR MATERIA ---
    html_sections = {}
    for materia in ['Matemática', 'Lengua']:
        df_m = df[df['Area temática'] == materia].copy()
        if df_m.empty: continue
        
        prefijo = 'MAT' if materia == 'Matemática' else 'LEC'
        
        # Tabla resumen
        conteo = df_m.groupby('Grado_Num')[col_doc].nunique().reset_index()
        tabla_html = f"<div class='card'><h2>👥 Resumen de Estudiantes - {materia}</h2><table><thead><tr><th>Grado</th><th>Estudiantes Evaluados</th></tr></thead><tbody>"
        for _, r in conteo.sort_values('Grado_Num').iterrows(): 
            if r['Grado_Num'] > 0:
                tabla_html += f"<tr><td>{int(r['Grado_Num'])}° Grado</td><td>{int(r[col_doc])}</td></tr>"
        tabla_html += "</tbody></table></div>"
        
        # Indicadores colapsables ¡AQUÍ ESTÁ EL CÁLCULO REAL!
        indicadores_html = ""
        grados = sorted([g for g in df_m['Grado_Num'].unique() if g > 0])
        for g in grados:
            df_g = df_m[df_m['Grado_Num'] == g].copy()
            item_cols = [c for c in df_g.columns if re.match(rf'^{prefijo}\d+$', str(c).upper())]
            
            # 1. Huella dactilar y agrupación
            indicator_groups = {}
            for item in item_cols:
                codigo_puro = re.sub(r'[^A-Z0-9]', '', str(item).upper())
                texto_ind = mapa_indicadores.get(codigo_puro, f"Indicador no disponible ({item})")
                fingerprint = re.sub(r'[\W_]+', '', texto_ind.lower())
                if fingerprint not in indicator_groups:
                    indicator_groups[fingerprint] = {'display_text': texto_ind, 'items': []}
                indicator_groups[fingerprint]['items'].append(item)
            
            # 2. Calcular % de éxito de cada grupo
            lista_final = []
            for fp, data in indicator_groups.items():
                total_corr, total_pos = 0, 0
                for it in data['items']:
                    df_g[it] = pd.to_numeric(df_g[it], errors='coerce')
                    validos = df_g[it].notna()
                    if validos.sum() > 0:
                        total_corr += (df_g.loc[validos, it] == 1).sum()
                        total_pos += validos.sum()
                if total_pos > 0:
                    pct = (total_corr / total_pos) * 100
                    cat, bg, border = clasificar_estricto(pct)
                    lista_final.append({'texto': data['display_text'], 'pct': pct, 'cat': cat, 'bg': bg, 'border': border})
            
            # 3. Ordenar y renderizar ítems HTML
            lista_final.sort(key=lambda x: x['pct'])
            html_items = ""
            for ind in lista_final:
                html_items += f"""
                <div class="indicator-item" style="background-color: {ind['bg']}; border-left: 6px solid {ind['border']};">
                    <div class="pct-box"><strong>{ind['pct']:.1f}%</strong><br><small>{ind['cat']}</small></div>
                    <div class="ind-text">{ind['texto']}</div>
                </div>"""
            
            indicadores_html += f"""
            <details class="grade-accordion">
                <summary>Grado {g}°: Mostrar/Ocultar Niveles de los Indicadores</summary>
                <div class="content">
                    {html_items if html_items else "<p>No hay indicadores evaluados en este grado.</p>"}
                </div>
            </details>"""
            
        html_sections[materia] = tabla_html + indicadores_html

    # --- PLANTILLA HTML CON TABS ---
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
            .indicator-item {{ display: flex; align-items: center; padding: 12px; margin-bottom: 10px; border-radius: 4px; border: 1px solid #e1ecef; }}
            .pct-box {{ width: 80px; text-align: center; border-right: 1px solid rgba(0,0,0,0.1); padding-right: 15px; margin-right: 15px; flex-shrink: 0; }}
            .pct-box strong {{ font-size: 16px; color: #2c3e50; }}
            .pct-box small {{ text-transform: uppercase; font-size: 10px; font-weight: bold; color: #555; }}
            .ind-text {{ font-size: 14px; color: #444; font-weight: 500; }}
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
    
    out_dir = os.path.join(config.PATH_REPORTS, "00_Reportes_Nacionales_Agregados")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"Reporte_Nacional_Grados_{config.MONTH_FOLDER}.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_final)
    print(f"  [OK] Reporte generado exitosamente en:\n       {out_path}")