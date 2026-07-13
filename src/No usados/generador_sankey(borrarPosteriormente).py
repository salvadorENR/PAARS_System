# src/generador_sankey.py
import os
import glob
import pandas as pd
import plotly.graph_objects as go
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# ==========================================
# DICCIONARIOS Y COLORES (Orden: Arriba -> Abajo)
# ==========================================
NIVELES = ['Excelente', 'Bueno', 'Medio', 'Bajo', 'Crítico']
COLORES_NODO = ['#065f46', '#84cc16', '#facc15', '#ff8c2e', '#991b1b']
COLORES_LINK = [
    'rgba(6, 95, 70, 0.4)',     # Excelente
    'rgba(132, 204, 22, 0.4)',  # Bueno
    'rgba(250, 204, 21, 0.4)',  # Medio
    'rgba(255, 140, 46, 0.4)',  # Bajo
    'rgba(153, 27, 27, 0.4)'    # Crítico
]

def clasificar_escuela(promedio):
    if pd.isna(promedio): return None
    if promedio <= 35: return 'Crítico'
    elif promedio <= 45: return 'Bajo'
    elif promedio <= 55: return 'Medio'
    elif promedio <= 65: return 'Bueno'
    else: return 'Excelente'

def cargar_promedios_mes(ruta_base):
    ruta_resultados = os.path.join(ruta_base, "Resultados")
    archivos = glob.glob(os.path.join(ruta_resultados, "*.csv"))
    
    lista_mat, lista_lec = [], []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            materia = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_centro = next((c for c in df.columns if 'nro de centro' in c.lower() or 'código' in c.lower() or 'codigo' in c.lower()), None)
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            
            if not col_theta or not col_centro: continue
            if col_anular:
                df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')]
                
            df['Código'] = df[col_centro].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df = df[df['Código'] != '99999']
            
            df['Puntaje'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=['Puntaje'])
            
            promedios = df.groupby('Código')['Puntaje'].mean().reset_index()
            if materia == 'Matemática': lista_mat.append(promedios)
            else: lista_lec.append(promedios)
        except: pass
        
    df_mat = pd.concat(lista_mat).groupby('Código')['Puntaje'].mean().reset_index() if lista_mat else pd.DataFrame()
    df_lec = pd.concat(lista_lec).groupby('Código')['Puntaje'].mean().reset_index() if lista_lec else pd.DataFrame()
    
    if not df_mat.empty: df_mat['Nivel'] = df_mat['Puntaje'].apply(clasificar_escuela)
    if not df_lec.empty: df_lec['Nivel'] = df_lec['Puntaje'].apply(clasificar_escuela)
    return df_mat, df_lec

def calc_y_centers(df_mes):
    """Fórmula adaptada para apilamiento perfecto proporcional."""
    totals = [len(df_mes[df_mes['Nivel'] == n]) for n in NIVELES]
    pad = 0.04
    grand = sum(totals) if sum(totals) > 0 else 1
    usable = 1.0 - pad * (len(NIVELES) - 1)
    
    y_centers = []
    cursor = 0.0
    for val in totals:
        h = max((val / grand) * usable, 0.005)
        y_centers.append(round(cursor + h / 2, 4))
        cursor += h + pad
    return y_centers

def dibujar_sankey_multimes(lista_meses_datos, materia, nombres_meses, out_dir):
    print(f"  -> Dibujando flujo Sankey Evolutivo para {materia}...")
    
    meses_validos, nombres_validos = [], []
    for i, df in enumerate(lista_meses_datos):
        if not df.empty:
            meses_validos.append(df)
            nombres_validos.append(nombres_meses[i])
            
    if len(meses_validos) < 2:
        print(f"     [!] No hay suficientes meses históricos con datos para {materia}.")
        return
        
    labels_nodos, colores_nodos_lista, node_x, node_y = [], [], [], []
    sources, targets, values, link_colors = [], [], [], []
    
    # --- CÁLCULO DE ESTADÍSTICAS PARA EL TÍTULO (Primer mes vs Último mes) ---
    df_inicial = meses_validos[0]
    df_final = meses_validos[-1]
    df_stats = pd.merge(df_inicial[['Código', 'Nivel']], df_final[['Código', 'Nivel']], on='Código', suffixes=('_Ini', '_Fin'), how='inner')
    
    n_centros = len(df_stats)
    mantienen, suben, bajan = 0, 0, 0
    
    if n_centros > 0:
        for _, row in df_stats.iterrows():
            if row['Nivel_Ini'] not in NIVELES or row['Nivel_Fin'] not in NIVELES: continue
            idx_i = NIVELES.index(row['Nivel_Ini'])
            idx_f = NIVELES.index(row['Nivel_Fin'])
            
            if idx_i == idx_f: mantienen += 1
            elif idx_f < idx_i: suben += 1 # Menor índice = Mejor nivel (0 es Excelente)
            else: bajan += 1
                
        pct_man = (mantienen / n_centros) * 100
        pct_sub = (suben / n_centros) * 100
        pct_baj = (bajan / n_centros) * 100
        # Usamos el símbolo literal '·' en lugar de '&middot;'
        subtitulo_html = f"<span style='font-size:13px;color:#888'>{n_centros} centros · Mantienen: {mantienen} ({pct_man:.1f}%) · Suben: {suben} ({pct_sub:.1f}%) · Bajan: {bajan} ({pct_baj:.1f}%)</span>"
    else:
        subtitulo_html = "<span style='font-size:13px;color:#888'>0 centros comparables</span>"
    
    # 1. Nodos anclados mediante cálculo proporcional
    for i, (nombre, df_mes) in enumerate(zip(nombres_validos, meses_validos)):
        x_val = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        centros_y = calc_y_centers(df_mes)
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])
            
    # 2. Inyectar Flujos "Fantasmas"
    for i in range(len(meses_validos) - 1):
        offset_origen = i * 5
        offset_destino = (i + 1) * 5
        for j in range(5):
            sources.append(offset_origen + j)
            targets.append(offset_destino + j)
            values.append(0.01) 
            link_colors.append('rgba(0,0,0,0)') 
            
    # 3. Flujos Reales
    for i in range(len(meses_validos) - 1):
        df_actual = meses_validos[i]
        df_siguiente = meses_validos[i+1]
        
        df_cruce = pd.merge(df_actual[['Código', 'Nivel']], df_siguiente[['Código', 'Nivel']], 
                            on='Código', suffixes=('_A', '_S'), how='inner')
        if df_cruce.empty: continue
        
        flujos = df_cruce.groupby(['Nivel_A', 'Nivel_S']).size().reset_index(name='Cantidad')
        
        offset_origen = i * 5
        offset_destino = (i + 1) * 5
        mapa_origen = {n: j + offset_origen for j, n in enumerate(NIVELES)}
        mapa_destino = {n: j + offset_destino for j, n in enumerate(NIVELES)}
        
        for _, row in flujos.iterrows():
            if row['Cantidad'] == 0: continue
            src_idx = mapa_origen[row['Nivel_A']]
            tgt_idx = mapa_destino[row['Nivel_S']]
            
            sources.append(src_idx)
            targets.append(tgt_idx)
            values.append(row['Cantidad'])
            link_colors.append(COLORES_LINK[src_idx % 5])
            
    if len(sources) <= len(meses_validos) * 5: 
        return
        
    fig = go.Figure(data=[go.Sankey(
        arrangement = "fixed", 
        node = dict(pad=20, thickness=24, line=dict(color="white", width=0.5), 
                    label=labels_nodos, color=colores_nodos_lista,
                    x=node_x, y=node_y), 
        link = dict(source=sources, target=targets, value=values, color=link_colors)
    )])
    
    ancho_dinamico = max(900, 800 + (len(meses_validos) * 150))
    fig.update_layout(
        title=dict(
            text=f"<span style='font-size:22px; color:#1a2b4c; font-family:Arial, sans-serif; font-weight:bold;'>Trayectorias de niveles — {materia}</span><br>{subtitulo_html}",
            x=0.01,
            y=0.95
        ),
        font_size=12, width=ancho_dinamico, height=600,
        margin=dict(l=20, r=20, t=100, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    
    os.makedirs(out_dir, exist_ok=True)
    
    # 1. Guardar interactivo (HTML)
    html_path = os.path.join(out_dir, f"Sankey_Evolutivo_{materia}.html")
    fig.write_html(html_path)
    
    # 2. Guardar estático (PNG)
    png_path = os.path.join(out_dir, f"Sankey_Evolutivo_{materia}.png")
    try: 
        fig.write_image(png_path, scale=2)
        print(f"     [OK] PNG e HTML generados para {materia}.")
    except Exception as e: 
        print(f"     [!] Advertencia: No se pudo generar el PNG para {materia}. Asegúrate de tener instalado 'kaleido' (pip install kaleido). Error: {e}")

def generar_sankeys_reporte_corto():
    print("\n[*] Iniciando Generador de Flujos Sankey Históricos...")
    carpetas_progreso = sorted([
        d for d in os.listdir(config.YEAR_DIR) 
        if os.path.isdir(os.path.join(config.YEAR_DIR, d)) and 'PROGRESO' in d.upper()
    ])
    
    carpetas_hasta_actual = []
    for c in carpetas_progreso:
        carpetas_hasta_actual.append(c)
        if c == config.MONTH_FOLDER: break
            
    if len(carpetas_hasta_actual) < 2:
        print("  [!] Solo hay un mes disponible en la línea de tiempo. No se puede generar evolución.")
        return
        
    nombres_meses = [c.split('_')[-1].capitalize() for c in carpetas_hasta_actual]
    out_dir = os.path.join(config.PATH_REPORTS, "Reporte_Corto_LaTeX", "Graficos_Sankey")
    
    lista_mat, lista_lec = [], []
    for c in carpetas_hasta_actual:
        ruta_interim = os.path.join(config.YEAR_DIR, c, "Interim_CSVs")
        print(f"  -> Extrayendo promedios de: {c}")
        d_mat, d_lec = cargar_promedios_mes(ruta_interim)
        lista_mat.append(d_mat)
        lista_lec.append(d_lec)
        
    dibujar_sankey_multimes(lista_mat, "Matemática", nombres_meses, out_dir)
    dibujar_sankey_multimes(lista_lec, "Lengua", nombres_meses, out_dir)
    print(f"  [OK] Gráficos de flujo guardados en:\n       {out_dir}")

if __name__ == "__main__":
    generar_sankeys_reporte_corto()