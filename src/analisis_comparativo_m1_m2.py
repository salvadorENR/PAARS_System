# src/analisis_comparativo_m1_m2.py
import os
import sys
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go

# Conexión con el mapa PAARS
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# ==========================================
# CONFIGURACIÓN DINÁMICA AUTOSUFICIENTE
# ==========================================
PATH_MES_ACTUAL = config.PATH_INTERIM
PATH_OUTPUT_PLOTS = os.path.join(config.PATH_REPORTS, "Tablas_Graficas_ReportCorto")

nombre_mes_act = config.MONTH_FOLDER.split('_')[-1].capitalize()
match_act = re.search(r'(\d+)_PROGRESO', config.MONTH_FOLDER, re.IGNORECASE)
num_act = int(match_act.group(1)) if match_act else "Actual"

carpetas_progreso = sorted([
    d for d in os.listdir(config.YEAR_DIR) 
    if os.path.isdir(os.path.join(config.YEAR_DIR, d)) and 'PROGRESO' in d.upper()
])

PATH_MES_ANTERIOR = None
nombre_mes_ant = "Mes_Anterior"
num_ant = "Anterior"

if config.MONTH_FOLDER in carpetas_progreso:
    idx = carpetas_progreso.index(config.MONTH_FOLDER)
    if idx > 0:
        carpeta_anterior = carpetas_progreso[idx-1]
        PATH_MES_ANTERIOR = os.path.join(config.YEAR_DIR, carpeta_anterior, "Interim_CSVs")
        nombre_mes_ant = carpeta_anterior.split('_')[-1].capitalize()
        match_ant = re.search(r'(\d+)_PROGRESO', carpeta_anterior, re.IGNORECASE)
        if match_ant: num_ant = int(match_ant.group(1))

LBL_MES_ANT = f"Prueba Anterior ({nombre_mes_ant})"
LBL_MES_ACT = f"Prueba Actual ({nombre_mes_act})"

COL_SCORE = 'theta.global (escala 0-100)'
GRADOS_TODOS = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
GRADOS_BOXPLOT = [3, 6, 9]

MAPEO_NOMBRES_GRADO = {2: "Segundo Grado", 3: "Tercer Grado", 4: "Cuarto Grado", 5: "Quinto Grado", 6: "Sexto Grado", 7: "Séptimo Grado", 8: "Octavo Grado", 9: "Noveno Grado", 10: "Décimo Grado", 11: "Undécimo Grado"}

# Paletas de color globales
COLOR_CATEGORIAS = {"Crítico": "#991b1b", "Bajo": "#ff8c2e", "Medio": "#facc15", "Bueno": "#84cc16", "Excelente": "#065f46"}
ORDEN_CATEGORIAS = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]

# ==========================================
# CONFIGURACIÓN (SANKEY EVOLUTIVO)
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


# ==========================================
# FUNCIONES: ANÁLISIS COMPARATIVO ESTÁNDAR
# ==========================================
def preparar_entorno():
    sns.set_theme(style="whitegrid")
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14
    os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)

def clasificar_puntaje_estudiante(val):
    if pd.isna(val): return "Crítico"
    val = float(val)
    if val <= 35: return "Crítico"
    elif val <= 45: return "Bajo"
    elif val <= 55: return "Medio"
    elif val <= 65: return "Bueno"
    else: return "Excelente"

def extraer_numero_grado(grado_str):
    if pd.isna(grado_str): return None
    match = re.search(r'\d+', str(grado_str).lower())
    return int(match.group()) if match else None

def cargar_y_procesar_mes_estudiantes(path_mes, etiqueta_mes):
    archivos = glob.glob(os.path.join(path_mes, "*.csv"))
    if not archivos: return pd.DataFrame()

    lista_dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            # AQUÍ SE CORRIGIÓ: Lectura -> Lengua
            materia = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            col_grado = next((c for c in df.columns if 'grado' in c.lower()), None)

            if not col_theta or not col_grado: continue

            if col_anular:
                df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')].copy()
                
            df[COL_SCORE] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=[COL_SCORE])
            if df.empty: continue

            df['Grado_Num'] = df[col_grado].apply(extraer_numero_grado)
            df['Nivel_Logro'] = df[COL_SCORE].apply(clasificar_puntaje_estudiante)
            df_final = df[[COL_SCORE, 'Grado_Num', 'Nivel_Logro']].copy()
            df_final['Materia'] = materia
            df_final['Mes'] = etiqueta_mes
            lista_dfs.append(df_final)
        except: pass
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()


# ==========================================
# FUNCIONES: GENERADOR SANKEY EVOLUTIVO
# ==========================================
def clasificar_escuela(promedio):
    if pd.isna(promedio): return None
    if promedio <= 35: return 'Crítico'
    elif promedio <= 45: return 'Bajo'
    elif promedio <= 55: return 'Medio'
    elif promedio <= 65: return 'Bueno'
    else: return 'Excelente'

def cargar_promedios_mes_escuelas(ruta_base):
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
            elif idx_f < idx_i: suben += 1 
            else: bajan += 1
                
        pct_man = (mantienen / n_centros) * 100
        pct_sub = (suben / n_centros) * 100
        pct_baj = (bajan / n_centros) * 100
        subtitulo_html = f"<span style='font-size:13px;color:#888'>{n_centros} centros · Mantienen: {mantienen} ({pct_man:.1f}%) · Suben: {suben} ({pct_sub:.1f}%) · Bajan: {bajan} ({pct_baj:.1f}%)</span>"
    else:
        subtitulo_html = "<span style='font-size:13px;color:#888'>0 centros comparables</span>"
    
    for i, (nombre, df_mes) in enumerate(zip(nombres_validos, meses_validos)):
        x_val = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        centros_y = calc_y_centers(df_mes)
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])
            
    for i in range(len(meses_validos) - 1):
        offset_origen = i * 5
        offset_destino = (i + 1) * 5
        for j in range(5):
            sources.append(offset_origen + j)
            targets.append(offset_destino + j)
            values.append(0.01) 
            link_colors.append('rgba(0,0,0,0)') 
            
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
    html_path = os.path.join(out_dir, f"Sankey_Evolutivo_{materia}.html")
    fig.write_html(html_path)
    
    png_path = os.path.join(out_dir, f"Sankey_Evolutivo_{materia}.png")
    try: 
        fig.write_image(png_path, scale=2)
    except Exception as e: 
        pass

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
        print("  [!] Solo hay un mes disponible en la línea de tiempo. Sankeys omitidos.")
        return
        
    nombres_meses = [c.split('_')[-1].capitalize() for c in carpetas_hasta_actual]
    out_dir = PATH_OUTPUT_PLOTS
    
    lista_mat, lista_lec = [], []
    for c in carpetas_hasta_actual:
        ruta_interim = os.path.join(config.YEAR_DIR, c, "Interim_CSVs")
        print(f"  -> Extrayendo promedios Sankey de: {c}")
        d_mat, d_lec = cargar_promedios_mes_escuelas(ruta_interim)
        lista_mat.append(d_mat)
        lista_lec.append(d_lec)
        
    dibujar_sankey_multimes(lista_mat, "Matemática", nombres_meses, out_dir)
    dibujar_sankey_multimes(lista_lec, "Lengua", nombres_meses, out_dir)


# ==========================================
# EJECUCIÓN PRINCIPAL
# ==========================================
def ejecutar_comparativo():
    print(f"\n[*] Iniciando Motor de Análisis Intermensual.\n    Todo se guardará en: {PATH_OUTPUT_PLOTS}")
    
    preparar_entorno()
    
    # 1. Generar Gráficos Sankey
    generar_sankeys_reporte_corto()

    # 2. Generar Comparativas Estadísticas y Gráficos
    if not PATH_MES_ANTERIOR or not os.path.exists(PATH_MES_ANTERIOR):
        print("\n[!] No se encontró el mes anterior. Omitiendo generación de KDE, Boxplots y Tablas.")
        return

    print("\n[*] Generando KDE, Boxplots y Barras Apiladas Intermensuales...")
    df_m1 = cargar_y_procesar_mes_estudiantes(PATH_MES_ANTERIOR, LBL_MES_ANT)
    df_m2 = cargar_y_procesar_mes_estudiantes(PATH_MES_ACTUAL, LBL_MES_ACT)
    
    if df_m1.empty or df_m2.empty:
        print("[!] Faltan datos en los CSV para graficar.")
    else:
        df_master = pd.concat([df_m1, df_m2], ignore_index=True)
        # Paleta suave para KDE y Boxplots (Celeste y Naranja suave)
        palette_meses = {LBL_MES_ANT: "#80b1d3", LBL_MES_ACT: "#fb8072"}

        for materia in df_master['Materia'].unique():
            print(f"    - Dibujando gráficas para {materia}...")
            df_sub = df_master[df_master['Materia'] == materia].copy()
            
            # ----------------------------------------------------
            # 2.1 TABLAS DE EXCEL
            # ----------------------------------------------------
            writer_path = os.path.join(PATH_OUTPUT_PLOTS, f"Tablas_Comparativas_{materia}.xlsx")
            with pd.ExcelWriter(writer_path) as writer:
                for mes_nombre in df_sub['Mes'].unique():
                    df_mes_stats = df_sub[df_sub['Mes'] == mes_nombre].copy()
                    if df_mes_stats.empty: continue
                    
                    desc = df_mes_stats.groupby('Grado_Num')[COL_SCORE].describe().round(2).reset_index()
                    desc['Asignatura/Grado'] = materia + " (" + desc['Grado_Num'].map(MAPEO_NOMBRES_GRADO) + ")"
                    desc = desc.rename(columns={'mean': 'Media', '50%': 'Mediana', 'std': 'D.S.'})
                    desc.sort_values('Grado_Num')[['Asignatura/Grado', 'Media', 'Mediana', 'D.S.']].to_excel(writer, sheet_name=mes_nombre[:31], index=False)

            # ----------------------------------------------------
            # 2.2 CURVAS KDE (DISTRIBUCIÓN)
            # ----------------------------------------------------
            plt.figure(figsize=(12, 7))
            sns.histplot(data=df_sub, x=COL_SCORE, hue='Mes', kde=True, element="step", stat="probability", common_norm=False, palette=palette_meses, alpha=0.6)
            plt.title(f"Evolución de Puntajes Theta: {materia}\n({nombre_mes_ant} vs {nombre_mes_act})")
            plt.xlim(0, 100)
            plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Distribucion_Evolutiva_{materia}.png"), dpi=300, bbox_inches='tight')
            plt.close()

            # ----------------------------------------------------
            # 2.3 BOXPLOTS (3°, 6°, 9°)
            # ----------------------------------------------------
            df_box = df_sub[df_sub['Grado_Num'].isin(GRADOS_BOXPLOT)].copy()
            if not df_box.empty:
                df_box['Grado_Str'] = df_box['Grado_Num'].astype(str) + "° Grado"
                plt.figure(figsize=(10, 6))
                sns.boxplot(data=df_box, x='Grado_Str', y=COL_SCORE, hue='Mes',
                            palette=palette_meses, showmeans=True,
                            meanprops={"marker":"^", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"8"})
                plt.title(f"Puntajes para 3°, 6° y 9°: {materia}")
                plt.ylim(0, 100)
                plt.ylabel("Puntajes (0-100)")
                plt.xlabel("Grado")
                plt.legend(title="Mes de Aplicación")
                plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Boxplot_Grados_Theta_{materia}.png"), dpi=300, bbox_inches='tight')
                plt.close()

            # ----------------------------------------------------
            # 2.4 BARRAS HORIZONTALES APILADAS (LADO A LADO)
            # ----------------------------------------------------
            df_bars = df_sub.copy()
            df_bars['Grado_Str'] = df_bars['Grado_Num'].astype(str) + "° Grado"
            
            grados_orden = [f"{g}° Grado" for g in sorted(df_bars['Grado_Num'].unique(), reverse=True)]
            
            fig, axes = plt.subplots(1, 2, figsize=(20, 8), sharey=True)
            
            meses_list = [(LBL_MES_ANT, f"Mes {num_ant} ({nombre_mes_ant})"), 
                          (LBL_MES_ACT, f"Mes {num_act} ({nombre_mes_act})")]
            
            for i, (mes_key, mes_title) in enumerate(meses_list):
                ax = axes[i]
                df_mes_bar = df_bars[df_bars['Mes'] == mes_key]
                if df_mes_bar.empty: 
                    ax.set_visible(False)
                    continue
                    
                ct = pd.crosstab(df_mes_bar['Grado_Str'], df_mes_bar['Nivel_Logro'], normalize='index') * 100
                
                for cat in ORDEN_CATEGORIAS:
                    if cat not in ct.columns: ct[cat] = 0
                ct = ct[ORDEN_CATEGORIAS]
                ct = ct.reindex(grados_orden)
                
                ct.plot(kind='barh', stacked=True, ax=ax, color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS], 
                        legend=False, edgecolor='white', width=0.6)
                
                ax.set_title(mes_title, fontsize=24, color='#005288', pad=20, fontweight='bold')
                ax.set_xlim(0, 100)
                ax.set_xlabel("Porcentaje de Estudiantes (%)")
                ax.set_ylabel("Grado" if i == 0 else "")
                
                if i == 1:
                    handles, labels = ax.get_legend_handles_labels()
                    ax.legend(handles, labels, title="Nivel de Logro", bbox_to_anchor=(1.02, 1), loc='upper left')
                
                for p in ax.patches:
                    width = p.get_width()
                    if width > 4: 
                        x = p.get_x() + width / 2
                        y = p.get_y() + p.get_height() / 2
                        ax.text(x, y, f"{width:.1f}%", ha='center', va='center', color='white', fontsize=10, fontweight='bold')
                        
            plt.suptitle(f"Distribución por Niveles: {materia}\n(De 2° a 11°)", fontsize=16)
            
            plt.tight_layout(rect=[0, 0, 0.9, 1]) 
            
            plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Barras_Horizontales_Comparativas_{materia}.png"), dpi=300, bbox_inches='tight')
            plt.close()
            
    print(f"\n[OK] ¡Kit completo de Gráficos y Tablas listos en:\n     {PATH_OUTPUT_PLOTS}")

if __name__ == "__main__":
    ejecutar_comparativo()