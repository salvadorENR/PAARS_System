import os
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS (INDEPENDIENTE)
# ==========================================
DIRECTORIOS_MESES = [
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados", "Mes 1 (Marzo)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados", "Mes 2 (Abril)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados", "Mes 3 (Mayo)")
]

PATH_METADATA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_OUTPUT_PLOTS = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Final_Reports\Tablas_Graficas_ReportCorto"
os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)

# ==========================================
# 2. CONFIGURACIÓN DE COLORES Y VARIABLES
# ==========================================
COL_SCORE_EXACTA = 'theta.global (escala 0-100)'

NIVELES = ['Excelente', 'Bueno', 'Medio', 'Bajo', 'Crítico']
ORDEN_CATEGORIAS = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
COLOR_CATEGORIAS = {"Crítico": "#991b1b", "Bajo": "#ff8c2e", "Medio": "#facc15", "Bueno": "#84cc16", "Excelente": "#065f46"}

COLORES_MESES = ["#80b1d3", "#fb8072", "#b3de69"]
COLORES_NODO = ['#065f46', '#84cc16', '#facc15', '#ff8c2e', '#991b1b']
COLORES_LINK = [
    'rgba(6, 95, 70, 0.4)',     # Excelente
    'rgba(132, 204, 22, 0.4)',  # Bueno
    'rgba(250, 204, 21, 0.4)',  # Medio
    'rgba(255, 140, 46, 0.4)',  # Bajo
    'rgba(153, 27, 27, 0.4)'    # Crítico
]

# ==========================================
# 3. FUNCIONES DE LECTURA Y PROCESAMIENTO
# ==========================================
def preparar_entorno():
    sns.set_theme(style="whitegrid")
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14

def clasificar_puntaje(val):
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

def cargar_matricula():
    print(f"\n[*] Cargando archivo de Matrícula desde: {PATH_METADATA}")
    if not os.path.exists(PATH_METADATA):
        print("  [!] ERROR: No se encontró el archivo de matrícula.")
        return pd.DataFrame()
        
    try:
        df_mat = pd.read_csv(PATH_METADATA, dtype=str, encoding_errors='ignore')
        df_mat.columns = df_mat.columns.str.strip().str.upper()
        
        if 'NIE' not in df_mat.columns:
            print("  [!] ERROR: No se encontró la columna 'NIE' en la matrícula.")
            return pd.DataFrame()
            
        df_mat['NIE_Limpio'] = df_mat['NIE'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        if 'GRUPO' in df_mat.columns:
            df_mat['GRUPO_Limpio'] = df_mat['GRUPO'].astype(str).str.strip().str.upper()
        else:
            df_mat['GRUPO_Limpio'] = 'DESCONOCIDO'
            
        return df_mat
    except Exception as e:
        print(f"  [!] Error leyendo matrícula: {e}")
        return pd.DataFrame()

def extraer_csvs_puntajes():
    lista_dfs = []
    print("[*] Extrayendo y combinando datos de puntajes de los meses indicados...")
    
    for ruta, etiqueta_mes in DIRECTORIOS_MESES:
        if not os.path.exists(ruta): continue
        archivos = glob.glob(os.path.join(ruta, "*.csv"))
        
        for f in archivos:
            if 'legend' in os.path.basename(f).lower(): continue
            try:
                df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
                df.columns = df.columns.str.strip()
                
                materia = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
                
                if COL_SCORE_EXACTA not in df.columns: continue
                col_theta = COL_SCORE_EXACTA
                
                col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
                col_grado = next((c for c in df.columns if 'grado' in c.lower()), None)
                col_codigo = next((c for c in df.columns if 'nro de centro' in c.lower() or 'código' in c.lower() or 'codigo' in c.lower() or ('centro' in c.lower() and 'id' in c.lower())), None)
                col_centro = next((c for c in df.columns if c.lower() == 'centro' or c.lower() == 'nombre centro' or c.lower() == 'institución'), None)
                if not col_centro:
                    col_centro = next((c for c in df.columns if 'centro' in c.lower() and 'nro' not in c.lower() and 'código' not in c.lower()), None)
                col_grupo = next((c for c in df.columns if 'grupo' in c.lower() or 'sección' in c.lower() or 'seccion' in c.lower()), None)
                col_doc = next((c for c in df.columns if 'documento' in c.lower() or 'nie' in c.lower() or 'id' == c.lower()), None)

                if not (col_grado and col_codigo and col_doc): continue

                if col_anular:
                    df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '') | (df[col_anular].astype(str).str.lower() == 'nan')].copy()
                    
                df[col_theta] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
                df = df.dropna(subset=[col_theta])
                if df.empty: continue

                df['Grado_Num'] = df[col_grado].apply(extraer_numero_grado)
                df['Grado_Str'] = df[col_grado].astype(str).str.strip()
                df['Nro de Centro'] = df[col_codigo].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df['Centro'] = df[col_centro].astype(str).str.strip() if col_centro else "Desconocido"
                df['Grupo'] = df[col_grupo].astype(str).str.strip() if col_grupo else "Desconocido"
                df['Documento'] = df[col_doc].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df['Nivel_Logro'] = df[col_theta].apply(clasificar_puntaje)
                
                df_final = df[[col_theta, 'Grado_Num', 'Grado_Str', 'Nivel_Logro', 'Nro de Centro', 'Centro', 'Grupo', 'Documento']].copy()
                df_final['Materia'] = materia
                df_final['Mes'] = etiqueta_mes
                
                df_final = df_final[df_final['Nro de Centro'] != '99999']
                lista_dfs.append(df_final)
            except: pass
            
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()

# ==========================================
# 4. EXPORTAR Y GRAFICAR B2 (NUEVO)
# ==========================================
def generar_excel_y_barras_b2(df_b2_mes3, df_matricula):
    if df_b2_mes3.empty:
        print("\n  [!] No hay estudiantes en el Grupo B2 para el Mes 3.")
        return
        
    print("\n[*] Generando Excel y Gráficas de Barras exclusivas para el Grupo B2...")
    
    # --- A. EXCEL DEL GRUPO B2 ---
    df_pivot = df_b2_mes3.pivot_table(
        index='Documento',
        columns='Materia',
        values=COL_SCORE_EXACTA,
        aggfunc='first'
    ).reset_index()
    
    if 'Matemática' not in df_pivot.columns: df_pivot['Matemática'] = pd.NA
    if 'Lengua' not in df_pivot.columns: df_pivot['Lengua'] = pd.NA
    
    df_pivot = df_pivot.rename(columns={'Matemática': 'Puntaje Matemática', 'Lengua': 'Puntaje Lengua'})
    
    cols_solicitadas = ['CODIGO', 'NOMBRE', 'CODIGO_SECCIÓN', 'GRADO', 'NIE', 'PRIMER NOMBRE', 'SEGUNDO NOMBRE', 'PRIMER APELLIDO', 'SEGUNDO APELLIDO']
    cols_existentes = [c for c in cols_solicitadas if c in df_matricula.columns]
    
    df_mat_export = df_matricula[cols_existentes + ['NIE_Limpio']].copy()
    
    df_excel_b2 = pd.merge(df_mat_export, df_pivot, left_on='NIE_Limpio', right_on='Documento', how='inner')
    df_excel_b2 = df_excel_b2.drop(columns=['NIE_Limpio', 'Documento'], errors='ignore')
    
    # Ordenar columnas al formato final solicitado
    final_cols = [c for c in cols_solicitadas if c in df_excel_b2.columns] + ['Puntaje Matemática', 'Puntaje Lengua']
    df_excel_b2 = df_excel_b2[final_cols]
    
    out_path = os.path.join(PATH_OUTPUT_PLOTS, "Resultados_Mes3_Grupo_B2.xlsx")
    df_excel_b2.to_excel(out_path, index=False)
    print(f"  [OK] Excel del Grupo B2 guardado en: {out_path}")
    
    # --- B. GRÁFICAS DE BARRAS B2 ---
    for materia in df_b2_mes3['Materia'].unique():
        df_sub = df_b2_mes3[df_b2_mes3['Materia'] == materia].copy()
        if df_sub.empty: continue
        
        df_sub['Grado_Str'] = df_sub['Grado_Num'].astype(str) + "° Grado"
        grados_orden = [f"{g}° Grado" for g in sorted(df_sub['Grado_Num'].unique(), reverse=True)]
        
        fig, ax = plt.subplots(figsize=(10, 8))
        
        ct = pd.crosstab(df_sub['Grado_Str'], df_sub['Nivel_Logro'], normalize='index') * 100
        for cat in ORDEN_CATEGORIAS:
            if cat not in ct.columns: ct[cat] = 0
        ct = ct[ORDEN_CATEGORIAS]
        ct = ct.reindex(grados_orden)
        
        ct.plot(kind='barh', stacked=True, ax=ax, color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS], 
                legend=False, edgecolor='white', width=0.6)
        
        ax.set_title(f"Mes 3 (Mayo) - Alumnos B2 - {materia}", fontsize=20, color='#005288', pad=20, fontweight='bold')
        ax.set_xlim(0, 100)
        ax.set_xlabel("Porcentaje de Estudiantes (%)")
        ax.set_ylabel("Grado")
        
        handles, labels = ax.get_legend_handles_labels()
        ax.legend(handles, labels, title="Nivel de Logro", bbox_to_anchor=(1.02, 1), loc='upper left')
        
        for p in ax.patches:
            width = p.get_width()
            if width > 4: 
                x = p.get_x() + width / 2
                y = p.get_y() + p.get_height() / 2
                ax.text(x, y, f"{width:.1f}%", ha='center', va='center', color='white', fontsize=10, fontweight='bold')
                
        plt.suptitle(f"Niveles de Logro Grupo B2 (Mes 3): {materia}\n(De 2° a 11°)", fontsize=16)
        plt.tight_layout(rect=[0, 0, 0.9, 1]) 
        
        ruta_guardado = os.path.join(PATH_OUTPUT_PLOTS, f"Barras_Horizontales_GrupoB2_{materia}.png")
        plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight')
        plt.close()

# ==========================================
# 5. EXPORTAR B1 CONSOLIDADO
# ==========================================
def exportar_base_consolidada_b1(df_master):
    print("\n[*] Generando Base de Datos Consolidada (Grupo B1)...")
    try:
        df_pivot = df_master.pivot_table(
            index=['Nro de Centro', 'Centro', 'Grupo', 'Documento', 'Grado_Str'],
            columns=['Materia', 'Mes'],
            values=COL_SCORE_EXACTA,
            aggfunc='first'
        ).reset_index()
        
        df_pivot = df_pivot.rename(columns={'Grado_Str': 'Grado'})
        
        nuevas_columnas = []
        for col in df_pivot.columns:
            if isinstance(col, tuple):
                materia, mes = col[0], col[1]
                nuevas_columnas.append(f"{COL_SCORE_EXACTA} (para {materia.lower()} {mes.lower()})")
            else:
                nuevas_columnas.append(col)
                
        df_pivot.columns = nuevas_columnas
        
        out_path = os.path.join(PATH_OUTPUT_PLOTS, "Base_Estudiantes_B1_Consolidada.xlsx")
        df_pivot.to_excel(out_path, index=False)
        print(f"  [OK] Excel consolidado B1 guardado en: {out_path}")
        
    except Exception as e:
        print(f"  [!] Error al exportar la base B1: {e}")

# ==========================================
# 6. SANKEY MULTIMES CON FIJACIÓN ESTRICTA
# ==========================================
def calc_y_centers(df_mes):
    totals = [len(df_mes[df_mes['Nivel'] == n]) for n in NIVELES]
    pad_frac = 0.04
    usable = 1.0 - (pad_frac * (len(NIVELES) - 1))
    
    grand = sum(totals) if sum(totals) > 0 else 1
    
    min_frac = 0.01 
    heights = [max((val / grand) * usable, min_frac) for val in totals]
    
    total_h = sum(heights)
    if total_h > usable:
        factor = usable / total_h
        heights = [h * factor for h in heights]
        
    y_centers = []
    cursor = 0.0
    for h in heights:
        y_centers.append(round(cursor + h / 2, 4))
        cursor += h + pad_frac
        
    return y_centers

def dibujar_sankey_multimes(df_master, materia, meses_labels):
    print(f"  -> Dibujando flujo Sankey para {materia} (Grupo B1)...")
    
    df_mat = df_master[df_master['Materia'] == materia].copy()
    if df_mat.empty: return
    
    meses_validos, nombres_validos = [], []
    for mes in meses_labels:
        df_m = df_mat[df_mat['Mes'] == mes]
        if not df_m.empty:
            promedios = df_m.groupby('Nro de Centro')[COL_SCORE_EXACTA].mean().reset_index()
            promedios['Nivel'] = promedios[COL_SCORE_EXACTA].apply(clasificar_puntaje)
            promedios = promedios.rename(columns={'Nro de Centro': 'Código'})
            meses_validos.append(promedios)
            nombres_validos.append(mes)
            
    if len(meses_validos) < 2: return
        
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
        subtitulo_html = f"<span style='font-size:13px;color:#888'>Población base (Grupo B1): {n_centros} centros <br>Mantienen: {mantienen} ({pct_man:.1f}%) · Suben: {suben} ({pct_sub:.1f}%) · Bajan: {bajan} ({pct_baj:.1f}%)</span>"
    else:
        subtitulo_html = ""
    
    for i, (nombre, df_mes) in enumerate(zip(nombres_validos, meses_validos)):
        x_val = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        centros_y = calc_y_centers(df_mes)
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])
            
    for i in range(len(meses_validos) - 1):
        offset_origen = i * len(NIVELES)
        offset_destino = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(offset_origen + j)
            targets.append(offset_destino + j)
            values.append(1e-9) 
            link_colors.append('rgba(0,0,0,0)') 
            
    for i in range(len(meses_validos) - 1):
        df_actual = meses_validos[i]
        df_siguiente = meses_validos[i+1]
        
        df_cruce = pd.merge(df_actual[['Código', 'Nivel']], df_siguiente[['Código', 'Nivel']], 
                            on='Código', suffixes=('_A', '_S'), how='inner')
        if df_cruce.empty: continue
        
        flujos = df_cruce.groupby(['Nivel_A', 'Nivel_S']).size().reset_index(name='Cantidad')
        
        offset_origen = i * len(NIVELES)
        offset_destino = (i + 1) * len(NIVELES)
        mapa_origen = {n: j + offset_origen for j, n in enumerate(NIVELES)}
        mapa_destino = {n: j + offset_destino for j, n in enumerate(NIVELES)}
        
        for _, row in flujos.iterrows():
            if row['Cantidad'] == 0: continue
            src_idx = mapa_origen[row['Nivel_A']]
            tgt_idx = mapa_destino[row['Nivel_S']]
            
            sources.append(src_idx)
            targets.append(tgt_idx)
            values.append(row['Cantidad'])
            link_colors.append(COLORES_LINK[src_idx % len(NIVELES)])
            
    if not sources: return
        
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
    
    html_path = os.path.join(PATH_OUTPUT_PLOTS, f"Sankey_Evolutivo_{materia}.html")
    fig.write_html(html_path)
    
    try: 
        png_path = os.path.join(PATH_OUTPUT_PLOTS, f"Sankey_Evolutivo_{materia}.png")
        fig.write_image(png_path, scale=2)
    except: pass

# ==========================================
# 7. GRÁFICOS ESTADÍSTICOS GRUPO B1
# ==========================================
def generar_graficos_estadisticos_b1(df_master, meses_labels):
    print(f"\n[*] Generando KDE, Boxplots y Barras Apiladas (Grupo B1)...")
    palette_meses = {mes: COLORES_MESES[i % len(COLORES_MESES)] for i, mes in enumerate(meses_labels)}

    for materia in df_master['Materia'].unique():
        print(f"    - Dibujando gráficas principales para {materia}...")
        df_sub = df_master[df_master['Materia'] == materia].copy()
        
        # A. CURVAS KDE
        plt.figure(figsize=(12, 7))
        sns.histplot(data=df_sub, x=COL_SCORE_EXACTA, hue='Mes', kde=True, element="step", stat="probability", common_norm=False, palette=palette_meses, alpha=0.6)
        titulo_meses = " vs ".join([m.split(' ')[-1].replace('(','').replace(')','') for m in meses_labels])
        plt.title(f"Evolución de Puntajes Theta: {materia}\n({titulo_meses} - Grupo B1)")
        plt.xlim(0, 100)
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Distribucion_Evolutiva_B1_{materia}.png"), dpi=300, bbox_inches='tight')
        plt.close()

        # B. BOXPLOT EVOLUTIVO (HORIZONTAL)
        plt.figure(figsize=(10, 6))
        sns.boxplot(data=df_sub, x=COL_SCORE_EXACTA, y='Mes',
                    palette=palette_meses, showmeans=True,
                    meanprops={"marker":"<", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"10"})
        plt.title(f"Distribución Global de Puntajes: {materia}\n(Grupo B1 - Alumnos constantes de M1 o M2)")
        plt.xlim(0, 100)
        plt.xlabel("Puntajes (0-100)")
        plt.ylabel("Prueba de Progreso")
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Boxplot_Evolutivo_Global_B1_{materia}.png"), dpi=300, bbox_inches='tight')
        plt.close()

        # C. BARRAS HORIZONTALES APILADAS
        df_bars = df_sub.copy()
        df_bars['Grado_Str'] = df_bars['Grado_Num'].astype(str) + "° Grado"
        grados_orden = [f"{g}° Grado" for g in sorted(df_bars['Grado_Num'].unique(), reverse=True)]
        
        n_meses = len(meses_labels)
        ancho_fig = max(10, n_meses * 10)
        fig, axes = plt.subplots(1, n_meses, figsize=(ancho_fig, 8), sharey=True)
        if n_meses == 1: axes = [axes]
        
        for i, mes_key in enumerate(meses_labels):
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
            
            ax.set_title(mes_key, fontsize=20, color='#005288', pad=20, fontweight='bold')
            ax.set_xlim(0, 100)
            ax.set_xlabel("Porcentaje de Estudiantes (%)")
            ax.set_ylabel("Grado" if i == 0 else "")
            
            if i == n_meses - 1:
                handles, labels = ax.get_legend_handles_labels()
                ax.legend(handles, labels, title="Nivel de Logro", bbox_to_anchor=(1.02, 1), loc='upper left')
            
            for p in ax.patches:
                width = p.get_width()
                if width > 4: 
                    x = p.get_x() + width / 2
                    y = p.get_y() + p.get_height() / 2
                    ax.text(x, y, f"{width:.1f}%", ha='center', va='center', color='white', fontsize=10, fontweight='bold')
                    
        plt.suptitle(f"Evolución Trimestral de Niveles Grupo B1: {materia}\n(De 2° a 11°)", fontsize=18)
        plt.tight_layout(rect=[0, 0, 0.9, 1]) 
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Barras_Horizontales_Comparativas_B1_{materia}.png"), dpi=300, bbox_inches='tight')
        plt.close()

# ==========================================
# 8. EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print(f"\n{'='*60}")
    print("  GENERADOR DE REPORTE EVOLUTIVO - GRUPOS B1 Y B2")
    print(f"{'='*60}")
    
    preparar_entorno()
    meses_labels = [etiqueta for _, etiqueta in DIRECTORIOS_MESES]
    
    # 1. Cargar Matrícula Oficial
    df_matricula = cargar_matricula()
    if df_matricula.empty:
        print("\n[!] El proceso se detuvo por falta de matrícula.")
        sys.exit()
        
    # 2. Extraer todos los puntajes
    df_all_scores = extraer_csvs_puntajes()
    if df_all_scores.empty:
        print("\n[!] No se extrajeron puntajes válidos.")
        sys.exit()
        
    # 3. Cruzar Puntajes con Matrícula (Documento == NIE)
    print("\n[*] Cruzando puntajes con la matrícula (Documento = NIE)...")
    df_merged = pd.merge(df_all_scores, df_matricula[['NIE_Limpio', 'GRUPO_Limpio']], left_on='Documento', right_on='NIE_Limpio', how='inner')
    
    # ----------------------------------------
    # PROCESAMIENTO GRUPO B1 (Evolutivo)
    # ----------------------------------------
    df_b1 = df_merged[df_merged['GRUPO_Limpio'] == 'B1'].copy()
    
    if df_b1.empty:
        print("  [!] No hay estudiantes cruzados que pertenezcan al Grupo B1.")
    else:
        # Aplicar filtro estricto a B1: Debe haber participado en M1 o M2
        if len(meses_labels) >= 3:
            meses_base = [meses_labels[0], meses_labels[1]]
            estudiantes_base = df_b1[df_b1['Mes'].isin(meses_base)][['Documento', 'Materia']].drop_duplicates()
            df_b1 = pd.merge(df_b1, estudiantes_base, on=['Documento', 'Materia'], how='inner')
            
        print(f"  [OK] Estudiantes válidos de B1 tras el filtro: {len(df_b1['Documento'].unique())}")
        
        exportar_base_consolidada_b1(df_b1)
        dibujar_sankey_multimes(df_b1, "Matemática", meses_labels)
        dibujar_sankey_multimes(df_b1, "Lengua", meses_labels)
        generar_graficos_estadisticos_b1(df_b1, meses_labels)
        
    # ----------------------------------------
    # PROCESAMIENTO GRUPO B2 (Solo Mes 3)
    # ----------------------------------------
    df_b2 = df_merged[df_merged['GRUPO_Limpio'] == 'B2'].copy()
    
    if not df_b2.empty and len(meses_labels) >= 3:
        mes_3_label = meses_labels[2]
        df_b2_mes3 = df_b2[df_b2['Mes'] == mes_3_label].copy()
        
        generar_excel_y_barras_b2(df_b2_mes3, df_matricula)
    else:
        print("\n  [!] No se encontraron datos para procesar el Grupo B2 en el Mes 3.")
        
    print(f"\n[OK] ¡Proceso completado con éxito! Todo está en:\n     {PATH_OUTPUT_PLOTS}\n")