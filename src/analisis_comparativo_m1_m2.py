# src/analisis_comparativo_m1_m2.py
import os
import sys
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import seaborn as sns

# Conexión con el mapa PAARS
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# ==========================================
# CONFIGURACIÓN DINÁMICA
# ==========================================
# Las rutas ahora vienen directo del menú que el usuario eligió en config.py
PATH_MES_ANTERIOR = config.PATH_PREV_INTERIM
PATH_MES_ACTUAL = config.PATH_INTERIM
PATH_OUTPUT_PLOTS = os.path.join(config.PATH_REPORTS, "Plots_Comparativos_Intermensuales")

# Extracción dinámica de nombres para las etiquetas (Ej: "Marzo" y "Abril")
if config.PREV_MONTH_FOLDER:
    nombre_mes_ant = config.PREV_MONTH_FOLDER.split('_')[-1].capitalize()
else:
    nombre_mes_ant = "Mes_Anterior"
    
nombre_mes_act = config.MONTH_FOLDER.split('_')[-1].capitalize()

# Etiquetas de las gráficas
LBL_MES_ANT = f"Prueba Anterior ({nombre_mes_ant})"
LBL_MES_ACT = f"Prueba Actual ({nombre_mes_act})"

COL_SCORE = 'theta.global (escala 0-100)'
GRADOS_TODOS = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
GRADOS_BOXPLOT = [3, 6, 9]

COLOR_CATEGORIAS = {"Crítico": "#991b1b", "Bajo": "#ff8c2e", "Medio": "#facc15", "Bueno": "#84cc16", "Excelente": "#065f46"}
ORDEN_CATEGORIAS = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
MAPEO_NOMBRES_GRADO = {2: "Segundo Grado", 3: "Tercer Grado", 4: "Cuarto Grado", 5: "Quinto Grado", 6: "Sexto Grado", 7: "Séptimo Grado", 8: "Octavo Grado", 9: "Noveno Grado", 10: "Décimo Grado", 11: "Undécimo Grado"}

def preparar_entorno():
    sns.set_theme(style="whitegrid")
    plt.rcParams['figure.figsize'] = (14, 10) 
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14
    os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)

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

def cargar_y_procesar_mes(path_mes, etiqueta_mes):
    print(f"  -> Extrayendo datos de: {etiqueta_mes}...")
    archivos = glob.glob(os.path.join(path_mes, "*.csv"))
    if not archivos: return pd.DataFrame()

    lista_dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            materia = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lectura'
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
            df['Nivel_Logro'] = df[COL_SCORE].apply(clasificar_puntaje)
            df_final = df[[COL_SCORE, 'Grado_Num', 'Nivel_Logro']].copy()
            df_final['Materia'] = materia
            df_final['Mes'] = etiqueta_mes
            lista_dfs.append(df_final)
        except: pass
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()

def ejecutar_comparativo():
    print("\n[*] Iniciando Motor de Análisis Comparativo Intermensual...")
    if not config.PATH_PREV_INTERIM:
        print("[!] No se seleccionó un mes anterior para comparar. Módulo omitido.")
        return

    preparar_entorno()
    df_m1 = cargar_y_procesar_mes(PATH_MES_ANTERIOR, LBL_MES_ANT)
    df_m2 = cargar_y_procesar_mes(PATH_MES_ACTUAL, LBL_MES_ACT)
    
    if df_m1.empty or df_m2.empty:
        print("[!] Faltan datos en uno de los meses para realizar la comparativa.")
        return
        
    df_master = pd.concat([df_m1, df_m2], ignore_index=True)
    palette_meses = {LBL_MES_ANT: "#80b1d3", LBL_MES_ACT: "#fb8072"}

    for materia in df_master['Materia'].unique():
        print(f"    - Dibujando comparativas para {materia}...")
        df_sub = df_master[df_master['Materia'] == materia].copy()
        
        # 1. ESTADÍSTICAS
        writer_path = os.path.join(PATH_OUTPUT_PLOTS, f"Tablas_Comparativas_{materia}.xlsx")
        with pd.ExcelWriter(writer_path) as writer:
            for mes_nombre in df_sub['Mes'].unique():
                df_mes_stats = df_sub[df_sub['Mes'] == mes_nombre].copy()
                if df_mes_stats.empty: continue
                
                desc = df_mes_stats.groupby('Grado_Num')[COL_SCORE].describe().round(2).reset_index()
                desc['Asignatura/Grado'] = materia + " (" + desc['Grado_Num'].map(MAPEO_NOMBRES_GRADO) + ")"
                desc = desc.rename(columns={'mean': 'Media', '50%': 'Mediana', 'std': 'D.S.'})
                desc.sort_values('Grado_Num')[['Asignatura/Grado', 'Media', 'Mediana', 'D.S.']].to_excel(writer, sheet_name=mes_nombre[:31], index=False)

        # 2. KDE PLOT
        plt.figure(figsize=(12, 7))
        sns.histplot(data=df_sub, x=COL_SCORE, hue='Mes', kde=True, element="step", stat="probability", common_norm=False, palette=palette_meses, alpha=0.6)
        plt.title(f"Evolución de Puntajes Theta: {materia}\n({nombre_mes_ant} vs {nombre_mes_act})")
        plt.xlim(0, 100)
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, f"Distribucion_Evolutiva_{materia}.png"), dpi=300, bbox_inches='tight')
        plt.close()

if __name__ == "__main__":
    ejecutar_comparativo()