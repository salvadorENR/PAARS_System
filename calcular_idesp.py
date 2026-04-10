# calcular_idesp.py - VERSIÓN DIRECTA MARZO
import os
import pandas as pd
import numpy as np
import glob

# --- 1. CONFIGURACIÓN DE RUTAS DIRECTAS (PARA EVITAR EL MENÚ) ---
RUTA_METADATA = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata"
RUTA_INTERIM  = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs"
RUTA_SALIDA   = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo"

# --- 2. CONFIGURACIÓN DE NIVELES IDESP ---
CORTE_CRITICO = 35.0
CORTE_BAJO    = 45.0
CORTE_MEDIO   = 55.0
CORTE_BUENO   = 65.0

PESO_CRITICO  = 4
PESO_BAJO     = 3
PESO_MEDIO    = 2
PESO_BUENO    = 1
PESO_EXCELENTE = 0
MAX_PESO      = 4 

def asignar_peso_rezago(score):
    if pd.isna(score): return np.nan
    val = float(score)
    if val <= CORTE_CRITICO: return PESO_CRITICO
    elif val <= CORTE_BAJO:  return PESO_BAJO
    elif val <= CORTE_MEDIO: return PESO_MEDIO
    elif val <= CORTE_BUENO: return PESO_BUENO
    else:                    return PESO_EXCELENTE

def ejecutar_calculo_directo():
    print("====================================================")
    print("   GENERANDO RANKING IDESP 100 - MARZO")
    print("====================================================\n")
    
    # 1. Cargar el archivo de promedios de la carpeta Metadata
    archivos_ranking = glob.glob(os.path.join(RUTA_METADATA, "Ranking_Centros_Semáforo*Marzo*.csv"))
    if not archivos_ranking:
        print(f"[!] ERROR: No se encontró el CSV de Semáforo en Metadata.")
        return
    
    df_avg_file = pd.read_csv(archivos_ranking[0], sep=';', dtype=str, encoding='utf-8-sig')
    df_avg_file.columns = df_avg_file.columns.str.strip()
    df_avg_file = df_avg_file.rename(columns={'Nro centro': 'cod_infra', 'Media': 'Promedio_Original'})
    df_avg_file['cod_infra'] = df_avg_file['cod_infra'].str.strip()

    # 2. Procesar archivos Geiser de la carpeta Interim de Marzo
    dfs_res = []
    print("[*] Leyendo datos de Geiser en Marzo...")
    for f in os.listdir(RUTA_INTERIM):
        if f.lower().endswith('.csv') and 'resultados' in f.lower():
            path = os.path.join(RUTA_INTERIM, f)
            try:
                df = pd.read_csv(path, dtype=str, encoding='utf-8-sig', on_bad_lines='skip')
                col_cod = next((c for c in df.columns if 'nro de centro' in str(c).lower() or 'código' in str(c).lower()), None)
                col_nom = next((c for c in df.columns if 'centro' in str(c).lower() and 'nro' not in str(c).lower()), None)
                col_theta = 'theta.global (escala 0-100)'

                if col_cod and col_nom and col_theta in df.columns:
                    temp = df[[col_cod, col_nom, col_theta]].copy()
                    temp.columns = ['cod_infra', 'Centro', 'Puntaje_Raw']
                    dfs_res.append(temp)
            except: pass

    if not dfs_res:
        print("[!] No hay archivos de resultados en la carpeta Interim de Marzo.")
        return

    # 3. Matemática IDESP
    df_master = pd.concat(dfs_res, ignore_index=True)
    df_master['Puntaje_Num'] = pd.to_numeric(df_master['Puntaje_Raw'].str.replace('"', '').str.replace(',', '.'), errors='coerce')
    df_master = df_master.dropna(subset=['Puntaje_Num'])
    df_master['Lag_Weight'] = df_master['Puntaje_Num'].apply(asignar_peso_rezago)

    df_idesp = df_master.groupby('cod_infra').agg(
        Centro=('Centro', 'first'),
        Evaluados=('Puntaje_Num', 'count'),
        D_Lag=('Lag_Weight', 'mean')
    ).reset_index()

    df_idesp['IDESP'] = (1 - (df_idesp['D_Lag'] / MAX_PESO)) * 100

    # 4. Cruce Final y Exportación
    print("[*] Cruzando con promedios originales...")
    df_final = pd.merge(df_idesp, df_avg_file[['cod_infra', 'Promedio_Original']], on='cod_infra', how='inner')
    df_final = df_final.sort_values(by='IDESP', ascending=False)
    
    # Redondeos y Limpieza de columnas
    df_final['IDESP'] = df_final['IDESP'].round(2)
    df_final['Promedio_Original'] = pd.to_numeric(df_final['Promedio_Original']).round(2)
    
    df_export = df_final[['cod_infra', 'Centro', 'Promedio_Original', 'IDESP', 'Evaluados']].copy()
    df_export.columns = ['Nro de Centro', 'Centro', 'Promedio General', 'IDESP (Calidad)', 'Total Evaluados']

    archivo_excel = os.path.join(RUTA_SALIDA, "Ranking_General_IDESP_100_Marzo.xlsx")
    df_export.to_excel(archivo_excel, index=False)

    print(f"\n[OK] ¡Archivo creado con éxito!")
    print(f"📁 Ruta: {archivo_excel}")

if __name__ == "__main__":
    ejecutar_calculo_directo()