# src/data_loader.py
import os
import pandas as pd
import re
import sys

# Conexión con el mapa de configuración
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def cargar_datos_limpios():
    print(f"--- Cargando archivos desde: {config.PATH_INTERIM}")
    
    archivos = [f for f in os.listdir(config.PATH_INTERIM) if f.endswith('.csv')]
    
    compl_list = []
    res_list = []

    for f in archivos:
        df = pd.read_csv(os.path.join(config.PATH_INTERIM, f), dtype=str)
        
        # Etiquetamos la materia según el nombre del archivo
        materia = 'Matematica' if 'MAT-' in f.upper() else 'Lengua'
        df['Area_Tematica'] = materia
        
        if '_compl' in f.lower():
            compl_list.append(df)
        else:
            res_list.append(df)

    if not compl_list or not res_list:
        print("[ERROR] No se encontraron archivos suficientes para unir.")
        return None, []

    # Unimos todos los fragmentos
    df_compl = pd.concat(compl_list, ignore_index=True)
    df_res = pd.concat(res_list, ignore_index=True)

    # --- LIMPIEZA CRÍTICA ---
    # Limpiamos el Documento (NIE) para que sea un texto puro sin '.0'
    for df in [df_compl, df_res]:
        df['Documento'] = df['Documento'].str.replace(r'\.0$', '', regex=True).str.strip()

    # Unimos los datos de los alumnos con sus resultados
    # Usamos Documento y Area_Tematica para que no se mezclen notas de MAT con LEC
    master_df = pd.merge(
        df_compl, 
        df_res[['Documento', 'Area_Tematica', 'theta.global (escala 0-100)', 'anular_prueba']], 
        on=['Documento', 'Area_Tematica'], 
        how='left'
    )

    # Detectamos automáticamente cuáles columnas son ítems (Ej: MAT01, LEC05)
    item_cols = [c for c in master_df.columns if re.match(r'^(MAT|LEC)\d+$', c)]
    
    # Convertimos la nota a número para poder hacer cálculos después
    master_df['theta.global (escala 0-100)'] = pd.to_numeric(master_df['theta.global (escala 0-100)'], errors='coerce')

    print(f"[OK] Datos integrados: {len(master_df)} registros encontrados.")
    print(f"[OK] items detectados: {len(item_cols)}")
    
    return master_df, item_cols

if __name__ == "__main__":
    df, items = cargar_datos_limpios()
    if df is not None:
        print("\nPrimeras 5 filas del sistema integrado:")
        print(df[['Documento', 'Area_Tematica', 'theta.global (escala 0-100)']].head())