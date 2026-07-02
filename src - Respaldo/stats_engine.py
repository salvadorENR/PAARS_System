# src/stats_engine.py
import pandas as pd
import numpy as np

def clasificar_puntaje(val):
    """Clasifica el puntaje theta en categorías pedagógicas institucionales."""
    if pd.isna(val) or val == '': return "Sin Datos"
    try:
        val = float(val)
        if val <= 35: return "Crítico"
        elif val <= 45: return "Bajo"
        elif val <= 55: return "Medio"
        elif val <= 65: return "Bueno"
        else: return "Excelente"
    except:
        return "Sin Datos"

def preprocesar_datos_validos(df):
    """Filtra pruebas anuladas y asegura conversiones numéricas para la estadística macro."""
    df_clean = df.copy()
    
    # 1. Quitar pruebas anuladas
    if 'anular_prueba' in df_clean.columns:
        df_clean = df_clean[df_clean['anular_prueba'].isna() | (df_clean['anular_prueba'].astype(str).str.strip() == '')]
    
    # 2. Tolerancia con el nombre de la columna de asignatura
    col_area = 'Area temática' if 'Area temática' in df_clean.columns else 'Area_Tematica'
    
    if col_area not in df_clean.columns:
        return pd.DataFrame(), col_area
        
    # 3. Limpiar y convertir puntajes
    df_clean['theta.global (escala 0-100)'] = pd.to_numeric(df_clean['theta.global (escala 0-100)'].astype(str).str.replace(',', '.'), errors='coerce')
    df_clean = df_clean.dropna(subset=['theta.global (escala 0-100)', col_area])
    
    return df_clean, col_area

def generar_estadisticas_basicas(df):
    """Calcula media, desviación estándar y participación de la población válida."""
    df_valid, col_area = preprocesar_datos_validos(df)
    if df_valid.empty: return pd.DataFrame()
    
    stats = df_valid.groupby(col_area)['theta.global (escala 0-100)'].agg(
        Media='mean',
        Desviacion_Std='std',
        Total_Evaluados='count'
    ).reset_index()
    
    return stats

def calcular_distribucion_niveles(df):
    """Calcula el porcentaje de alumnos válidos en cada nivel de desempeño."""
    df_valid, col_area = preprocesar_datos_validos(df)
    if df_valid.empty: return pd.DataFrame()
    
    # Aplicamos la clasificación
    df_valid['Nivel'] = df_valid['theta.global (escala 0-100)'].apply(clasificar_puntaje)
    
    # Contamos y calculamos porcentajes
    dist = df_valid.groupby([col_area, 'Nivel']).size().unstack(fill_value=0)
    dist_pct = dist.div(dist.sum(axis=1), axis=0) * 100
    
    return dist_pct.round(1)

def obtener_quintiles(df):
    """Divide a los estudiantes en 5 grupos iguales (Quintiles)."""
    df_valid, col_area = preprocesar_datos_validos(df)
    
    # Candado de seguridad: Evitar error si hay muy pocos datos
    if len(df_valid) < 5:
        return pd.DataFrame(columns=['Q1 (20%)', 'Q2 (40%)', 'Q3 (60%)', 'Q4 (80%)'])
    
    quintiles = df_valid.groupby(col_area)['theta.global (escala 0-100)'].quantile(
        [0.2, 0.4, 0.6, 0.8]
    ).unstack()
    
    quintiles.columns = ['Q1 (20%)', 'Q2 (40%)', 'Q3 (60%)', 'Q4 (80%)']
    return quintiles

if __name__ == "__main__":
    print("--- Motor Estadístico cargado y listo para procesar ---")