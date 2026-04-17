# src/stats_engine.py
import pandas as pd
import numpy as np

def clasificar_puntaje(val):
    """Clasifica el puntaje theta en categorías pedagógicas."""
    if pd.isna(val): return "Sin Datos"
    val = float(val)
    if val <= 35: return "Crítica"
    elif val <= 45: return "Baja"
    elif val <= 55: return "Media"
    elif val <= 65: return "Buena"
    else: return "Excelente"

def generar_estadisticas_basicas(df):
    """Calcula media, desviación estándar y participación."""
    stats = df.groupby('Area_Tematica')['theta.global (escala 0-100)'].agg(
        Media='mean',
        Desviacion_Std='std',
        Total_Evaluados='count'
    ).reset_index()
    return stats

def calcular_distribucion_niveles(df):
    """Calcula el porcentaje de alumnos en cada nivel de desempeño."""
    # Aplicamos la clasificación
    df['Nivel'] = df['theta.global (escala 0-100)'].apply(clasificar_puntaje)
    
    # Contamos y calculamos porcentajes
    dist = df.groupby(['Area_Tematica', 'Nivel']).size().unstack(fill_value=0)
    dist_pct = dist.div(dist.sum(axis=1), axis=0) * 100
    
    return dist_pct.round(1)

# src/stats_engine.py (Modifica esta función)

def obtener_quintiles(df):
    """Divide a los estudiantes en 5 grupos iguales (Quintiles)."""
    # Solo intentamos calcular si hay notas validas
    df_valid = df.dropna(subset=['theta.global (escala 0-100)'])
    
    if df_valid.empty:
        # Si no hay notas, devolvemos una tabla vacia con los nombres de columnas
        return pd.DataFrame(columns=['Q1 (20%)', 'Q2 (40%)', 'Q3 (60%)', 'Q4 (80%)'])
    
    quintiles = df_valid.groupby('Area_Tematica')['theta.global (escala 0-100)'].quantile(
        [0.2, 0.4, 0.6, 0.8]
    ).unstack()
    
    quintiles.columns = ['Q1 (20%)', 'Q2 (40%)', 'Q3 (60%)', 'Q4 (80%)']
    return quintiles

if __name__ == "__main__":
    # Prueba rápida (esto fallará si se corre solo, se debe llamar desde el integrador)
    print("--- Motor Estadístico cargado y listo para procesar ---")