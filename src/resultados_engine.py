# src/resultados_engine.py
import pandas as pd
import numpy as np

def calcular_metricas_resultados(df_master, items_cols):
    """
    Calcula todas las estadísticas necesarias para el reporte LaTeX de Resultados.
    Asume que df_master ya viene limpio de Geiser.
    """
    print("  [Motor Resultados] Calculando estadísticas, quintiles y métricas de ítems...")
    
    report_data = {}
    
    # Asegurar que tenemos el grado numérico para iterar
    df_master['Grado_Num'] = df_master['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
    
    grados = sorted([g for g in df_master['Grado_Num'].unique() if g >= 3])
    
    for grado in grados:
        report_data[grado] = {'Matemática': None, 'Lengua': None}
        
        for materia in ['Matemática', 'Lengua']:
            # Filtrar por grado y materia, y asegurar que tengan puntaje IRT válido
            df_subset = df_master[(df_master['Grado_Num'] == grado) & 
                                  (df_master['Area temática'] == materia)].copy()
            
            df_subset['theta.global (escala 0-100)'] = pd.to_numeric(df_subset['theta.global (escala 0-100)'], errors='coerce')
            df_subset = df_subset.dropna(subset=['theta.global (escala 0-100)'])
            
            if df_subset.empty:
                continue
                
            N_validos = len(df_subset)
            puntajes = df_subset['theta.global (escala 0-100)']
            
            # 1. ESTADÍSTICAS DESCRIPTIVAS
            media = puntajes.mean()
            mediana = puntajes.median()
            ds = puntajes.std(ddof=1) if N_validos > 1 else 0
            minimo = puntajes.min()
            maximo = puntajes.max()
            
            desc_stats = {
                'N': N_validos,
                'Media': media,
                'Mediana': mediana,
                'DS': ds,
                'Mínimo': minimo,
                'Máximo': maximo
            }
            
            # 2. QUINTILES
            df_subset['Quintil'] = pd.qcut(puntajes.rank(method='first'), 5, labels=[1, 2, 3, 4, 5])
            quintiles_data = []
            for q in range(1, 6):
                grupo = df_subset[df_subset['Quintil'] == q]
                if not grupo.empty:
                    q_min = grupo['theta.global (escala 0-100)'].min()
                    q_max = grupo['theta.global (escala 0-100)'].max()
                    n_grupo = len(grupo)
                    pct_grupo = (n_grupo / N_validos) * 100
                    quintiles_data.append({
                        'Q': q,
                        'Rango': f"{q_min:.1f} - {q_max:.1f}",
                        'N': n_grupo,
                        '%': pct_grupo
                    })
            
            # 3. DIFICULTAD DE ÍTEMS (Porcentaje de aciertos)
            # Buscar qué ítems corresponden a esta materia y grado
            prefijo = 'MAT' if materia == 'Matemática' else 'LEC'
            items_materia = [col for col in items_cols if col.startswith(prefijo) and col in df_subset.columns]
            
            item_diff = []
            for item in items_materia:
                # Convertir a numérico por si acaso (1=Correcto, 0=Incorrecto)
                df_subset[item] = pd.to_numeric(df_subset[item], errors='coerce').fillna(0)
                # Solo tomamos en cuenta los ítems que fueron respondidos/evaluados
                if df_subset[item].notna().any():
                    pct_acierto = (df_subset[item] == 1).mean() * 100
                    item_diff.append({'Item': item, 'Pct': pct_acierto})
            
            df_items = pd.DataFrame(item_diff)
            
            # 4. TEXTOS DINÁMICOS (Para inyectar en LaTeX)
            texto_dinamico = {}
            if not df_items.empty:
                idx_min = df_items['Pct'].idxmin()
                idx_max = df_items['Pct'].idxmax()
                
                texto_dinamico = {
                    'item_min_nombre': df_items.loc[idx_min, 'Item'],
                    'item_min_val': df_items.loc[idx_min, 'Pct'],
                    'item_max_nombre': df_items.loc[idx_max, 'Item'],
                    'item_max_val': df_items.loc[idx_max, 'Pct']
                }
            
            # 5. DATOS PARA LAS GRÁFICAS
            promedios_escuelas = df_subset.groupby('Centro')['theta.global (escala 0-100)'].mean().reset_index()
            
            # Guardamos todo el paquete para este grado y materia
            report_data[grado][materia] = {
                'desc_stats': desc_stats,
                'quintiles': quintiles_data,
                'items_df': df_items,
                'textos': texto_dinamico,
                'df_estudiantes': df_subset[['Centro', 'theta.global (escala 0-100)']].copy(),
                'df_escuelas': promedios_escuelas
            }
            
    return report_data