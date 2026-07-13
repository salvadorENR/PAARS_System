import os
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ==========================================
# CONFIGURACIÓN (Orden estricto de Arriba hacia Abajo)
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

MESES = ['Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre']
NUM_ESCUELAS = 150

def simular_datos_8_meses():
    print("[*] Simulando datos para 150 escuelas durante 8 meses...")
    np.random.seed(42)
    codigos_escuelas = [f"INFRA-{i:04d}" for i in range(1, NUM_ESCUELAS + 1)]
    lista_meses_dfs = []
    
    prob_inicial = [0.02, 0.08, 0.15, 0.35, 0.40]
    niveles_m1 = np.random.choice(NIVELES, size=NUM_ESCUELAS, p=prob_inicial)
    df_actual = pd.DataFrame({'Código': codigos_escuelas, 'Nivel': niveles_m1})
    lista_meses_dfs.append(df_actual)
    
    for mes in range(1, 8):
        df_previo = lista_meses_dfs[-1]
        nuevos_niveles = []
        for nivel in df_previo['Nivel']:
            idx_actual = NIVELES.index(nivel)
            salto = np.random.choice([1, 0, -1, -2], p=[0.10, 0.50, 0.30, 0.10])
            idx_nuevo = max(0, min(4, idx_actual + salto))
            nuevos_niveles.append(NIVELES[idx_nuevo])
        df_nuevo = pd.DataFrame({'Código': codigos_escuelas, 'Nivel': nuevos_niveles})
        lista_meses_dfs.append(df_nuevo)
        
    return lista_meses_dfs

def calc_y_centers(df_mes):
    """Calcula los centros Y dinámicamente para apilar sin superponer (Top-Down)."""
    totals = [len(df_mes[df_mes['Nivel'] == n]) for n in NIVELES]
    pad = 0.04
    grand = sum(totals) if sum(totals) > 0 else 1
    usable = 1.0 - pad * (len(NIVELES) - 1)
    
    y_centers = []
    cursor = 0.0
    for val in totals:
        h = max((val / grand) * usable, 0.005) # Altura mínima para nodos vacíos
        y_centers.append(round(cursor + h / 2, 4))
        cursor += h + pad
    return y_centers

def dibujar_sankey_simulado(lista_meses_datos):
    print("[*] Construyendo diagrama Sankey longitudinal anclado...")
    
    labels_nodos, colores_nodos_lista, node_x, node_y = [], [], [], []
    sources, targets, values, link_colors = [], [], [], []
    
    # 1. Crear los nodos apilados matemáticamente
    for i, (nombre_mes, df_mes) in enumerate(zip(MESES, lista_meses_datos)):
        x_val = 0.01 + i * (0.98 / max(1, len(MESES) - 1))
        centros_y = calc_y_centers(df_mes)
        
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre_mes}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])
            
    # 2. Inyectar flujos invisibles para forzar estructura
    for i in range(len(MESES) - 1):
        offset_origen = i * 5
        offset_destino = (i + 1) * 5
        for j in range(5):
            sources.append(offset_origen + j)
            targets.append(offset_destino + j)
            values.append(0.01) 
            link_colors.append('rgba(0,0,0,0)')
            
    # 3. Flujos Reales
    for i in range(len(lista_meses_datos) - 1):
        df_actual = lista_meses_datos[i]
        df_siguiente = lista_meses_datos[i+1]
        
        df_cruce = pd.merge(df_actual, df_siguiente, on='Código', suffixes=('_A', '_S'))
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
            
    fig = go.Figure(data=[go.Sankey(
        arrangement = "fixed", # Fija las coordenadas generadas por nuestra función
        node = dict(
            pad = 20, thickness = 24, # Match exacto con tu código de referencia
            line = dict(color = "white", width = 0.5),
            label = labels_nodos, color = colores_nodos_lista,
            x = node_x, y = node_y
        ),
        link = dict(
            source = sources, target = targets,
            value = values, color = link_colors
        )
    )])
    
    fig.update_layout(
        title_text="SIMULACIÓN: Evolución Longitudinal de Categorías Institucionales (8 Meses)<br><sup>Trayectoria de 150 Escuelas (Marzo a Octubre)</sup>",
        font_size=12, width=1600, height=700,
        margin=dict(l=20, r=20, t=90, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    
    out_file = "Simulacion_Sankey_8_Meses.html"
    fig.write_html(out_file)
    print(f"\n[OK] ¡Gráfico anclado generado exitosamente!")

if __name__ == "__main__":
    datos_simulados = simular_datos_8_meses()
    dibujar_sankey_simulado(datos_simulados)