# src/visualizer.py
import matplotlib.pyplot as plt
import seaborn as sns
import os
import sys
import numpy as np

# Conexión con el mapa de configuración para saber dónde guardar
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

# --- PALETA INSTITUCIONAL UNIFICADA ---
COLOR_CATEGORIAS = {
    "Crítico": "#991b1b",
    "Bajo": "#ff8c2e",
    "Medio": "#facc15",
    "Bueno": "#84cc16",
    "Excelente": "#065f46"
}

def configurar_estilo():
    """Prepara el papel y los colores para que se vean profesionales."""
    sns.set_theme(style="whitegrid")
    plt.rcParams['figure.figsize'] = (12, 7) # Un poco más ancho para que quepan las etiquetas
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14

def generar_grafico_niveles(dist_pct):
    """Dibuja barras de colores para los niveles institucionales."""
    configurar_estilo()
    
    # Asegurarnos de que el orden de las columnas sea el correcto
    orden_cols = [c for c in COLOR_CATEGORIAS.keys() if c in dist_pct.columns]
    dist_pct = dist_pct[orden_cols]
    
    # Obtener los colores correspondientes
    colores = [COLOR_CATEGORIAS[c] for c in orden_cols]
    
    # Dibujamos las barras
    ax = dist_pct.plot(kind='barh', stacked=True, color=colores, width=0.7, edgecolor='white')
    
    plt.title("Distribución Global de Estudiantes por Nivel de Logro", fontsize=16, pad=15)
    plt.xlabel("Porcentaje de Estudiantes (%)", fontsize=12)
    plt.ylabel("Asignatura", fontsize=12)
    plt.legend(title="Nivel de Logro", bbox_to_anchor=(1.02, 1), loc='upper left')
    plt.xlim(0, 105) # Un poco de espacio extra
    
    # --- MAGIA: Poner las etiquetas de porcentaje adentro de las barras ---
    for container in ax.containers:
        # Solo mostrar la etiqueta si el porcentaje es mayor a 2% para que no se amontone
        labels = [f"{val:.1f}%" if val > 2 else "" for val in container.datavalues]
        ax.bar_label(container, labels=labels, label_type='center', color='white', fontweight='bold', fontsize=11)

    # Guardamos la foto en la carpeta de reportes
    ruta_guardado = os.path.join(config.PATH_REPORTS, "distribucion_niveles.png")
    plt.tight_layout()
    plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [OK] Gráfica ejecutiva guardada: distribucion_niveles.png")

def generar_grafico_quintiles(quintiles):
    """Dibuja una línea de tiempo para ver los cortes de puntaje y les pone etiquetas."""
    configurar_estilo()
    
    plt.figure(figsize=(10, 6))
    
    # Paleta para diferenciar materias
    colores_materias = ['#2980b9', '#8e44ad', '#27ae60', '#e67e22']
    
    for idx, materia in enumerate(quintiles.index):
        color = colores_materias[idx % len(colores_materias)]
        x = quintiles.columns
        y = quintiles.loc[materia]
        
        plt.plot(x, y, marker='o', label=materia, linewidth=3, markersize=8, color=color)
        
        # Poner las etiquetas numéricas en cada punto de corte
        for i, txt in enumerate(y):
            plt.annotate(f"{txt:.1f}", 
                         (x[i], y[i]), 
                         textcoords="offset points", 
                         xytext=(0,10), 
                         ha='center',
                         fontsize=11,
                         fontweight='bold',
                         color=color)
    
    plt.title("Cortes de Puntaje por Quintil (Q1 - Q4)", fontsize=16, pad=15)
    plt.ylabel("Puntaje Theta (0-100)", fontsize=12)
    plt.xlabel("Quintiles", fontsize=12)
    plt.ylim(0, 100) # Mantener la escala global 0-100
    plt.legend(bbox_to_anchor=(1.02, 1), loc='upper left')
    
    ruta_guardado = os.path.join(config.PATH_REPORTS, "analisis_quintiles.png")
    plt.tight_layout()
    plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  [OK] Gráfica ejecutiva guardada: analisis_quintiles.png")