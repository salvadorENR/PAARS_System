# src/visualizer.py
import matplotlib.pyplot as plt
import seaborn as sns
import os
import sys

# Conexión con el mapa de configuración para saber dónde guardar
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def configurar_estilo():
    """Prepara el papel y los colores para que se vean profesionales."""
    sns.set_theme(style="whitegrid")
    plt.rcParams['figure.figsize'] = (10, 6)

def generar_grafico_niveles(dist_pct):
    """Dibuja barras de colores para los niveles (Crítica a Excelente)."""
    configurar_estilo()
    
    # Colores: Rojo (Crítica), Naranja (Baja), Amarillo (Media), Verde claro (Buena), Verde oscuro (Excelente)
    colores = ["#c0392b", "#d35400", "#f1c40f", "#1abc9c", "#196f3d"]
    
    # Dibujamos las barras
    ax = dist_pct.plot(kind='barh', stacked=True, color=colores, width=0.8)
    
    plt.title("Distribución de Estudiantes por Nivel de Desempeño", fontsize=14)
    plt.xlabel("Porcentaje (%)")
    plt.ylabel("Materia")
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.xlim(0, 100)
    
    # Guardamos la foto en la carpeta de reportes del Drive
    ruta_guardado = os.path.join(config.PATH_REPORTS, "distribucion_niveles.png")
    plt.tight_layout()
    plt.savefig(ruta_guardado, dpi=300)
    plt.close()
    print(f"  [OK] Foto guardada: distribucion_niveles.png")

def generar_grafico_quintiles(quintiles):
    """Dibuja una línea de tiempo para ver cómo suben los puntajes."""
    configurar_estilo()
    
    plt.figure()
    for materia in quintiles.index:
        plt.plot(quintiles.columns, quintiles.loc[materia], marker='o', label=materia, linewidth=3)
    
    plt.title("Cortes de Puntaje por Quintil (Q1 - Q4)", fontsize=14)
    plt.ylabel("Puntaje (0-100)")
    plt.legend()
    
    ruta_guardado = os.path.join(config.PATH_REPORTS, "analisis_quintiles.png")
    plt.tight_layout()
    plt.savefig(ruta_guardado, dpi=300)
    plt.close()
    print(f"  [OK] Foto guardada: analisis_quintiles.png")