# config.py
import os

# --- 1. CONFIGURACIÓN DEL ENTORNO (LA RUTA BASE) ---
# Esta es la parte fija de tu Google Drive
DRIVE_PATH = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse"

# --- 2. SELECCIÓN DE PROCESO (LO QUE CAMBIA CADA MES) ---
YEAR = "2026"
MONTH_FOLDER = "03_PROGRESO_Marzo"
EXAM_TYPE = "PROGRESO" 

# --- 3. CONSTRUCCIÓN AUTOMÁTICA DE RUTAS ---
CURRENT_MONTH_PATH = os.path.join(DRIVE_PATH, YEAR, MONTH_FOLDER)

PATH_RAW = os.path.join(CURRENT_MONTH_PATH, "Raw_Data")
PATH_INTERIM = os.path.join(CURRENT_MONTH_PATH, "Interim_CSVs")
PATH_REPORTS = os.path.join(CURRENT_MONTH_PATH, "Final_Reports")
PATH_METADATA = os.path.join(DRIVE_PATH, "00_Metadata")

# Quitamos el emoji para evitar el error de 'charmap'
print(f"[OK] Sistema configurado para: {MONTH_FOLDER}")