# config.py
import os
import sys
from dotenv import load_dotenv

# 1. Cargar las variables (busca el archivo .env automáticamente)
load_dotenv() 

# 2. Asignar la ruta extrayéndola de la variable
DRIVE_PATH = os.getenv("BASE_DRIVE")

# 3. Validación de seguridad
if not DRIVE_PATH:
    print("\n[!] ERROR CRÍTICO: No se encontró la variable BASE_DRIVE en el archivo .env.")
    sys.exit(1)

# 4. DEFINICIÓN DEL AÑO (¡Esta es la parte que falta!)
YEAR = "2026"
YEAR_DIR = os.path.join(DRIVE_PATH, YEAR)

if not os.path.exists(YEAR_DIR):
    print(f"\n[!] ERROR CRÍTICO: No se encontró la ruta {YEAR_DIR}")
    sys.exit(1)

# --- NIVEL 1 & 2 ---
# (El resto de tus menús sigue aquí...)

# --- NIVEL 1 & 2 ---
while True:
    print("\n¿Qué acción requiere realizar el sistema?\n  (1) Teoría de Respuesta al ítem\n  (2) Generación de Reportes")
    mod_choice = input("-> ").strip()
    if mod_choice == '1': print("\n[!] Disponible en el futuro."); continue
    if mod_choice == '2': break

while True:
    print("\nSeleccione el tipo de examen:\n  (1) Prueba de Resultados\n  (2) Prueba de Progreso\n  (3) Prueba de Fundamento")
    exam_choice = input("-> ").strip()
    if exam_choice == '1': EXAM_TYPE = "RESULTADOS"; break
    if exam_choice == '2': EXAM_TYPE = "PROGRESO"; break
    if exam_choice == '3': EXAM_TYPE = "FUNDAMENTO"; break

# --- NIVEL 3 (Selección Dinámica) ---
carpetas_disponibles = sorted([d for d in os.listdir(YEAR_DIR) if os.path.isdir(os.path.join(YEAR_DIR, d)) and EXAM_TYPE in d.upper()])

while True:
    print(f"\nCarpetas disponibles ({EXAM_TYPE}):")
    for i, f in enumerate(carpetas_disponibles): print(f"  ({i+1}) {f}")
    try:
        idx = int(input("-> Seleccione el número: ")) - 1
        MONTH_FOLDER = carpetas_disponibles[idx]; break
    except: print("Selección inválida.")

# --- NIVEL 4: MENÚ GRANULAR ACTUALIZADO ---
while True:
    print(f"\nMenú de Acciones para {EXAM_TYPE}:")
    if EXAM_TYPE == "RESULTADOS":
        print("  (1) Integrar Archivos Crudos (Unificar MAT y LENGUA en un Dataset)")
        print("  (2) Control de Calidad")
        print("  (3) Reporte Formal (LaTeX Tables/Graphs)")
        print("  (4) Informes por Escuela")
        print("  (5) Reporte Nacional por Grados (HTML)")
        print("  (6) Rankings Macro y Evolución (Excel)")
        print("  (7) TODO lo anterior")
        validas = ['1','2','3','4','5','6','7']
    elif EXAM_TYPE == "PROGRESO":
        print("  (1) Integrar Archivos Crudos (Unificar MAT y LENGUA en un Dataset)")
        print("  (2) Control de Calidad")
        print("  (3) Informes para Profesores")
        print("  (4) Informes para Directores")
        print("  (5) Reporte Nacional por Grados (HTML)")
        print("  (6) Reporte Corto Beamer (LaTeX)")
        print("  (7) Rankings Macro y Evolución (Excel)")
        print("  (8) TODO lo anterior")
        validas = ['1','2','3','4','5','6','7','8']
    elif EXAM_TYPE == "FUNDAMENTO":
        print("  (1) Sistematizar CSV Crudo a Datasets (Excel) y Reporte Gráfico (HTML)")
        validas = ['1']
    
    REPORT_CHOICE = input("-> Elija una opción: ").strip()
    if REPORT_CHOICE in validas: break

# --- CONFIGURACIÓN DE RUTAS ---
CURRENT_MONTH_PATH = os.path.join(YEAR_DIR, MONTH_FOLDER)
PATH_RAW = os.path.join(CURRENT_MONTH_PATH, "Raw_Data")

if EXAM_TYPE == "FUNDAMENTO":
    PATH_INTERIM = os.path.join(CURRENT_MONTH_PATH, "02_Datasets_Procesados")
    PATH_REPORTS = os.path.join(CURRENT_MONTH_PATH, "03_Reportes")
else:
    PATH_INTERIM = os.path.join(CURRENT_MONTH_PATH, "Interim_CSVs")
    PATH_REPORTS = os.path.join(CURRENT_MONTH_PATH, "Final_Reports")

PATH_METADATA = os.path.join(DRIVE_PATH, "00_Metadata")
PATH_PREV_INTERIM = None