# config.py
import os
import sys

# --- 1. CONFIGURACION DEL ENTORNO (LA RUTA BASE) ---
DRIVE_PATH = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse"
YEAR = "2026"
YEAR_DIR = os.path.join(DRIVE_PATH, YEAR)

# --- 2. ESCANEO AUTOMATICO DE CARPETAS ---
try:
    carpetas_disponibles = [d for d in os.listdir(YEAR_DIR) if os.path.isdir(os.path.join(YEAR_DIR, d))]
except FileNotFoundError:
    print(f"\n[!] ERROR CRITICO: No se encontro la ruta {YEAR_DIR}")
    print("Asegurate de que tu disco G: y Google Drive esten conectados.")
    sys.exit(1)

# --- 3. MENU INTERACTIVO EN CONSOLA ---
print("\n" + "="*50)
print("   PANEL DE CONFIGURACION PAARS")
print("="*50)
print("\nCarpetas encontradas en tu Drive:")
for i, folder in enumerate(carpetas_disponibles):
    print(f"  [{i+1}] {folder}")

# 3.1 Elegir el mes a procesar
while True:
    try:
        seleccion = int(input("\n-> Elige el NUMERO del mes que deseas PROCESAR: ")) - 1
        if 0 <= seleccion < len(carpetas_disponibles):
            MONTH_FOLDER = carpetas_disponibles[seleccion]
            break
        else:
            print("  [!] Numero fuera de rango. Intenta de nuevo.")
    except ValueError:
        print("  [!] Por favor, ingresa solo el numero.")

# 3.2 Determinar automaticamente el tipo de examen
if "PROGRESO" in MONTH_FOLDER.upper():
    EXAM_TYPE = "PROGRESO"
else:
    EXAM_TYPE = "RESULTADOS" 

# 3.3 Menu dinamico dependiente del tipo de examen
PREV_MONTH_FOLDER = None

if EXAM_TYPE == "PROGRESO":
    print("\nQue tipo de reportes deseas generar hoy para PROGRESO?")
    print("  [1] SOLO Reportes por Seccion (Profesores)")
    print("  [2] SOLO Reportes Comparativos (Directores)")
    print("  [3] AMBOS Reportes (Completo)")
    while True:
        try:
            REPORT_CHOICE = int(input("\n-> Elige una opcion (1, 2 o 3): "))
            if REPORT_CHOICE in [1, 2, 3]:
                break
            else:
                print("  [!] Opcion invalida.")
        except ValueError:
            print("  [!] Ingresa solo el numero.")
            
    # Elegir mes para comparar si eligió reportes de directores (Opcion 2 o 3)
    if REPORT_CHOICE in [2, 3]:
        print(f"\nCon que mes deseas COMPARAR los resultados de {MONTH_FOLDER}?")
        print("  [0] Ninguno (Es la primera prueba del anio)")
        for i, folder in enumerate(carpetas_disponibles):
            print(f"  [{i+1}] {folder}")
        
        while True:
            try:
                seleccion_prev = int(input("\n-> Elige el NUMERO del mes ANTERIOR (o 0 para omitir): "))
                if seleccion_prev == 0:
                    break
                elif 1 <= seleccion_prev <= len(carpetas_disponibles):
                    PREV_MONTH_FOLDER = carpetas_disponibles[seleccion_prev - 1]
                    break
                else:
                    print("  [!] Numero fuera de rango. Intenta de nuevo.")
            except ValueError:
                print("  [!] Por favor, ingresa solo el numero.")

else:
    # --- EL NUEVO MENÚ PARA LA PRUEBA DE RESULTADOS ---
    print("\nQue tipo de reportes deseas generar hoy para la Prueba de RESULTADOS?")
    print("  [1] SOLO Reporte Formal (LaTeX)")
    print("  [2] SOLO Reporte Corto (En construccion)")
    print("  [3] AMBOS Reportes")
    while True:
        try:
            REPORT_CHOICE = int(input("\n-> Elige una opcion (1, 2 o 3): "))
            if REPORT_CHOICE in [1, 2, 3]:
                break
            else:
                print("  [!] Opcion invalida.")
        except ValueError:
            print("  [!] Ingresa solo el numero.")

# --- MENSAJE DE TRANQUILIDAD ---
print("\n" + "="*50)
print("  [INICIANDO EL MOTOR PAARS... POR FAVOR ESPERA]")
print("  (Cargando librerias y leyendo archivos...)")
print("="*50 + "\n")

# --- 4. CONSTRUCCION AUTOMATICA DE RUTAS ---
CURRENT_MONTH_PATH = os.path.join(DRIVE_PATH, YEAR, MONTH_FOLDER)

PATH_RAW = os.path.join(CURRENT_MONTH_PATH, "Raw_Data")
PATH_INTERIM = os.path.join(CURRENT_MONTH_PATH, "Interim_CSVs")
PATH_REPORTS = os.path.join(CURRENT_MONTH_PATH, "Final_Reports")
PATH_METADATA = os.path.join(DRIVE_PATH, "00_Metadata")

if PREV_MONTH_FOLDER:
    PATH_PREV_INTERIM = os.path.join(DRIVE_PATH, YEAR, PREV_MONTH_FOLDER, "Interim_CSVs")
else:
    PATH_PREV_INTERIM = None