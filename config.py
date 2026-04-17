# config.py
import os
import sys

# =============================================================================
# 0. CONFIGURACIÓN DEL ENTORNO Y VALIDACIÓN DE RUTAS
# =============================================================================
DRIVE_PATH = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse"
YEAR = "2026"
YEAR_DIR = os.path.join(DRIVE_PATH, YEAR)

print("\n" + "="*55)
print("   SISTEMA DE EVALUACIÓN PAARS 2026 - INICIO")
print("="*55)

if not os.path.exists(YEAR_DIR):
    print(f"\n[!] ERROR CRÍTICO: No se encontró la ruta base:")
    print(f"    {YEAR_DIR}")
    print("    Asegúrese de que su disco G: y Google Drive estén conectados.")
    sys.exit(1)

# =============================================================================
# NIVEL 1: SELECCIÓN DE MÓDULO
# =============================================================================
while True:
    print("\n¿Qué acción requiere realizar el sistema?")
    print("  (1) Teoría de Respuesta al ítem")
    print("  (2) Generación de Reportes")
    
    try:
        mod_choice = int(input("-> Ingrese una opción (1 o 2): "))
        if mod_choice == 1:
            print("\n[!] El módulo de Teoría de Respuesta al ítem estará disponible en el futuro.")
        elif mod_choice == 2:
            break
        else:
            print("  [!] Opción inválida. Intente de nuevo.")
    except ValueError:
        print("  [!] Por favor, ingrese un número válido.")

# =============================================================================
# NIVEL 2: TIPO DE EXAMEN
# =============================================================================
while True:
    print("\nSeleccione el tipo de examen:")
    print("  (1) Prueba de Resultados")
    print("  (2) Prueba de Progreso")
    
    try:
        exam_choice = int(input("-> Ingrese una opción (1 o 2): "))
        if exam_choice == 1:
            EXAM_TYPE = "RESULTADOS"
            break
        elif exam_choice == 2:
            EXAM_TYPE = "PROGRESO"
            break
        else:
            print("  [!] Opción inválida. Intente de nuevo.")
    except ValueError:
        print("  [!] Por favor, ingrese un número válido.")

# =============================================================================
# NIVEL 3: SELECCIÓN DINÁMICA DEL MES/EXAMEN
# =============================================================================
# Filtramos las carpetas basándonos en el tipo de examen elegido
carpetas_disponibles = sorted([
    d for d in os.listdir(YEAR_DIR) 
    if os.path.isdir(os.path.join(YEAR_DIR, d)) and EXAM_TYPE in d.upper()
])

if not carpetas_disponibles:
    print(f"\n[!] ERROR: No se encontraron carpetas para 'Prueba de {EXAM_TYPE.capitalize()}' en {YEAR}.")
    sys.exit(1)

while True:
    print(f"\nExámenes disponibles detectados (Prueba de {EXAM_TYPE.capitalize()}):")
    for i, folder in enumerate(carpetas_disponibles):
        # Formateo amigable (Ej: "01_PROGRESO_Marzo" -> "Prueba de Progreso 1 (Marzo)")
        partes = folder.split('_')
        etiqueta = f"Prueba de {EXAM_TYPE.capitalize()} {partes[0].lstrip('0')} ({partes[-1]})" if len(partes) >= 3 else folder
        print(f"  ({i+1}) {etiqueta} [{folder}]")
        
    try:
        folder_choice = int(input(f"-> Seleccione el examen (1 al {len(carpetas_disponibles)}): ")) - 1
        if 0 <= folder_choice < len(carpetas_disponibles):
            MONTH_FOLDER = carpetas_disponibles[folder_choice]
            break
        else:
            print("  [!] Opción fuera de rango. Intente de nuevo.")
    except ValueError:
        print("  [!] Por favor, ingrese un número válido.")

# =============================================================================
# NIVEL 4: MENÚ DE ACCIONES (Dependiente del Tipo de Examen)
# =============================================================================
while True:
    print("\nSeleccione la acción a realizar:")
    
    if EXAM_TYPE == "RESULTADOS":
        print("  (1) Control de calidad de las bases de datos")
        print("  (2) Generar tablas y gráficas")
        print("  (3) Generar informes por escuelas")
        print("  (4) Generar tablas, gráficas e informes por escuela")
        opciones_validas = [1, 2, 3, 4]
        
    elif EXAM_TYPE == "PROGRESO":
        print("  (1) Control de calidad de las bases de datos")
        print("  (2) Generar informes para el estudiante (Aún no disponible)")
        print("  (3) Generar informes para profesores")
        print("  (4) Generar informes para directores")
        print("  (5) Generar los informes de estudiantes, profesores y directores")
        opciones_validas = [1, 2, 3, 4, 5]

    try:
        REPORT_CHOICE = int(input("-> Ingrese la opción deseada: "))
        if REPORT_CHOICE in opciones_validas:
            break
        else:
            print("  [!] Opción inválida para este tipo de examen.")
    except ValueError:
        print("  [!] Por favor, ingrese un número válido.")

# =============================================================================
# LÓGICA ADICIONAL: MES PREVIO PARA COMPARATIVAS (Opciones de Directores)
# =============================================================================
PREV_MONTH_FOLDER = None
if EXAM_TYPE == "PROGRESO" and REPORT_CHOICE in [4, 5]:
    print(f"\n¿Con qué mes desea COMPARAR los resultados de {MONTH_FOLDER}?")
    print("  (0) Ninguno (Es la primera prueba del año / No comparar)")
    for i, folder in enumerate(carpetas_disponibles):
        print(f"  ({i+1}) {folder}")
        
    while True:
        try:
            prev_choice = int(input("-> Seleccione el mes anterior (o 0 para omitir): "))
            if prev_choice == 0:
                break
            elif 1 <= prev_choice <= len(carpetas_disponibles):
                PREV_MONTH_FOLDER = carpetas_disponibles[prev_choice - 1]
                break
            else:
                print("  [!] Opción fuera de rango.")
        except ValueError:
            print("  [!] Por favor, ingrese un número válido.")

# =============================================================================
# 5. CONSTRUCCIÓN AUTOMÁTICA DE RUTAS (EXPORTACIÓN GLOBAL)
# =============================================================================
CURRENT_MONTH_PATH = os.path.join(YEAR_DIR, MONTH_FOLDER)

PATH_RAW = os.path.join(CURRENT_MONTH_PATH, "Raw_Data")
PATH_INTERIM = os.path.join(CURRENT_MONTH_PATH, "Interim_CSVs")
PATH_REPORTS = os.path.join(CURRENT_MONTH_PATH, "Final_Reports")
PATH_METADATA = os.path.join(DRIVE_PATH, "00_Metadata")

if PREV_MONTH_FOLDER:
    PATH_PREV_INTERIM = os.path.join(YEAR_DIR, PREV_MONTH_FOLDER, "Interim_CSVs")
else:
    PATH_PREV_INTERIM = None

print("\n" + "="*55)
print("  [CONFIGURACIÓN EXITOSA - INICIANDO MOTOR PAARS]")
print("="*55 + "\n")