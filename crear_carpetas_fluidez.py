import os
import pandas as pd
import re

def crear_carpetas_escuelas():
    # 1. Definición de Rutas
    FILE_PATH = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
    BASE_DIR = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\FLUIDEZ"

    # 2. Crear la carpeta principal FLUIDEZ si no existe
    os.makedirs(BASE_DIR, exist_ok=True)
    print(f"[*] Verificando carpeta base: {BASE_DIR}")

    # 3. Leer el archivo CSV de matrícula (separado por punto y coma)
    try:
        df = pd.read_csv(FILE_PATH, sep=';', dtype=str, encoding='utf-8')
    except Exception as e:
        print(f"[!] Error al leer el archivo: {e}")
        return

    # 4. Identificar las columnas de Código y Nombre dinámicamente
    col_cod = next((c for c in df.columns if 'codigo' in c.lower() or 'código' in c.lower()), None)
    col_nom = next((c for c in df.columns if 'nombre' in c.lower() and 'secc' not in c.lower()), None)

    if not col_cod or not col_nom:
        print("[!] ERROR: No se encontraron las columnas de Código o Nombre del centro.")
        print(f"Columnas disponibles: {list(df.columns)}")
        return

    # 5. Obtener lista única de escuelas (sin duplicados)
    escuelas_unicas = df[[col_cod, col_nom]].drop_duplicates().dropna()
    
    print(f"[*] Se encontraron {len(escuelas_unicas)} centros educativos únicos.")
    print("[*] Creando carpetas...")

    carpetas_creadas = 0

    # 6. Iterar sobre cada escuela y crear su carpeta
    for index, row in escuelas_unicas.iterrows():
        codigo = str(row[col_cod]).strip().replace('.0', '')
        nombre_original = str(row[col_nom]).strip()
        
        # Limpiar caracteres que Windows no permite en nombres de carpetas (\ / : * ? " < > |)
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_original)
        
        # Ensamblar el nombre de la carpeta (Ej: 10747854-CENTRO ESCOLAR COLONIA LAS CAÑAS)
        nombre_carpeta = f"{codigo}-{nombre_limpio}"
        ruta_carpeta = os.path.join(BASE_DIR, nombre_carpeta)
        
        # Crear la carpeta
        os.makedirs(ruta_carpeta, exist_ok=True)
        carpetas_creadas += 1

    print("\n" + "="*50)
    print(f" ✅ ¡PROCESO COMPLETADO!")
    print(f" Se crearon o verificaron {carpetas_creadas} carpetas en:")
    print(f" {BASE_DIR}")
    print("="*50)

if __name__ == "__main__":
    crear_carpetas_escuelas()