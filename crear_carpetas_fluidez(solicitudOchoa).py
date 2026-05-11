import os
import pandas as pd
import re

def crear_carpetas_por_grupo():
    # 1. Definición de Rutas
    FILE_PATH = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\2036 CE Modernización Educativa.csv"
    BASE_DIR = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\FLUIDEZ"

    # 2. Crear la carpeta principal FLUIDEZ si no existe
    os.makedirs(BASE_DIR, exist_ok=True)
    print(f"[*] Verificando carpeta base: {BASE_DIR}")

    # 3. Leer el archivo CSV
    try:
        # Se intenta utf-8, si falla por acentos se intenta con latin1
        try:
            df = pd.read_csv(FILE_PATH, sep=';', dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(FILE_PATH, sep=';', dtype=str, encoding='latin1')
    except Exception as e:
        print(f"[!] Error al leer el archivo: {e}")
        return

    # Limpiar espacios en los nombres de las columnas
    df.columns = df.columns.str.strip()

    # 4. Validar las columnas requeridas
    col_cod = 'CÓDIGO CE'
    col_nom = 'NOMBRE CE'
    col_grupo = 'GRUPO'

    if col_cod not in df.columns or col_nom not in df.columns or col_grupo not in df.columns:
        print("[!] ERROR: No se encontraron las columnas especificadas en el archivo.")
        print(f"Columnas disponibles: {list(df.columns)}")
        return

    # 5. Filtrar las escuelas (Omitir aquellas donde 'GRUPO' está en blanco o es nulo)
    df = df.dropna(subset=[col_grupo])
    df = df[df[col_grupo].str.strip() != '']

    # Obtener lista única de escuelas con su grupo para evitar duplicados
    escuelas_unicas = df[[col_cod, col_nom, col_grupo]].drop_duplicates()
    
    print(f"[*] Se encontraron {len(escuelas_unicas)} centros educativos con un GRUPO válido.")
    print("[*] Generando árbol de carpetas y subcarpetas de grados...")

    carpetas_creadas = 0
    grupos_creados = set()
    
    # Lista de subcarpetas a crear en cada escuela
    grados_a_crear = ["2° Grado", "3° Grado", "4° Grado"]

    # 6. Iterar sobre cada escuela y crear las carpetas
    for index, row in escuelas_unicas.iterrows():
        grupo_val = str(row[col_grupo]).strip()
        codigo = str(row[col_cod]).strip().replace('.0', '')
        nombre_original = str(row[col_nom]).strip()
        
        # Limpiar caracteres que Windows no permite en nombres de carpetas (\ / : * ? " < > |)
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_original)
        
        # Definir nombres
        nombre_grupo_folder = f"GRUPO {grupo_val}"
        nombre_escuela_folder = f"{codigo}-{nombre_limpio}"
        
        # Construir la ruta base de la escuela: FLUIDEZ / GRUPO X / CODIGO-NOMBRE
        ruta_carpeta_escuela = os.path.join(BASE_DIR, nombre_grupo_folder, nombre_escuela_folder)
        
        # Crear las subcarpetas de los grados (esto también creará las carpetas padre si no existen)
        for grado in grados_a_crear:
            ruta_subcarpeta_grado = os.path.join(ruta_carpeta_escuela, grado)
            os.makedirs(ruta_subcarpeta_grado, exist_ok=True)
        
        grupos_creados.add(nombre_grupo_folder)
        carpetas_creadas += 1

    print("\n" + "="*60)
    print(f" ✅ ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print(f" Se crearon {len(grupos_creados)} carpetas de grupos: {', '.join(sorted(grupos_creados))}")
    print(f" Se organizaron en total {carpetas_creadas} carpetas de escuelas.")
    print(f" Y se añadieron 3 subcarpetas (2°, 3° y 4° Grado) dentro de cada escuela.")
    print(f" Directorio: {BASE_DIR}")
    print("="*60)

if __name__ == "__main__":
    crear_carpetas_por_grupo()