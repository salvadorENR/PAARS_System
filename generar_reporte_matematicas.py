import pandas as pd
import os
import glob
import numpy as np
import re

def generar_reporte_consolidado_matematicas():
    print("="*60)
    print("🚀 INICIANDO SISTEMATIZACIÓN DE REPORTES DE MATEMÁTICAS")
    print("="*60)

    # ---------------------------------------------------------
    # 1. CONFIGURACIÓN DE RUTAS Y PARÁMETROS
    # ---------------------------------------------------------
    PATH_METADATA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata"
    FILE_MATRICULA = os.path.join(PATH_METADATA, "MatriculaProgresoMes2.csv")
    FOLDER_SCORED = os.path.join(PATH_METADATA, r"CML_OCT_2025_MINED\MAT_SCORED_CSV")
    OUTPUT_DIR = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents"
    OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Reporte_Notas_Matematicas_CML_Oct2025.xlsx")

    # Grados a procesar
    GRADOS_TARGET = [3, 4, 5, 6, 7, 8, 9]

    # Palabras clave para excluir columnas de metadatos y dejar solo los ítems del examen
    # (Añadimos 'codin' y 'codigo' porque 9° grado usa nombres diferentes)
    METADATA_KEYWORDS = ['id', 'email', 'time', 'duration', 'name', 'nombre', 'nie', 
                         'sede', 'codin', 'codigo', 'código', 'grado', 'seccion', 'score', 
                         'n°', 'nº', 'no.', 'numero', 'puntuaci', 'total', 'nota', 'estado', 'status']

    # ---------------------------------------------------------
    # 2. CARGAR FILTRO MAESTRO DE ESCUELAS (MATRÍCULA)
    # ---------------------------------------------------------
    print(f"[*] Cargando filtro de escuelas desde Matrícula...")
    try:
        try:
            df_mat = pd.read_csv(FILE_MATRICULA, sep=';', dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            df_mat = pd.read_csv(FILE_MATRICULA, sep=';', dtype=str, encoding='latin1')
            
        df_mat.columns = df_mat.columns.str.strip().str.upper()
        
        col_cod = next((c for c in df_mat.columns if 'CODIGO' in c or 'CÓDIGO' in c), None)
        col_nom = next((c for c in df_mat.columns if 'NOMBRE' in c and 'SECC' not in c), None)
        
        if not col_cod or not col_nom:
            print(f"[!] ERROR: No se encontraron columnas CODIGO/NOMBRE en {FILE_MATRICULA}")
            return

        df_mat_clean = df_mat[[col_cod, col_nom]].dropna().drop_duplicates()
        df_mat_clean[col_cod] = df_mat_clean[col_cod].astype(str).str.strip().str.replace('.0', '', regex=False)
        map_escuelas = pd.Series(df_mat_clean[col_nom].values, index=df_mat_clean[col_cod]).to_dict()
        
        codigos_escuelas_validas = set(map_escuelas.keys())
        print(f"    -> {len(codigos_escuelas_validas)} escuelas autorizadas cargadas.")

    except Exception as e:
        print(f"[!] Error crítico al cargar archivo de matrícula: {e}")
        return

    # ---------------------------------------------------------
    # 3. PROCESAMIENTO DE SCORED CSVs POR GRADO
    # ---------------------------------------------------------
    csv_files_in_folder = glob.glob(os.path.join(FOLDER_SCORED, "*.csv"))
    hojas_a_exportar = {}

    for grado_num in GRADOS_TARGET:
        grado_label = f"{grado_num}°"
        print(f"\n[*] Procesando {grado_label} Grado...")
        
        # Búsqueda robusta del archivo del grado (Ej: _9_, -9-, _9no_, _09_, _9º_)
        pattern = re.compile(rf"[_\-\s](0?{grado_num})([a-zA-Zº°]*)[_\-\s\.]", re.IGNORECASE)
        
        file_path = None
        for f in csv_files_in_folder:
            nombre_archivo = os.path.basename(f)
            
            # Evitar conflictos con 1 y 2
            if grado_num > 2 and ("_1" in nombre_archivo or "_2" in nombre_archivo): continue
            
            if pattern.search(nombre_archivo):
                file_path = f
                break
        
        if not file_path:
            print(f"    [!] No se encontró archivo SCORED para el grado {grado_label}. Saltando.")
            continue

        print(f"    -> Leyendo archivo: {os.path.basename(file_path)}")
        
        try:
            df_scored = pd.read_csv(file_path, dtype=str, encoding='utf-8')
            df_scored.columns = df_scored.columns.str.strip().str.lower()
            
            # Buscar dinámicamente la columna SEDE o CODIN
            col_sede = next((c for c in df_scored.columns if any(kw in c for kw in ['sede', 'codin', 'codigo', 'código'])), None)
            
            if not col_sede:
                print(f"    [!] ERROR: No se encontró columna 'sede' ni 'codin' en {os.path.basename(file_path)}. Saltando.")
                continue

            df_scored[col_sede] = df_scored[col_sede].astype(str).str.strip().str.replace('.0', '', regex=False)

            # --- A. FILTRADO EXCLUSIVO ---
            df_scored_filtered = df_scored[df_scored[col_sede].isin(codigos_escuelas_validas)].copy()
            
            if df_scored_filtered.empty:
                print(f"    [!] Ningún estudiante pertenece a las escuelas autorizadas. Saltando.")
                continue

            # --- B. DETECCIÓN DINÁMICA DE ÍTEMS ---
            cols_items = [c for c in df_scored_filtered.columns if not any(kw in c for kw in METADATA_KEYWORDS)]
            total_items_examen = len(cols_items)
            
            if total_items_examen == 0:
                print(f"    [!] ERROR: No se detectaron columnas de ítems (preguntas). Saltando.")
                continue
                
            print(f"    -> Detectados {total_items_examen} ítems en el examen.")

            # Convertir a numérico (asegurando sumar 1s y 0s)
            for col in cols_items:
                df_scored_filtered[col] = pd.to_numeric(df_scored_filtered[col], errors='coerce').fillna(0)

            df_scored_filtered['puntaje_calculado'] = df_scored_filtered[cols_items].sum(axis=1)

            # --- C. AGREGACIÓN POR ESCUELA ---
            grouped = df_scored_filtered.groupby(col_sede).agg(
                Estudiantes=(col_sede, 'size'),
                Promedio_Puntaje=('puntaje_calculado', 'mean')
            ).reset_index()

            # --- D. CONSTRUCCIÓN DE COLUMNAS DEL REPORTE ---
            grouped['Centro Educativo'] = grouped[col_sede].apply(
                lambda x: f"{x} - {map_escuelas.get(x, 'NOMBRE NO ENCONTRADO')}"
            )

            grouped['Prom. Bruto'] = grouped['Promedio_Puntaje'].apply(
                lambda x: f"{x:.2f} / {total_items_examen}"
            )

            grouped['redondeado_num'] = (grouped['Promedio_Puntaje'] + 0.000001).round(0).astype(int)
            grouped['redondeado_num'] = grouped['redondeado_num'].clip(upper=total_items_examen)
            
            grouped['Redondeado'] = grouped['redondeado_num'].apply(
                lambda x: f"{x} / {total_items_examen}"
            )

            denominator = total_items_examen if total_items_examen > 0 else 1
            grouped['% Promedio'] = (grouped['Promedio_Puntaje'] / denominator * 100).apply(
                lambda x: f"{x:.2f}%"
            )

            grouped['Nota (0-10)'] = (grouped['Promedio_Puntaje'] / denominator * 10).round(2)

            # --- E. GUARDAR EN MEMORIA ---
            df_final_sheet = grouped[[
                'Centro Educativo', 'Estudiantes', 'Prom. Bruto', 
                'Redondeado', '% Promedio', 'Nota (0-10)'
            ]].sort_values(by='Centro Educativo')

            hojas_a_exportar[grado_label] = df_final_sheet
            print(f"    -> ✔️ Hoja de {grado_label} procesada con {len(df_final_sheet)} escuelas.")

        except Exception as e:
            print(f"    [!] Error procesando el grado {grado_label}: {e}")

    # ---------------------------------------------------------
    # 4. EXPORTAR A EXCEL (CON FORMATO DE COLUMNAS)
    # ---------------------------------------------------------
    if len(hojas_a_exportar) > 0:
        print("\n[*] Generando archivo Excel...")
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            # Ordenamos las hojas para que queden siempre en orden (3°, 4°, etc.)
            for grado_label in sorted(hojas_a_exportar.keys()):
                df_sheet = hojas_a_exportar[grado_label]
                df_sheet.to_excel(writer, sheet_name=grado_label, index=False)
                
                # Ajustar el ancho de las columnas
                worksheet = writer.sheets[grado_label]
                
                # Diccionario de anchos: {Letra de Columna : Ancho}
                anchos_columnas = {
                    'A': 65,  # Centro Educativo (largo)
                    'B': 15,  # Estudiantes
                    'C': 15,  # Prom. Bruto
                    'D': 15,  # Redondeado
                    'E': 15,  # % Promedio
                    'F': 15   # Nota (0-10)
                }
                
                for letra, ancho in anchos_columnas.items():
                    worksheet.column_dimensions[letra].width = ancho
                
        print("="*60)
        print(f" ✅ ¡PROCESO COMPLETADO!")
        print(f" Se procesaron {len(hojas_a_exportar)} grados exitosamente.")
        print(f" Archivo guardado en: {OUTPUT_FILE}")
        print("="*60)
    else:
        print("\n" + "="*60)
        print(" [!] ATENCIÓN: No se pudo generar el reporte.")
        print(" Ninguno de los archivos procesados tenía datos válidos o compatibles.")
        print("="*60)

if __name__ == "__main__":
    # Asegurar que openpyxl esté instalado para el ajuste de columnas
    try:
        import openpyxl
        generar_reporte_consolidado_matematicas()
    except ImportError:
        print("[!] ERROR: La librería 'openpyxl' no está instalada.")
        print("    Por favor ejecute: pip install openpyxl")