import pandas as pd
import os

# 1. Definir la ruta del archivo CSV
ruta_csv = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"

# 2. Ruta de guardado en "Documentos"
ruta_documentos = os.path.join(os.path.expanduser("~"), "Documents")
ruta_excel_salida = os.path.join(ruta_documentos, "Secciones_A_Reiniciar.xlsx")

try:
    # Cargar los datos
    df = pd.read_csv(ruta_csv, sep=';', encoding='utf-8')

    # --- FILTROS ---
    
    # 1. Escuelas solicitadas
    escuelas = [11244, 74005]
    
    # 2. Turnos: Matutino y Jornada Completa
    regex_turnos = r'(Mañana|Matutino|Jornada Completa)'
    
    # 3. Grados: Desde Segundo Grado hasta Segundo Año 
    # Esta lista garantiza que tome los grados correctos y sus variantes (Ej. "Segundo Grado - Año 1")
    regex_grados = r'^(Segundo Grado|Tercer Grado|Cuarto Grado|Quinto Grado|Sexto Grado|Séptimo Grado|Octavo Grado|Noveno Grado|Primer Año|Segundo Año)'

    # Aplicar todas las condiciones a la vez
    condicion_final = (
        (df['CODIGO'].isin(escuelas)) & 
        (df['TURNO'].str.contains(regex_turnos, case=False, na=False)) &
        (df['GRADO'].astype(str).str.contains(regex_grados, case=False, na=False, regex=True))
    )

    df_filtrado = df[condicion_final]

    # --- COLUMNAS Y GUARDADO ---
    
    columnas_finales = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NOMBRE_SECCIÓN', 'TURNO']
    
    # Eliminar duplicados para obtener la lista limpia de secciones
    df_resultado = df_filtrado[columnas_finales].drop_duplicates()

    # Guardar en Documentos
    df_resultado.to_excel(ruta_excel_salida, index=False)
    
    print(f"--- ¡ÉXITO! ---")
    print(f"Se encontraron {len(df_resultado)} secciones para reiniciar.")
    print(f"Archivo guardado en: {ruta_excel_salida}")

except Exception as e:
    print(f"❌ Ocurrió un error: {e}")