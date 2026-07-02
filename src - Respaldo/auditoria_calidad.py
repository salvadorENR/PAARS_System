# src/auditoria_calidad.py
import os
import glob
import re
import pandas as pd
import sys

# Conexión con el mapa PAARS
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def clean_option(x):
    """Filtro estricto para limpieza de opciones (A, B, C, D, N/R)."""
    if pd.isna(x): return 'N/R'
    val = str(x).strip().replace('"', '').replace('.', '').upper()
    if val in ['A', 'B', 'C', 'D']: return val
    return 'N/R'

def auditoria_vacios(df_master, output_dir):
    """
    Identifica ítems vacíos (NaN) y los cruza con la columna 'anular_prueba' 
    para separar descuido estudiantil de invalidaciones legítimas.
    """
    print("    -> Ejecutando: Auditoría de Ítems Vacíos y Anulaciones...")
    
    # Identificar columnas de ítems (Ej: MAT1, LEC12)
    item_cols = [c for c in df_master.columns if re.match(r'^(MAT|LEC)\d+$', str(c).upper())]
    if not item_cols:
        print("       [!] No se encontraron columnas de ítems para auditar vacíos.")
        return

    # Asegurar que exista la columna anular_prueba
    col_anular = next((c for c in df_master.columns if 'anular' in c.lower()), None)
    if not col_anular:
        df_master['anular_prueba'] = 'Sin Observación'
        col_anular = 'anular_prueba'

    # Llenar vacíos en la columna de anulación para agrupar correctamente
    df_master[col_anular] = df_master[col_anular].fillna('Sin Observación').replace('', 'Sin Observación')

    # Contar NaNs por ítem agrupado por estado de anulación
    res_list = []
    for estado, grupo in df_master.groupby(col_anular):
        n_estudiantes = len(grupo)
        for item in item_cols:
            vacios = grupo[item].isna().sum()
            if vacios > 0:
                res_list.append({
                    'Estado_Prueba': estado,
                    'Item': item,
                    'Total_Estudiantes_Grupo': n_estudiantes,
                    'Respuestas_Vacias (NaN)': vacios,
                    '%_Vacio_en_Grupo': round((vacios / n_estudiantes) * 100, 2)
                })
                
    if res_list:
        df_vacios = pd.DataFrame(res_list).sort_values(by=['Estado_Prueba', '%_Vacio_en_Grupo'], ascending=[True, False])
        df_vacios.to_excel(os.path.join(output_dir, "Auditoria_Vacios_Descuido_vs_Anulacion.xlsx"), index=False)
        print("       [OK] Reporte de vacíos generado.")
    else:
        print("       [OK] No se detectaron ítems vacíos alarmantes.")

def auditoria_cobertura_curricular(output_dir):
    """
    Lee la metadata dinámica y cuenta la cobertura curricular de la prueba (indicadores e intervalos).
    """
    print("    -> Ejecutando: Auditoría de Cobertura Curricular...")
    
    # Buscar el maestro de ítems correspondiente al mes actual
    archivos_meta = glob.glob(os.path.join(config.PATH_METADATA, "*.xlsx")) + glob.glob(os.path.join(config.PATH_METADATA, "*.csv"))
    
    # Extraer el número del mes del config (Ej: 01_PROGRESO_Marzo -> 1)
    match_mes = re.search(r'^0?(\d+)_', config.MONTH_FOLDER)
    mes_num = match_mes.group(1) if match_mes else "1"
    
    archivo_maestro = None
    for f in archivos_meta:
        nombre = os.path.basename(f).lower()
        if config.EXAM_TYPE.lower() in nombre and (f"mes{mes_num}" in nombre or f"_{mes_num}_" in nombre or f"mes {mes_num}" in nombre):
            archivo_maestro = f
            break

    if not archivo_maestro:
        print("       [!] No se encontró metadata para auditar cobertura curricular.")
        return

    try:
        if archivo_maestro.endswith('.csv'):
            df_meta = pd.read_csv(archivo_maestro, encoding='utf-8-sig', sep=None, engine='python')
        else:
            df_meta = pd.read_excel(archivo_maestro)
            
        col_indicador = next((c for c in df_meta.columns if 'indicador' in c.lower()), None)
        col_grado = next((c for c in df_meta.columns if 'grado' in c.lower()), None)
        
        if col_indicador and col_grado:
            cobertura = df_meta.groupby(col_grado)[col_indicador].nunique().reset_index()
            cobertura.columns = ['Grado', 'Total_Indicadores_Unicos_Evaluados']
            cobertura.to_excel(os.path.join(output_dir, "Auditoria_Cobertura_Curricular.xlsx"), index=False)
            print("       [OK] Reporte de cobertura curricular generado.")
        else:
            print("       [!] Faltan columnas clave (Indicador/Grado) en la metadata.")
    except Exception as e:
        print(f"       [!] Error leyendo metadata: {e}")

def auditoria_frecuencia_opciones(output_dir):
    """
    Busca datos RAW en TXT, limpia con clean_option y calcula frecuencias A,B,C,D,N/R.
    """
    print("    -> Ejecutando: Auditoría de Frecuencia de Opciones de Respuesta...")
    
    # Buscar TXTs en Raw_Data
    txt_files = glob.glob(os.path.join(config.PATH_RAW, "**", "*.txt"), recursive=True)
    if not txt_files:
        print("       [!] No se encontraron archivos TXT en Raw_Data para procesar opciones.")
        return
        
    print(f"       [*] Procesando {len(txt_files)} archivos TXT. Esto puede tardar unos segundos...")
    
    lista_dfs = []
    for file in txt_files:
        try:
            df = pd.read_csv(file, sep='|', quotechar='"', dtype=str, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(file, sep='|', quotechar='"', dtype=str, encoding='latin-1')
            except:
                continue
        df.columns = df.columns.str.strip().str.replace('"', '')
        lista_dfs.append(df)
        
    if not lista_dfs: return
    
    df_raw = pd.concat(lista_dfs, ignore_index=True)
    item_cols = [c for c in df_raw.columns if re.match(r'^(MAT|LEC)\d+$', str(c).upper())]
    
    if not item_cols:
        print("       [!] No se encontraron columnas de ítems en los TXT.")
        return
        
    # Optimización: Derretir (melt) solo las columnas necesarias
    df_melt = df_raw.melt(id_vars=['Grado'], value_vars=item_cols, var_name='Item', value_name='Respuesta')
    df_melt['Respuesta'] = df_melt['Respuesta'].apply(clean_option)
    
    # Calcular frecuencias
    freq = df_melt.groupby(['Grado', 'Item', 'Respuesta']).size().unstack(fill_value=0)
    for col in ['A', 'B', 'C', 'D', 'N/R']:
        if col not in freq.columns: freq[col] = 0
        
    # Calcular porcentajes
    total = freq.sum(axis=1).replace(0, 1)
    freq_pct = freq.div(total, axis=0) * 100
    freq_pct = freq_pct.round(2).reset_index()
    
    freq_pct.to_excel(os.path.join(output_dir, "Auditoria_Frecuencias_Opciones.xlsx"), index=False)
    print("       [OK] Reporte de frecuencia de opciones generado.")

def auditoria_representatividad(df_master, output_dir, lista_escuelas=None):
    """
    Verifica la representatividad de una escuela o grupo de escuelas (Ej: 99999 - Virtual) 
    frente al total nacional evaluado, usando drop_duplicates por Documento.
    """
    print("    -> Ejecutando: Auditoría de Representatividad Demográfica...")
    
    col_doc = next((c for c in df_master.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), None)
    col_centro = next((c for c in df_master.columns if 'nro de centro' in str(c).lower() or 'centro' in str(c).lower() or 'código' in str(c).lower()), None)
    
    if not col_doc or not col_centro:
        print("       [!] Faltan columnas de NIE o Código de Centro para calcular representatividad.")
        return

    # Limpiar duplicados reales (Un estudiante es una persona, sin importar cuántas pruebas hizo)
    df_unicos = df_master.drop_duplicates(subset=[col_doc]).copy()
    df_unicos[col_centro] = df_unicos[col_centro].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    total_nacional = len(df_unicos)
    if total_nacional == 0: return
    
    # Análisis general de las escuelas top y virtuales
    conteo_escuelas = df_unicos.groupby(col_centro).size().reset_index(name='Estudiantes_Evaluados')
    conteo_escuelas['%_del_Total_Nacional'] = (conteo_escuelas['Estudiantes_Evaluados'] / total_nacional * 100).round(3)
    conteo_escuelas = conteo_escuelas.sort_values(by='Estudiantes_Evaluados', ascending=False)
    
    conteo_escuelas.to_excel(os.path.join(output_dir, "Auditoria_Representatividad_Nacional.xlsx"), index=False)
    
    # Búsqueda específica solicitada
    if lista_escuelas:
        print(f"       [*] Buscando representatividad específica de los códigos: {lista_escuelas}")
        for cod in lista_escuelas:
            datos_escuela = conteo_escuelas[conteo_escuelas[col_centro] == str(cod)]
            if not datos_escuela.empty:
                pct = datos_escuela['%_del_Total_Nacional'].iloc[0]
                n_est = datos_escuela['Estudiantes_Evaluados'].iloc[0]
                print(f"           - Código {cod}: {n_est} estudiantes ({pct}% del total nacional).")
            else:
                print(f"           - Código {cod}: No detectado en la base de datos validada.")
                
    print("       [OK] Reporte de representatividad generado.")

# =============================================================================
# ORQUESTADOR PRINCIPAL
# =============================================================================
def auditoria_calidad():
    """
    Función principal que invoca las auditorías secuencialmente.
    Esta función es llamada directamente desde main.py (Opción 1).
    """
    # 1. Crear carpeta de salida específica para la auditoría
    AUDIT_DIR = os.path.join(config.PATH_REPORTS, "00_Auditoria_Calidad_DB")
    os.makedirs(AUDIT_DIR, exist_ok=True)
    
    print(f"\n[INICIANDO BATERÍA DE AUDITORÍAS]")
    print(f"Directorio de salida: {AUDIT_DIR}\n")
    
    # 2. Cargar la base de datos unificada de Interim
    # Cargamos CSVs de la carpeta Interim para las auditorías de vacíos y representatividad
    archivos_interim = glob.glob(os.path.join(config.PATH_INTERIM, "*.csv"))
    df_master = pd.DataFrame()
    
    if archivos_interim:
        lista_interim = []
        for f in archivos_interim:
            try:
                df = pd.read_csv(f, dtype=str, encoding='utf-8-sig', sep=None, engine='python')
                lista_interim.append(df)
            except Exception:
                pass
        if lista_interim:
            df_master = pd.concat(lista_interim, ignore_index=True)
            # Limpieza básica de nombres de columnas
            df_master.columns = df_master.columns.str.strip()

    if df_master.empty:
        print("[!] No se pudo cargar la base de datos Interim para la auditoría.")
        return

    # 3. Ejecutar funciones en secuencia
    try:
        auditoria_vacios(df_master, AUDIT_DIR)
    except Exception as e:
        print(f"    [!] Falló la auditoría de vacíos: {e}")

    try:
        auditoria_cobertura_curricular(AUDIT_DIR)
    except Exception as e:
        print(f"    [!] Falló la auditoría de cobertura: {e}")

    try:
        # Se asume que el Centro Virtual tiene el código '99999'
        auditoria_representatividad(df_master, AUDIT_DIR, lista_escuelas=['99999'])
    except Exception as e:
        print(f"    [!] Falló la auditoría de representatividad: {e}")
        
    try:
        auditoria_frecuencia_opciones(AUDIT_DIR)
    except Exception as e:
        print(f"    [!] Falló la auditoría de frecuencia de opciones: {e}")

if __name__ == "__main__":
    # Para pruebas directas
    auditoria_calidad()