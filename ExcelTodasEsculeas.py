import os
import sys
import glob
import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ==========================================
# 1. CONFIGURACIÓN GENERAL Y RUTAS EXACTAS
# ==========================================
PATH_GEISER_R1 = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_Resultados_Febrero\Interim_CSVs\Resultados"
PATH_GEISER_M1 = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados"
PATH_GEISER_M2 = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"

# ARCHIVO OFICIAL (FUENTE DE LA VERDAD) PARA INYECTAR MEDIAS
ARCHIVO_OFICIAL_MEDIAS = "2_Evolucion_Medias_Progreso_2.csv"

DIR_SALIDA = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Final_Reports"
os.makedirs(DIR_SALIDA, exist_ok=True)

COLORES_CAT = {
    "Crítico": {"bg": "991B1B", "font": "FFFFFF"}, 
    "Bajo": {"bg": "FF8C2E", "font": "FFFFFF"},
    "Medio": {"bg": "FACC15", "font": "000000"}, 
    "Bueno": {"bg": "84CC16", "font": "000000"},
    "Excelente": {"bg": "065F46", "font": "FFFFFF"}, 
    "Sin Datos": {"bg": "E0E0E0", "font": "000000"}
}

# ==========================================
# 2. FUNCIONES DE BASE Y LIMPIEZA
# ==========================================
def extraer_grado_geiser(grado_str):
    if pd.isna(grado_str): return None
    match = re.search(r'\d+', str(grado_str))
    return int(match.group()) if match else None

def clasificar_puntaje(val):
    if pd.isna(val) or val == "": return "Sin Datos"
    try:
        media = float(val)
        if 0 <= media <= 35: return "Crítico"
        elif 35 < media <= 45: return "Bajo"
        elif 45 < media <= 55: return "Medio"
        elif 55 < media <= 65: return "Bueno"
        elif 65 < media <= 100: return "Excelente"
        else: return "Sin Datos"
    except:
        return "Sin Datos"

def es_vacio(x):
    if pd.isna(x): return True
    val = str(x).strip().lower()
    if val in ['', 'nan', 'null', '-', 'nr', 'n/a', 'na']: return True
    return False

# =========================================================================
# 3. EXTRACCIÓN DE DATOS CRUDOS DIRECTO DEL DISCO
# =========================================================================
def extraer_datos_crudos(ruta, mes_label):
    """Extrae TODOS los registros de la ruta (Para el motor de promedios)."""
    csv_files = glob.glob(os.path.join(ruta, "*.csv"))
    lista_dfs = []
    
    for f in csv_files:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            materia = 'Mat' if 'MAT' in os.path.basename(f).upper() else 'Len'
            
            col_cod = next((c for c in df.columns if 'nro de centro' in str(c).lower() or 'infra' in str(c).lower() or 'código' in str(c).lower() or 'centro' in str(c).lower()), None)
            col_centro = next((c for c in df.columns if 'centro' in str(c).lower() or 'institu' in str(c).lower() or 'escuela' in str(c).lower() and c != col_cod), None)
            col_nie = next((c for c in df.columns if 'documento' in c.lower() or 'nie' in c.lower()), None)
            col_nombre = next((c for c in df.columns if 'nombre' in c.lower()), None)
            col_apellido = next((c for c in df.columns if 'apellido' in c.lower()), None)
            col_grado = next((c for c in df.columns if 'grado' in c.lower()), None)
            col_grupo = next((c for c in df.columns if 'grupo' in c.lower() or 'sección' in c.lower()), None)
            col_theta = next((c for c in df.columns if 'escala 0-100' in str(c).lower()), 'theta.global (escala 0-100)')
            col_anular = next((c for c in df.columns if 'anular_prueba' in c.lower()), None)
            
            if not col_cod or col_theta not in df.columns: continue
            
            temp = pd.DataFrame()
            temp['codigo'] = df[col_cod].astype(str).str.strip().str.replace('.0', '', regex=False)
            temp['Centro_Nom'] = df[col_centro].astype(str).str.strip().str.upper() if col_centro else "DESCONOCIDO"
            temp['nie'] = df[col_nie].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() if col_nie else ""
            nom = df[col_nombre].fillna('').str.strip() if col_nombre else ""
            ape = df[col_apellido].fillna('').str.strip() if col_apellido else ""
            temp['Apellido+Nombre'] = ape + " " + nom
            temp['grado_num'] = df[col_grado].apply(extraer_grado_geiser) if col_grado else np.nan
            temp['Grupo'] = df[col_grupo].astype(str).str.strip() if col_grupo else ""
            temp['Puntaje'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            temp['Materia'] = materia
            temp['Mes'] = mes_label
            temp['Anulado'] = df[col_anular].apply(lambda x: False if es_vacio(x) or str(x).strip().lower() in ['0', 'false', 'falso', 'no'] else True) if col_anular else False
            
            lista_dfs.append(temp)
        except Exception: pass
        
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()

# =========================================================================
# 4. MOTOR DE INYECCIÓN (FUERZA BRUTA DEL CSV OFICIAL)
# =========================================================================
def inyectar_medias_oficiales(df_centros):
    """
    Busca el archivo CSV oficial y sobrescribe la Media 1 y Media 2.
    """
    if os.path.exists(ARCHIVO_OFICIAL_MEDIAS):
        print(f"   -> [INYECCIÓN] Archivo '{ARCHIVO_OFICIAL_MEDIAS}' detectado. Sincronizando promedios oficiales...")
        try:
            df_oficial = pd.read_csv(ARCHIVO_OFICIAL_MEDIAS, sep=';', dtype=str, encoding_errors='ignore')
            df_oficial.columns = df_oficial.columns.astype(str).str.strip()
            
            col_cod = next((c for c in df_oficial.columns if 'nro de centro' in c.lower() or 'código' in c.lower()), None)
            col_m1 = next((c for c in df_oficial.columns if c.lower() == 'media 1' or 'media 1' in c.lower()), None)
            col_m2 = next((c for c in df_oficial.columns if c.lower() == 'media 2' or 'media 2' in c.lower()), None)
            
            if col_cod:
                df_oficial[col_cod] = df_oficial[col_cod].astype(str).str.strip().str.replace('.0', '', regex=False)
                
                if col_m1 and 'Media 1 Cons' in df_centros.columns:
                    df_oficial[col_m1] = pd.to_numeric(df_oficial[col_m1].astype(str).str.replace(',', '.'), errors='coerce')
                    map_m1 = dict(zip(df_oficial[col_cod], df_oficial[col_m1]))
                    df_centros['Media 1 Cons'] = df_centros['Código de infraestructura'].map(map_m1).fillna(df_centros['Media 1 Cons'])
                    df_centros['Categoría 1 Cons'] = df_centros['Media 1 Cons'].apply(clasificar_puntaje)
                
                if col_m2 and 'Media 2 Cons' in df_centros.columns:
                    df_oficial[col_m2] = pd.to_numeric(df_oficial[col_m2].astype(str).str.replace(',', '.'), errors='coerce')
                    map_m2 = dict(zip(df_oficial[col_cod], df_oficial[col_m2]))
                    df_centros['Media 2 Cons'] = df_centros['Código de infraestructura'].map(map_m2).fillna(df_centros['Media 2 Cons'])
                    df_centros['Categoría 2 Cons'] = df_centros['Media 2 Cons'].apply(clasificar_puntaje)
                    
                print("      [+] Medias 1 y 2 oficiales inyectadas y sincronizadas exitosamente.")
        except Exception as e:
            print(f"      [!] Hubo un error al procesar el archivo oficial de inyección: {e}")
    else:
        print(f"      [-] No se detectó '{ARCHIVO_OFICIAL_MEDIAS}'. Se utilizará el algoritmo de contingencia.")
    
    return df_centros

# =========================================================================
# 5. CÁLCULO DEL ALGORITMO ORIGINAL (CONTINGENCIA / BASE)
# =========================================================================
def calcular_matriz_historica(df_crudo, index_cols):
    """
    Aplica el algoritmo de la "Bolsa Única" de generador_rankings.py.
    Agrupa absolutamente todos los puntajes por Centro o Sección y redondea a 1 decimal.
    """
    if df_crudo.empty: return pd.DataFrame()
    
    df_crudo_limpio = df_crudo.dropna(subset=['Puntaje', 'codigo'])
    
    # 1. Medias separadas por Materia
    df_matlen = df_crudo_limpio.pivot_table(index=index_cols, columns=['Mes', 'Materia'], values='Puntaje', aggfunc='mean').reset_index()
    df_matlen.columns = [' '.join(col).strip() if isinstance(col, tuple) else col for col in df_matlen.columns.values]
    
    # 2. Consolidado ("Bolsa Única" del Ranking)
    df_cons = df_crudo_limpio.groupby(index_cols + ['Mes'])['Puntaje'].mean().reset_index()
    df_cons['Puntaje'] = df_cons['Puntaje'].round(1) 
    
    df_cons_pivot = df_cons.pivot_table(index=index_cols, columns=['Mes'], values='Puntaje', aggfunc='first').reset_index()
    df_cons_pivot.columns = [str(c) + ' Consolidado' if c in ['R1', 'M1', 'M2'] else c for c in df_cons_pivot.columns.values]
    
    # 3. Unir Todo
    df_res = pd.merge(df_matlen, df_cons_pivot, on=index_cols, how='left')
    
    rep = pd.DataFrame()
    for col in index_cols:
        if col == 'codigo': rep['Código de infraestructura'] = df_res[col]
        elif col == 'Centro_Nom': rep['Centro'] = df_res[col]
        elif col == 'grado_num': rep['Grado'] = df_res[col].apply(lambda x: f"{int(x)}°" if pd.notna(x) else "")
        elif col == 'Grupo': rep['Sección'] = df_res[col]

    for prefix in ['R1', '1', '2']:
        m_prefix = prefix if prefix == 'R1' else 'M' + prefix
        
        rep[f'Media {prefix} Mat'] = df_res[f'{m_prefix} Mat'].round(1) if f'{m_prefix} Mat' in df_res.columns else np.nan
        rep[f'Categoría {prefix} Mat'] = rep[f'Media {prefix} Mat'].apply(clasificar_puntaje)
        
        rep[f'Media {prefix} Len'] = df_res[f'{m_prefix} Len'].round(1) if f'{m_prefix} Len' in df_res.columns else np.nan
        rep[f'Categoría {prefix} Len'] = rep[f'Media {prefix} Len'].apply(clasificar_puntaje)
        
        # El valor de contingencia (luego será sobreescrito por la inyección para Centros)
        rep[f'Media {prefix} Cons'] = df_res[f'{m_prefix} Consolidado'] if f'{m_prefix} Consolidado' in df_res.columns else np.nan
        rep[f'Categoría {prefix} Cons'] = rep[f'Media {prefix} Cons'].apply(clasificar_puntaje)
            
    return rep

# ==========================================
# 6. ORQUESTADOR Y GENERADOR DEL EXCEL INTEGRADO
# ==========================================
def generar_reporte_integrado():
    print("\n=======================================================================")
    print("  INICIANDO GENERADOR MAESTRO INTEGRADO (ALGORITMO RANKING EXACTO)")
    print("=======================================================================\n")

    print("[1/4] Leyendo bases de datos crudas desde el disco (sin filtros)...")
    df_r1 = extraer_datos_crudos(PATH_GEISER_R1, "R1")
    df_m1 = extraer_datos_crudos(PATH_GEISER_M1, "M1")
    df_m2 = extraer_datos_crudos(PATH_GEISER_M2, "M2")
    
    if df_r1.empty and df_m1.empty and df_m2.empty:
        print("\n[!] ERROR CRÍTICO: No se encontraron datos CSV en las rutas.")
        sys.exit()
        
    df_crudo_master = pd.concat([df_r1, df_m1, df_m2], ignore_index=True)
    
    escuelas_universo = df_r1['codigo'].unique()
    num_escuelas = len(escuelas_universo)
    
    if num_escuelas == 0:
        print("\n[!] No se encontraron escuelas en los archivos de Febrero (R1).")
        sys.exit()
        
    print(f"      -> {num_escuelas} escuelas detectadas en el universo (Febrero).")
    
    # Filtramos todo al universo de escuelas de Febrero
    df_master = df_crudo_master[df_crudo_master['codigo'].isin(escuelas_universo)].copy()

    # -------------------------------------------------------------
    # HOJA 1: ESTUDIANTES (Aplicando Filtros Estrictos para limpieza visual)
    # -------------------------------------------------------------
    print("\n[2/4] Generando matriz individual de Estudiantes (Filtro Limpio)...")
    df_estudiantes_limpios = df_master[
        (df_master['Anulado'] == False) & 
        (df_master['grado_num'].notna()) &
        (df_master['grado_num'] >= 2) & 
        (df_master['grado_num'] <= 11) &
        (df_master['nie'] != "") &
        (df_master['Puntaje'].notna())
    ].copy()
    
    df_p = df_estudiantes_limpios.pivot_table(index=['codigo', 'Centro_Nom', 'grado_num', 'Grupo', 'nie', 'Apellido+Nombre'], columns=['Mes', 'Materia'], values='Puntaje', aggfunc='first').reset_index()
    df_p.columns = [' '.join(col).strip() for col in df_p.columns.values]
    
    for col in ['R1 Mat', 'M1 Mat', 'M2 Mat', 'R1 Len', 'M1 Len', 'M2 Len']:
        if col not in df_p.columns: df_p[col] = np.nan

    rep_estudiantes = pd.DataFrame()
    rep_estudiantes['Código de infraestructura'] = df_p['codigo']
    rep_estudiantes['Centro'] = df_p['Centro_Nom']
    rep_estudiantes['Grado'] = df_p['grado_num'].apply(lambda x: f"{int(x)}°" if pd.notna(x) else "")
    rep_estudiantes['Sección'] = df_p['Grupo']
    rep_estudiantes['NIE'] = df_p['nie']
    rep_estudiantes['Apellido+Nombre'] = df_p['Apellido+Nombre']
    
    rep_estudiantes['Puntaje R1 Mat'] = df_p['R1 Mat'].round(2); rep_estudiantes['Categoría R1 Mat'] = rep_estudiantes['Puntaje R1 Mat'].apply(clasificar_puntaje)
    rep_estudiantes['Puntaje 1 Mat'] = df_p['M1 Mat'].round(2); rep_estudiantes['Categoría 1 Mat'] = rep_estudiantes['Puntaje 1 Mat'].apply(clasificar_puntaje)
    rep_estudiantes['Puntaje 2 Mat'] = df_p['M2 Mat'].round(2); rep_estudiantes['Categoría 2 Mat'] = rep_estudiantes['Puntaje 2 Mat'].apply(clasificar_puntaje)
    rep_estudiantes['Puntaje R1 Len'] = df_p['R1 Len'].round(2); rep_estudiantes['Categoría R1 Len'] = rep_estudiantes['Puntaje R1 Len'].apply(clasificar_puntaje)
    rep_estudiantes['Puntaje 1 Len'] = df_p['M1 Len'].round(2); rep_estudiantes['Categoría 1 Len'] = rep_estudiantes['Puntaje 1 Len'].apply(clasificar_puntaje)
    rep_estudiantes['Puntaje 2 Len'] = df_p['M2 Len'].round(2); rep_estudiantes['Categoría 2 Len'] = rep_estudiantes['Puntaje 2 Len'].apply(clasificar_puntaje)
    
    for col in ['Puntaje R1 Mat', 'Puntaje 1 Mat', 'Puntaje 2 Mat', 'Puntaje R1 Len', 'Puntaje 1 Len', 'Puntaje 2 Len']: rep_estudiantes[col] = rep_estudiantes[col].fillna("")

    # -------------------------------------------------------------
    # HOJA 2: CENTROS (Datos Crudos + INYECCIÓN CSV OFICIAL)
    # -------------------------------------------------------------
    print("\n[3/4] Generando promedios por Centro y ejecutando inyección oficial...")
    rep_centros = calcular_matriz_historica(df_master, ['codigo', 'Centro_Nom'])
    rep_centros = inyectar_medias_oficiales(rep_centros)
    
    for c in rep_centros.columns:
        if 'Media' in c: rep_centros[c] = rep_centros[c].fillna("")

    # -------------------------------------------------------------
    # HOJA 3: SECCIONES (Datos Crudos, algoritmo histórico)
    # -------------------------------------------------------------
    print("\n[4/4] Generando promedios por Sección...")
    rep_secciones = calcular_matriz_historica(df_master, ['codigo', 'Centro_Nom', 'grado_num', 'Grupo'])
    
    for c in rep_secciones.columns:
        if 'Media' in c: rep_secciones[c] = rep_secciones[c].fillna("")

    # -------------------------------------------------------------
    # ESCRITURA DEL ARCHIVO INTEGRADO FINAL
    # -------------------------------------------------------------
    ruta_final = os.path.join(DIR_SALIDA, f"Reporte_General_Integrado_{num_escuelas}_Escuelas.xlsx")
    print(f"\n[*] Ensamblando y aplicando estilos al archivo maestro: {os.path.basename(ruta_final)} ...")
    
    with pd.ExcelWriter(ruta_final, engine='openpyxl') as writer:
        # Pestaña 1
        rep_estudiantes.to_excel(writer, index=False, sheet_name='Resultados por Estudiante')
        ws1 = writer.book['Resultados por Estudiante']
        cols_cat1 = [rep_estudiantes.columns.get_loc(c) + 1 for c in ['Categoría R1 Mat', 'Categoría 1 Mat', 'Categoría 2 Mat', 'Categoría R1 Len', 'Categoría 1 Len', 'Categoría 2 Len']]
        for row in range(2, len(rep_estudiantes) + 2):
            for col_idx in cols_cat1:
                val = ws1.cell(row=row, column=col_idx).value
                if val in COLORES_CAT:
                    ws1.cell(row=row, column=col_idx).fill = PatternFill(start_color=COLORES_CAT[val]['bg'], end_color=COLORES_CAT[val]['bg'], fill_type="solid")
                    ws1.cell(row=row, column=col_idx).font = Font(color=COLORES_CAT[val]['font'], bold=True)
                    
        # Pestaña 2
        rep_centros.to_excel(writer, index=False, sheet_name='Promedios por Centro')
        ws2 = writer.book['Promedios por Centro']
        cols_cat2 = [i+1 for i, c in enumerate(rep_centros.columns) if 'Categoría' in c]
        for row in range(2, len(rep_centros) + 2):
            for col_idx in cols_cat2:
                val = ws2.cell(row=row, column=col_idx).value
                if val in COLORES_CAT:
                    ws2.cell(row=row, column=col_idx).fill = PatternFill(start_color=COLORES_CAT[val]['bg'], end_color=COLORES_CAT[val]['bg'], fill_type="solid")
                    ws2.cell(row=row, column=col_idx).font = Font(color=COLORES_CAT[val]['font'], bold=True)
                    
        # Pestaña 3
        rep_secciones.to_excel(writer, index=False, sheet_name='Promedios por Sección')
        ws3 = writer.book['Promedios por Sección']
        cols_cat3 = [i+1 for i, c in enumerate(rep_secciones.columns) if 'Categoría' in c]
        for row in range(2, len(rep_secciones) + 2):
            for col_idx in cols_cat3:
                val = ws3.cell(row=row, column=col_idx).value
                if val in COLORES_CAT:
                    ws3.cell(row=row, column=col_idx).fill = PatternFill(start_color=COLORES_CAT[val]['bg'], end_color=COLORES_CAT[val]['bg'], fill_type="solid")
                    ws3.cell(row=row, column=col_idx).font = Font(color=COLORES_CAT[val]['font'], bold=True)

    print("\n=======================================================================")
    print("  [OK] ¡PROCESO COMPLETADO! Archivo maestro disponible en Final_Reports.")
    print("=======================================================================\n")

if __name__ == "__main__":
    generar_reporte_integrado()