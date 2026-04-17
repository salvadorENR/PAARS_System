import os
import re
import glob
import pandas as pd
import numpy as np

# ==========================================
# CONFIGURACIÓN DE RUTAS
# ==========================================
TXT_ROOT_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026"
METADATA_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata"

def clean_documento(doc):
    if pd.isna(doc): return ""
    return str(doc).strip().replace('-', '').replace('.0', '')

def clean_grade(g):
    if pd.isna(g): return ""
    match = re.search(r'(\d+)', str(g))
    if match: return f"{match.group(1)}°"
    return str(g).strip()

def clean_option(x):
    """
    FILTRO INVENCIBLE: Extrae pura y estrictamente la letra de la opción,
    ignorando puntos, espacios, paréntesis o números basura (ej. "A (1556263)").
    """
    if pd.isna(x): return 'N/R'
    
    # Destruye todo lo que no sea letra o número
    val = re.sub(r'[^A-Za-z0-9]', '', str(x)).upper()
    
    # Si quedó vacío o es NR (No responde)
    if not val or val in ['NR', 'N']: 
        return 'N/R'
        
    # Toma estrictamente el primer caracter válido
    letra = val[0] 
    
    if letra in ['A', 'B', 'C', 'D', 'E', 'V', 'F']:
        return letra
        
    # Si la plataforma usó números en lugar de letras
    if letra == '1': return 'A'
    if letra == '2': return 'B'
    if letra == '3': return 'C'
    if letra == '4': return 'D'
    if letra == '5': return 'E'
    
    return 'N/R'

def get_col_name(df, posibles_nombres):
    for col in df.columns:
        if str(col).strip().lower() in posibles_nombres:
            return col
    return None

def get_item_number(item_code):
    match = re.search(r'\d+', str(item_code))
    return int(match.group()) if match else 0

def calcular_limites_sote(group):
    tiempos = group['Tiempo total de prueba'].dropna()
    if tiempos.empty:
        group['Clasificación SOTE'] = 'Sin tiempo'
        return group
        
    q3 = tiempos.quantile(0.75)
    q1 = tiempos.quantile(0.25)
    iqr = q3 - q1
    limite_maximo = q3 + 3 * iqr
    
    group['Clasificación SOTE'] = group['Tiempo total de prueba'].apply(
        lambda x: 'Confiable' if (pd.notna(x) and 5 <= x <= limite_maximo) else 'No confiable'
    )
    return group

def auditoria_calidad():
    print("======================================================")
    print("   AUDITORÍA TOTAL DE CALIDAD (QA) - PAARS")
    print("======================================================\n")

    print(f"[*] Escaneando exámenes en el disco:\n    {TXT_ROOT_DIR}...")
    patron_progreso = os.path.join(TXT_ROOT_DIR, "*", "Raw_Data", "Reportes Prueba Progreso * letras", "*.txt")
    patron_resultados = os.path.join(TXT_ROOT_DIR, "*", "Raw_Data", "Reportes Prueba Resultados * letras", "*.txt")
    
    todos_los_txt = glob.glob(patron_progreso) + glob.glob(patron_resultados)
    
    examenes = {}
    for f in todos_los_txt:
        raw_data_dir = os.path.dirname(os.path.dirname(f))
        exam_folder_path = os.path.dirname(raw_data_dir)
        report_folder = os.path.basename(os.path.dirname(f))
        match = re.search(r'Reportes Prueba (Progreso|Resultados?) (\d+) letras', report_folder, re.IGNORECASE)
        
        if match:
            tipo = match.group(1).capitalize()
            if tipo == 'Resultado': tipo = 'Resultados'
            numero = match.group(2)
            clave_examen = (exam_folder_path, raw_data_dir, tipo, numero)
            if clave_examen not in examenes: examenes[clave_examen] = []
            examenes[clave_examen].append(f)

    for (exam_folder_path, raw_data_dir, tipo, numero), archivos_txt in examenes.items():
        nombre_examen = f"{tipo} {numero}"
        print(f"\n------------------------------------------------------")
        print(f"[*] Auditando: {nombre_examen} ({os.path.basename(exam_folder_path)})")

        # -------------------------------------------------------------
        # 1. CARGAR DATOS SOTE (TXT)
        # -------------------------------------------------------------
        lista_sote = []
        for file in archivos_txt:
            try:
                try: df = pd.read_csv(file, sep='|', quotechar='"', dtype=str, encoding='utf-8')
                except: df = pd.read_csv(file, sep='|', quotechar='"', dtype=str, encoding='latin-1')
                df.columns = df.columns.str.strip().str.replace('"', '')
                lista_sote.append(df)
            except: pass
        
        if not lista_sote: continue
        df_sote = pd.concat(lista_sote, ignore_index=True)
        
        df_sote['Documento_Limpio'] = df_sote['Documento'].apply(clean_documento)
        df_sote = df_sote[df_sote['Documento_Limpio'] != ""]
        df_sote['Grado_Limpio'] = df_sote['Grado'].apply(clean_grade)
        
        if 'Area temática' not in df_sote.columns: df_sote['Area temática'] = 'Desconocida'
        df_sote['Materia'] = df_sote['Area temática'].apply(lambda x: 'MAT' if 'mat' in str(x).lower() else ('LEC' if 'lec' in str(x).lower() or 'leng' in str(x).lower() else 'UNK'))
        
        col_inicio = get_col_name(df_sote, ['fecha-hora de inicio'])
        col_fin = get_col_name(df_sote, ['fecha-hora de fin'])
        
        if col_inicio and col_fin:
            df_sote['Inicio'] = pd.to_datetime(df_sote[col_inicio], errors='coerce')
            df_sote['Fin'] = pd.to_datetime(df_sote[col_fin], errors='coerce')
            df_sote['Tiempo total de prueba'] = (df_sote['Fin'] - df_sote['Inicio']).dt.total_seconds() / 60.0
            df_sote = df_sote.groupby(['Grado_Limpio', 'Materia'], group_keys=False).apply(calcular_limites_sote)
        else:
            df_sote['Tiempo total de prueba'] = np.nan
            df_sote['Clasificación SOTE'] = 'Faltan datos de hora'

        # -------------------------------------------------------------
        # 2. CARGAR DATOS GEISER (EXCEL MULTI-PESTAÑA)
        # -------------------------------------------------------------
        lista_geiser = []
        mapeo_archivos = [("MAT", "MAT-resultados.xlsx"), ("LEC", "LEC-resultados.xlsx")]
        
        for materia_fija, nombre_fijo in mapeo_archivos:
            archivos_encontrados = glob.glob(os.path.join(raw_data_dir, nombre_fijo))
            archivos_encontrados = [f for f in archivos_encontrados if not os.path.basename(f).startswith('~')]
            
            for file in archivos_encontrados:
                try:
                    xls = pd.read_excel(file, sheet_name=None, dtype=str)
                    for sheet_name, df in xls.items():
                        df['Materia'] = materia_fija
                        if get_col_name(df, ['grado', 'grado_aplicado']) is None:
                            df['Grado_Inferido'] = sheet_name
                        lista_geiser.append(df)
                except Exception as e: pass
        
        if not lista_geiser: continue
        df_geiser = pd.concat(lista_geiser, ignore_index=True)
        
        col_doc_geiser = get_col_name(df_geiser, ['documento', 'nie'])
        col_grado_geiser = get_col_name(df_geiser, ['grado', 'grado_aplicado', 'grado_inferido'])
        
        df_geiser['Documento_Limpio'] = df_geiser[col_doc_geiser].apply(clean_documento) if col_doc_geiser else ""
        df_geiser = df_geiser[df_geiser['Documento_Limpio'] != ""]
        df_geiser['Grado_Limpio'] = df_geiser[col_grado_geiser].apply(clean_grade) if col_grado_geiser else ""

        col_anular = get_col_name(df_geiser, ['anular_prueba', 'anular prueba'])
        col_duracion = get_col_name(df_geiser, ['duration', 'duración', 'duracion'])
        
        if col_anular:
            df_geiser['Clasificación archivos TRI'] = df_geiser[col_anular].apply(
                lambda x: 'Confiable' if pd.isna(x) or str(x).strip() == '' else 'No confiable'
            )
        else:
            df_geiser['Clasificación archivos TRI'] = 'No encontrada'

        # -------------------------------------------------------------
        # 3. REPORTES 1, 2, 3 (Comparativas Globales)
        # -------------------------------------------------------------
        out_dir = os.path.join(exam_folder_path, "Final_Reports", "Quality_Control")
        os.makedirs(out_dir, exist_ok=True)
        print("  -> Generando reportes de auditoría de tiempos y registros globales...")

        conteo_sote = df_sote.groupby(['Grado_Limpio', 'Materia']).size().reset_index(name='Total_SOTE_TXT')
        conteo_geiser = df_geiser.groupby(['Grado_Limpio', 'Materia']).size().reset_index(name='Total_Geiser_Excel')
        comp_registros = pd.merge(conteo_sote, conteo_geiser, on=['Grado_Limpio', 'Materia'], how='outer').fillna(0)
        comp_registros['Diferencia (Geiser - SOTE)'] = comp_registros['Total_Geiser_Excel'] - comp_registros['Total_SOTE_TXT']
        
        if not comp_registros.empty:
            comp_registros['sort_grado'] = comp_registros['Grado_Limpio'].apply(get_item_number)
            comp_registros = comp_registros.sort_values(by=['sort_grado', 'Materia']).drop(columns=['sort_grado'])

        totales = pd.DataFrame([{'Grado_Limpio': 'GRAN TOTAL', 'Materia': '-', 'Total_SOTE_TXT': comp_registros['Total_SOTE_TXT'].sum(), 'Total_Geiser_Excel': comp_registros['Total_Geiser_Excel'].sum(), 'Diferencia (Geiser - SOTE)': comp_registros['Diferencia (Geiser - SOTE)'].sum()}])
        comp_registros = pd.concat([comp_registros, totales], ignore_index=True)

        items_sote = [c for c in df_sote.columns if re.match(r'^(MAT|LEC)\d+', c.strip().upper())]
        items_geiser = [c for c in df_geiser.columns if re.match(r'^(MAT|LEC)\d+', c.strip().upper())]
        
        sote_items_by_grade = []
        for grado, group in df_sote.groupby('Grado_Limpio'):
            if not grado: continue
            for col in items_sote:
                valid_vals = group[col].dropna().astype(str).str.strip()
                if (valid_vals != "").any():
                    materia = 'MAT' if col.upper().startswith('MAT') else 'LEC'
                    sote_items_by_grade.append({'Grado': grado, 'Materia': materia, 'Código del Ítem': col.upper(), 'SOTE (TXT FILES)': 'x'})

        geiser_items_by_grade = []
        for grado, group in df_geiser.groupby('Grado_Limpio'):
            if not grado: continue
            for col in items_geiser:
                valid_vals = group[col].dropna().astype(str).str.strip()
                if (valid_vals != "").any():
                    materia = 'MAT' if col.upper().startswith('MAT') else 'LEC'
                    geiser_items_by_grade.append({'Grado': grado, 'Materia': materia, 'Código del Ítem': col.upper(), 'TRI (EXCEL FILES)': 'x'})

        df_items_sote = pd.DataFrame(sote_items_by_grade) if sote_items_by_grade else pd.DataFrame(columns=['Grado', 'Materia', 'Código del Ítem', 'SOTE (TXT FILES)'])
        df_items_geiser = pd.DataFrame(geiser_items_by_grade) if geiser_items_by_grade else pd.DataFrame(columns=['Grado', 'Materia', 'Código del Ítem', 'TRI (EXCEL FILES)'])

        df_items_comp = pd.merge(df_items_sote, df_items_geiser, on=['Grado', 'Materia', 'Código del Ítem'], how='outer').fillna('-')
        
        resumen_comunes_lista = []
        if not df_items_comp.empty:
            for (grado, materia), group in df_items_comp.groupby(['Grado', 'Materia']):
                comunes = group[(group['SOTE (TXT FILES)'] == 'x') & (group['TRI (EXCEL FILES)'] == 'x')]['Código del Ítem'].tolist()
                no_comunes = group[(group['SOTE (TXT FILES)'] == '-') | (group['TRI (EXCEL FILES)'] == '-')]['Código del Ítem'].tolist()
                
                comunes.sort(key=get_item_number)
                no_comunes.sort(key=get_item_number)
                
                resumen_comunes_lista.append({
                    'Examen (Grado-Subject)': f"{grado} - {materia}",
                    'Número de ítems comunes': len(comunes),
                    'Ítems comunes': ", ".join(comunes) if comunes else "Ninguno",
                    'Número de ítems no comunes': len(no_comunes),
                    'Ítems no comunes': ", ".join(no_comunes) if no_comunes else "Ninguno"
                })
                
        df_resumen_comunes = pd.DataFrame(resumen_comunes_lista)
        if not df_resumen_comunes.empty:
            df_resumen_comunes['sort_grado'] = df_resumen_comunes['Examen (Grado-Subject)'].apply(get_item_number)
            df_resumen_comunes = df_resumen_comunes.sort_values(by=['sort_grado', 'Examen (Grado-Subject)']).drop(columns=['sort_grado'])

        if not df_items_comp.empty:
            df_items_comp['sort_grado'] = df_items_comp['Grado'].apply(get_item_number)
            df_items_comp['sort_item'] = df_items_comp['Código del Ítem'].apply(get_item_number)
            df_items_comp = df_items_comp.sort_values(['sort_grado', 'sort_item']).drop(columns=['sort_grado', 'sort_item'])

        df_items_mat = df_items_comp[df_items_comp['Materia'] == 'MAT'].drop(columns=['Materia']) if not df_items_comp.empty else pd.DataFrame()
        df_items_lec = df_items_comp[df_items_comp['Materia'] == 'LEC'].drop(columns=['Materia']) if not df_items_comp.empty else pd.DataFrame()

        r1_name = f'1_QC_Global_Comparison_{tipo}_{numero}.xlsx'
        try:
            with pd.ExcelWriter(os.path.join(out_dir, r1_name)) as writer:
                comp_registros.to_excel(writer, sheet_name='Comparacion_Registros', index=False)
                if not df_resumen_comunes.empty:
                    df_resumen_comunes.to_excel(writer, sheet_name='Resumen_Items_Comunes', index=False)
                if not df_items_mat.empty:
                    df_items_mat.to_excel(writer, sheet_name='Items_Matematica', index=False)
                if not df_items_lec.empty:
                    df_items_lec.to_excel(writer, sheet_name='Items_Lengua', index=False)
            print(f"  [OK] Creado: {r1_name}")
        except PermissionError:
            print(f"  [!] ERROR DE PERMISO: El archivo {r1_name} está abierto en Excel. Ciérralo.")

        # COMPARATIVA DE TIEMPOS
        columnas_sote_merge = ['Documento_Limpio', 'Materia', 'Nombre', 'Apellido', 'Nro de centro', 'Centro', 'Grado_Limpio', 'Grupo', col_inicio, col_fin, 'Tiempo total de prueba', 'Clasificación SOTE']
        columnas_sote_merge = [c for c in columnas_sote_merge if c in df_sote.columns]
        columnas_geiser_merge = ['Documento_Limpio', 'Materia', col_duracion, col_anular, 'Clasificación archivos TRI']
        columnas_geiser_merge = [c for c in columnas_geiser_merge if c in df_geiser.columns]

        df_cruce = pd.merge(df_sote[columnas_sote_merge], df_geiser[columnas_geiser_merge], on=['Documento_Limpio', 'Materia'], how='inner')
        df_cruce['Coincide SOTE y TRI'] = np.where(df_cruce['Clasificación SOTE'] == df_cruce['Clasificación archivos TRI'], 'Coincide', 'No coincide')
        
        renombres = {col_inicio: 'Fecha-Hora de Inicio (TXT)', col_fin: 'Fecha-Hora de Fin (TXT)', col_duracion: 'Duración (Geiser)', col_anular: 'Anular Prueba (Geiser)', 'Documento_Limpio': 'Documento', 'Grado_Limpio': 'Grado'}
        df_cruce.rename(columns=renombres, inplace=True)

        r2_name = f'2_QC_Student_Time_{tipo}_{numero}.xlsx'
        try:
            with pd.ExcelWriter(os.path.join(out_dir, r2_name), engine='openpyxl') as writer:
                for grado in df_cruce['Grado'].unique():
                    if pd.isna(grado) or grado == "": continue
                    df_grado = df_cruce[df_cruce['Grado'] == grado]
                    df_grado.to_excel(writer, sheet_name=str(grado), index=False)
            print(f"  [OK] Creado: {r2_name}")
        except PermissionError: pass

        resumen_lista = []
        for (grado, materia), group in df_cruce.groupby(['Grado', 'Materia']):
            total_comunes = len(group)
            coinciden = len(group[group['Coincide SOTE y TRI'] == 'Coincide'])
            no_coinciden = len(group[group['Coincide SOTE y TRI'] == 'No coincide'])
            porcentaje = (coinciden / total_comunes) * 100 if total_comunes > 0 else 0
            resumen_lista.append({'Grado': grado, 'Materia': materia, 'Total Estudiantes Comunes': total_comunes, 'Coinciden SOTE y TRI': coinciden, 'No Coinciden': no_coinciden, '% de Coincidencia': f"{porcentaje:.1f}%"})
            
        r3_name = f'3_QC_Summary_Coincidences_{tipo}_{numero}.xlsx'
        df_resumen_3 = pd.DataFrame(resumen_lista)
        
        # ¡ORDENAMIENTO DE 2° A 11° PARA REPORTE 3!
        if not df_resumen_3.empty:
            df_resumen_3['sort_grado'] = df_resumen_3['Grado'].apply(get_item_number)
            df_resumen_3 = df_resumen_3.sort_values(by=['sort_grado', 'Materia']).drop(columns=['sort_grado'])

        try:
            df_resumen_3.to_excel(os.path.join(out_dir, r3_name), index=False)
            print(f"  [OK] Creado: {r3_name}")
        except PermissionError: pass

        # -------------------------------------------------------------
        # 4. AUDITORÍA A NIVEL DE ÍTEM 
        # -------------------------------------------------------------
        print("  -> Generando Reportes de Calificación por Ítem (Auditoría Geiser vs SOTE)...")
        
        archivos_metadata = glob.glob(os.path.join(METADATA_DIR, "*.csv"))
        archivo_maestro = None
        for arch in archivos_metadata:
            nombre_arch = os.path.basename(arch).lower()
            if tipo.lower() in nombre_arch:
                nombre_sin_ano = re.sub(r'20\d\d', '', nombre_arch)
                if f"mes{numero}" in nombre_sin_ano or f"mes {numero}" in nombre_sin_ano or f"_{numero}_" in nombre_sin_ano:
                    archivo_maestro = arch
                    break

        if not archivo_maestro:
            print(f"  [!] ADVERTENCIA: No se encontró el archivo de claves. Omitiendo validación por ítem.")
            continue

        try:
            try: df_claves = pd.read_csv(archivo_maestro, sep=';', encoding='utf-8-sig', dtype=str)
            except: df_claves = pd.read_csv(archivo_maestro, sep=';', encoding='latin-1', dtype=str)
            if 'ItemCodigo' not in df_claves.columns:
                try: df_claves = pd.read_csv(archivo_maestro, sep=',', encoding='utf-8-sig', dtype=str)
                except: df_claves = pd.read_csv(archivo_maestro, sep=',', encoding='latin-1', dtype=str)

            col_g_clave = get_col_name(df_claves, ['grado_aplicado', 'pruebatitulo', 'grado'])
            df_claves['Grado_Limpio'] = df_claves[col_g_clave].apply(clean_grade) if col_g_clave else ""
            
            df_claves['Opcion Correcta'] = df_claves['Opcion Correcta'].apply(clean_option)
            df_claves['ItemCodigo'] = df_claves['ItemCodigo'].astype(str).str.strip().str.upper()
            df_claves = df_claves[['ItemCodigo', 'Grado_Limpio', 'Opcion Correcta']].dropna()
        except Exception as e:
            continue

        df_eval_list = []
        grados_disponibles_en_claves = df_claves['Grado_Limpio'].unique()

        for grado_val in grados_disponibles_en_claves:
            if not grado_val: continue
            
            for materia_val in ['MAT', 'LEC']:
                claves_g_m = df_claves[(df_claves['Grado_Limpio'] == grado_val) & (df_claves['ItemCodigo'].str.startswith(materia_val))]
                if claves_g_m.empty: continue
                
                items_permitidos = set(claves_g_m['ItemCodigo'])
                
                sote_g_m = df_sote[(df_sote['Grado_Limpio'] == grado_val) & (df_sote['Materia'] == materia_val)].copy()
                geiser_g_m = df_geiser[(df_geiser['Grado_Limpio'] == grado_val) & (df_geiser['Materia'] == materia_val)].copy()
                
                if sote_g_m.empty or geiser_g_m.empty: continue
                
                sote_cols = [c for c in sote_g_m.columns if c.strip().upper() in items_permitidos]
                geiser_cols = [c for c in geiser_g_m.columns if c.strip().upper() in items_permitidos]
                
                if not sote_cols or not geiser_cols: continue
                
                sote_melt = sote_g_m.melt(id_vars=['Documento_Limpio'], value_vars=sote_cols, var_name='ItemCodigo', value_name='Respuesta_TXT')
                sote_melt['ItemCodigo'] = sote_melt['ItemCodigo'].str.strip().str.upper()
                sote_melt = sote_melt.dropna(subset=['Respuesta_TXT'])
                sote_melt['Respuesta_TXT'] = sote_melt['Respuesta_TXT'].apply(clean_option)
                
                for col_g in geiser_cols:
                    geiser_g_m[col_g] = geiser_g_m[col_g].astype(str)
                    
                geiser_melt = geiser_g_m.melt(id_vars=['Documento_Limpio'], value_vars=geiser_cols, var_name='ItemCodigo', value_name='Puntaje_Geiser')
                geiser_melt['ItemCodigo'] = geiser_melt['ItemCodigo'].str.strip().str.upper()
                geiser_melt = geiser_melt.dropna(subset=['Puntaje_Geiser'])
                
                def clean_geiser_score(x):
                    if pd.isna(x): return '0'
                    val = str(x).strip().replace('.0', '').replace('nan', '')
                    return '1' if val == '1' else '0'
                geiser_melt['Puntaje_Geiser'] = geiser_melt['Puntaje_Geiser'].apply(clean_geiser_score)
                
                df_eval_g_m = pd.merge(sote_melt, geiser_melt, on=['Documento_Limpio', 'ItemCodigo'], how='inner')
                df_eval_g_m = pd.merge(df_eval_g_m, claves_g_m[['ItemCodigo', 'Opcion Correcta']], on=['ItemCodigo'], how='inner')
                
                df_eval_g_m['Grado_Limpio'] = grado_val
                df_eval_g_m['Materia'] = materia_val
                df_eval_list.append(df_eval_g_m)

        if not df_eval_list: continue

        df_eval = pd.concat(df_eval_list, ignore_index=True)
        
        df_eval['Puntaje_Esperado'] = np.where(df_eval['Respuesta_TXT'] == df_eval['Opcion Correcta'], '1', '0')
        df_eval['Validacion'] = np.where(df_eval['Puntaje_Esperado'] == df_eval['Puntaje_Geiser'], 'Coincide', 'No coincide')
        df_eval['Resultado_Matriz'] = np.where(df_eval['Validacion'] == 'Coincide', 'Correcto', 'Incorrecto')

        def guardar_matriz(df_sub, filename):
            if df_sub.empty: return
            try:
                with pd.ExcelWriter(os.path.join(out_dir, filename), engine='openpyxl') as writer:
                    for g_matriz in df_sub['Grado_Limpio'].unique():
                        if not g_matriz: continue
                        d_g = df_sub[df_sub['Grado_Limpio'] == g_matriz]
                        pivot = d_g.pivot_table(index='Documento_Limpio', columns='ItemCodigo', values='Resultado_Matriz', aggfunc='first').reset_index()
                        pivot.rename(columns={'Documento_Limpio': 'Documento'}, inplace=True)
                        pivot.to_excel(writer, sheet_name=str(g_matriz), index=False)
                print(f"  [OK] Creado: {filename}")
            except PermissionError:
                print(f"  [!] ERROR DE PERMISO: El archivo {filename} está abierto en Excel. Ciérralo.")

        guardar_matriz(df_eval[df_eval['Materia'] == 'MAT'], f'4_QC_Item_Scoring_Math_{tipo}_{numero}.xlsx')
        guardar_matriz(df_eval[df_eval['Materia'] == 'LEC'], f'5_QC_Item_Scoring_Lengua_{tipo}_{numero}.xlsx')

        resumen_scoring = df_eval.groupby(['Grado_Limpio', 'Materia', 'Validacion']).size().unstack(fill_value=0).reset_index()
        for col in ['Coincide', 'No coincide']:
            if col not in resumen_scoring: resumen_scoring[col] = 0

        resumen_scoring['Total Evaluaciones'] = resumen_scoring['Coincide'] + resumen_scoring['No coincide']
        resumen_scoring['% Coincidencia'] = (resumen_scoring['Coincide'] / resumen_scoring['Total Evaluaciones'] * 100).round(2).astype(str) + '%'
        resumen_scoring.rename(columns={'Grado_Limpio': 'Grado'}, inplace=True)
        
        # ¡ORDENAMIENTO DE 2° A 11° PARA REPORTE 6!
        if not resumen_scoring.empty:
            resumen_scoring['sort_grado'] = resumen_scoring['Grado'].apply(get_item_number)
            resumen_scoring = resumen_scoring.sort_values(by=['sort_grado', 'Materia']).drop(columns=['sort_grado'])

        r6_name = f'6_QC_Summary_Item_Scoring_{tipo}_{numero}.xlsx'
        try:
            resumen_scoring.to_excel(os.path.join(out_dir, r6_name), index=False)
            print(f"  [OK] Creado: {r6_name}")
        except PermissionError:
            print(f"  [!] ERROR DE PERMISO: El archivo {r6_name} está abierto en Excel.")

    print("\n¡Auditoría Total completada con éxito!")

if __name__ == "__main__":
    auditoria_calidad()