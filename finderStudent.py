import os
import sys
import glob
import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS Y ARCHIVOS
# ==========================================
PATH_MATRICULA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_AUTOGENERADOS = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\registros autogenerados nie-2026-05-14.csv"
PATH_RESULTADOS = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados"
DIR_SALIDA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Final_Reports"

os.makedirs(DIR_SALIDA, exist_ok=True)

# ==========================================
# 2. FUNCIONES DE LIMPIEZA Y NORMALIZACIÓN
# ==========================================
def normalizar_y_formatear_grado(grado_str):
    """Extrae el número del grado y le agrega el símbolo °. Ignora Tercer Año."""
    if pd.isna(grado_str): return None
    texto = str(grado_str).lower().strip()
    
    if "tercer año" in texto or "3er año" in texto:
        return None 
        
    num = None
    if "primer año" in texto or "1er año" in texto: num = 10
    elif "segundo año" in texto or "2do año" in texto: num = 11
    else:
        match = re.search(r'\d+', texto)
        if match: 
            num = int(match.group())
        else:
            palabras = {"segundo": 2, "tercer": 3, "cuarto": 4, "quinto": 5, "sexto": 6, "séptimo": 7, "septimo": 7, "octavo": 8, "noveno": 9}
            for p, v in palabras.items():
                if p in texto: 
                    num = v
                    break
                    
    return f"{num}°" if num else None

def limpiar_nombre(ape1, ape2, nom1, nom2):
    """Concatena los nombres y apellidos omitiendo los valores nulos."""
    partes = [ape1, ape2, nom1, nom2]
    partes_limpias = [str(p).strip() for p in partes if pd.notna(p) and str(p).strip() != ""]
    return " ".join(partes_limpias)

def normalizar_texto(texto):
    """Elimina acentos, caracteres especiales y espacios múltiples para búsquedas robustas por nombre."""
    if pd.isna(texto): return ""
    t = str(texto).lower().strip()
    t = re.sub(r'[áäâà]', 'a', t)
    t = re.sub(r'[éëêè]', 'e', t)
    t = re.sub(r'[íïîì]', 'i', t)
    t = re.sub(r'[óöôò]', 'o', t)
    t = re.sub(r'[úüûù]', 'u', t)
    t = re.sub(r'\s+', ' ', t)
    return t

# ==========================================
# 3. MOTOR PRINCIPAL
# ==========================================
def generar_reporte_nominal():
    print(f"\n[*] INICIANDO BÚSQUEDA EXPANDIDA CON INFORMACIÓN DE CENTROS EDUCATIVOS...")

    target_nies = ["10103662", "10083400"]
    target_names_norm = [
        "jonathan enmanuel cabrera carranza",
        "kimberly anahy ardon vasquez"
    ]
    
    def cumple_criterio_nombre(nombre_completo):
        nom_norm = normalizar_texto(nombre_completo)
        if not nom_norm: return False
        for tn in target_names_norm:
            if tn in nom_norm or nom_norm in tn:
                return True
        return False

    df_base_parts = []

    # ---------------------------------------------------------
    # PASO 1: Universo de Alumnos (Matrícula + Autogenerados)
    # ---------------------------------------------------------
    
    # --- FUENTE A: Matrícula Oficial ---
    if os.path.exists(PATH_MATRICULA):
        print("   -> Escaneando Matrícula Oficial...")
        df_mat = pd.read_csv(PATH_MATRICULA, sep=';', dtype=str, encoding_errors='ignore')
        
        col_nie = next((c for c in df_mat.columns if 'nie' in c.lower()), None)
        col_grado = next((c for c in df_mat.columns if 'grado' in c.lower()), None)
        col_sec_nom = next((c for c in df_mat.columns if 'nombre_secc' in c.lower()), None)
        col_sec_cod = next((c for c in df_mat.columns if 'código_secc' in c.lower() or 'codigo_secc' in c.lower()), None)
        
        # Columnas de centro (NUEVO)
        col_cod_centro = next((c for c in df_mat.columns if 'codigo' in c.lower() or 'código' in c.lower() or 'centro' in c.lower()), None)
        col_nom_centro = next((c for c in df_mat.columns if 'nombre_centro' in c.lower() or 'nombre_inst' in c.lower() or ('nombre' in c.lower() and 'centro' in c.lower())), None)
        
        col_a1 = next((c for c in df_mat.columns if 'primer_apellido' in c.lower()), None)
        col_a2 = next((c for c in df_mat.columns if 'segundo_apellido' in c.lower()), None)
        col_n1 = next((c for c in df_mat.columns if 'primer_nombre' in c.lower()), None)
        col_n2 = next((c for c in df_mat.columns if 'segundo_nombre' in c.lower()), None)

        if col_nie:
            df_mat['temp_nombre'] = df_mat.apply(lambda row: limpiar_nombre(
                row[col_a1] if col_a1 else "", 
                row[col_a2] if col_a2 else "", 
                row[col_n1] if col_n1 else "", 
                row[col_n2] if col_n2 else ""
            ), axis=1)
            df_mat['temp_nie'] = df_mat[col_nie].astype(str).str.strip().str.replace('.0', '', regex=False)
            
            mask_nie = df_mat['temp_nie'].isin(target_nies)
            mask_name = df_mat['temp_nombre'].apply(cumple_criterio_nombre)
            df_mat_filtrado = df_mat[mask_nie | mask_name].copy()
            
            if not df_mat_filtrado.empty:
                df_mat_filtrado['NIE'] = df_mat_filtrado['temp_nie']
                df_mat_filtrado['Nombre Completo'] = df_mat_filtrado['temp_nombre']
                df_mat_filtrado['Grado'] = df_mat_filtrado[col_grado].apply(normalizar_y_formatear_grado) if col_grado else None
                
                # Extraer datos de escuela
                df_mat_filtrado['Código Centro'] = df_mat_filtrado[col_cod_centro].astype(str).str.strip().str.replace('.0', '', regex=False) if col_cod_centro else ""
                df_mat_filtrado['Nombre Centro'] = df_mat_filtrado[col_nom_centro].astype(str).str.strip() if col_nom_centro else ""
                
                def armar_seccion(nom, cod):
                    n = str(nom).strip() if pd.notna(nom) else ""
                    c = str(cod).strip().replace('.0', '') if pd.notna(cod) else ""
                    if n and c: return f"{n} ({c})"
                    elif n: return n
                    elif c: return f"({c})"
                    return "Sin Sección"
                    
                df_mat_filtrado['Sección'] = df_mat_filtrado.apply(lambda row: armar_seccion(
                    row[col_sec_nom] if col_sec_nom else "", 
                    row[col_sec_cod] if col_sec_cod else ""
                ), axis=1)
                
                df_base_parts.append(df_mat_filtrado[['NIE', 'Nombre Completo', 'Código Centro', 'Nombre Centro', 'Sección', 'Grado']])
    else:
        print(f"   [!] ADVERTENCIA: No se encontró el archivo de matrícula en: {PATH_MATRICULA}")

    # --- FUENTE B: Registros Autogenerados ---
    if os.path.exists(PATH_AUTOGENERADOS):
        print(f"   -> Escaneando Registros Autogenerados (registros autogenerados nie-2026-05-14.csv)...")
        df_auto = pd.read_csv(PATH_AUTOGENERADOS, sep=None, engine='python', dtype=str, encoding_errors='ignore')
        
        col_auto_nie = 'alumno_nie'
        col_auto_nom = 'solicitante_nombre_completo'
        col_auto_gra = 'alumno_grado_curso'
        
        if col_auto_nie in df_auto.columns and col_auto_nom in df_auto.columns:
            df_auto['temp_nie'] = df_auto[col_auto_nie].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_auto['temp_nombre'] = df_auto[col_auto_nom].astype(str).str.strip()
            
            mask_auto_nie = df_auto['temp_nie'].isin(target_nies)
            mask_auto_name = df_auto['temp_nombre'].apply(cumple_criterio_nombre)
            df_auto_filtrado = df_auto[mask_auto_nie | mask_auto_name].copy()
            
            if not df_auto_filtrado.empty:
                df_auto_filtrado['NIE'] = df_auto_filtrado['temp_nie']
                df_auto_filtrado['Nombre Completo'] = df_auto_filtrado['temp_nombre']
                df_auto_filtrado['Grado'] = df_auto_filtrado[col_auto_gra].apply(normalizar_y_formatear_grado) if col_auto_gra in df_auto_filtrado.columns else None
                df_auto_filtrado['Sección'] = "Autogenerado"
                df_auto_filtrado['Código Centro'] = ""
                df_auto_filtrado['Nombre Centro'] = ""
                
                df_base_parts.append(df_auto_filtrado[['NIE', 'Nombre Completo', 'Código Centro', 'Nombre Centro', 'Sección', 'Grado']])
    else:
        print(f"   [!] ADVERTENCIA: No se encontró el archivo autogenerado en: {PATH_AUTOGENERADOS}")

    # Consolidar ambas fuentes priorizando filas que contengan escuela cargada
    if df_base_parts:
        df_base = pd.concat(df_base_parts)
        df_base['tiene_centro'] = df_base['Código Centro'].apply(lambda x: 1 if str(x).strip() != "" else 0)
        df_base = df_base.sort_values(by=['NIE', 'tiene_centro'], ascending=[True, False])
        df_base = df_base.drop_duplicates(subset=['NIE'], keep='first').drop(columns=['tiene_centro'])
    else:
        df_base = pd.DataFrame(columns=['NIE', 'Nombre Completo', 'Código Centro', 'Nombre Centro', 'Sección', 'Grado'])

    if df_base.empty:
        print(f"   [!] ERROR: No se localizó a los estudiantes en los archivos de origen.")
        return

    df_base = df_base.dropna(subset=['Grado'])
    found_nies = [nie for nie in df_base['NIE'].unique() if nie and str(nie).lower() != 'nan' and str(nie).strip() != '']

    # ---------------------------------------------------------
    # PASO 2: Extraer Resultados de Geiser e Información del Centro
    # ---------------------------------------------------------
    print("   -> Buscando puntajes y metadatos de escuelas en Geiser...")
    archivos_csv = glob.glob(os.path.join(PATH_RESULTADOS, "*.csv"))
    
    lista_mat = []
    lista_len = []

    for f in archivos_csv:
        try:
            df_res = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            
            c_nie = next((c for c in df_res.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), None)
            c_theta = next((c for c in df_res.columns if 'escala 0-100' in str(c).lower()), 'theta.global (escala 0-100)')
            c_fecha = next((c for c in df_res.columns if 'fecha' in str(c).lower() and 'inicio' in str(c).lower()), None)
            if not c_fecha: c_fecha = next((c for c in df_res.columns if 'inicio' in str(c).lower()), None)
            
            # Nuevos detectores para código y nombre de escuela dentro de Geiser
            c_cod = next((c for c in df_res.columns if 'nro de centro' in str(c).lower() or 'código' in str(c).lower() or 'centro' in str(c).lower()), None)
            c_nom_centro = next((c for c in df_res.columns if 'nombre_centro' in str(c).lower() or 'nombre de centro' in str(c).lower() or 'institucion' in str(c).lower() or 'nombre_inst' in str(c).lower()), None)
            
            if not c_nie or c_theta not in df_res.columns: continue
            
            df_res['temp_nie'] = df_res[c_nie].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_match = df_res[df_res['temp_nie'].isin(found_nies)].copy()
            
            if df_match.empty: continue
            
            df_match['NIE'] = df_match['temp_nie']
            df_match['Puntaje'] = pd.to_numeric(df_match[c_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df_match['Fecha'] = df_match[c_fecha] if c_fecha else ""
            df_match['Geiser_Cod_Centro'] = df_match[c_cod].astype(str).str.strip().str.replace('.0', '', regex=False) if c_cod else ""
            df_match['Geiser_Nom_Centro'] = df_match[c_nom_centro].astype(str).str.strip() if c_nom_centro else ""
            
            materia = 'Mat' if 'MAT' in os.path.basename(f).upper() else 'Len'
            df_match = df_match.sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
            
            if materia == 'Mat':
                lista_mat.append(df_match[['NIE', 'Puntaje', 'Fecha', 'Geiser_Cod_Centro', 'Geiser_Nom_Centro']])
            else:
                lista_len.append(df_match[['NIE', 'Puntaje', 'Fecha', 'Geiser_Cod_Centro', 'Geiser_Nom_Centro']])
                
        except Exception: pass

    # Consolidación de Notas
    if lista_mat:
        df_mat = pd.concat(lista_mat).sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
        df_mat.columns = ['NIE', 'Puntaje mat', 'Fecha de aplicación Mat', 'Cod_Centro_Mat', 'Nom_Centro_Mat']
    else:
        df_mat = pd.DataFrame(columns=['NIE', 'Puntaje mat', 'Fecha de aplicación Mat', 'Cod_Centro_Mat', 'Nom_Centro_Mat'])

    if lista_len:
        df_len = pd.concat(lista_len).sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
        df_len.columns = ['NIE', 'Puntaje len', 'Fecha de aplicación Len', 'Cod_Centro_Len', 'Nom_Centro_Len']
    else:
        df_len = pd.DataFrame(columns=['NIE', 'Puntaje len', 'Fecha de aplicación Len', 'Cod_Centro_Len', 'Nom_Centro_Len'])

    # ---------------------------------------------------------
    # PASO 3: Cruce de Datos y Cruce de Escuelas Extraviadas
    # ---------------------------------------------------------
    print("   -> Cruzando universos y corrigiendo escuelas vacías...")
    df_final = pd.merge(df_base, df_mat, on='NIE', how='left')
    df_final = pd.merge(df_final, df_len, on='NIE', how='left')
    
    # Lógica Backfill: Si la escuela venía en blanco, heredar de Geiser Matemática o Geiser Lengua
    def rellenar_centro(row):
        cod = str(row['Código Centro']).strip()
        nom = str(row['Nombre Centro']).strip()
        
        if cod == "" or cod.lower() == "nan":
            if 'Cod_Centro_Mat' in row and str(row['Cod_Centro_Mat']).strip() != "" and str(row['Cod_Centro_Mat']).lower() != "nan":
                cod = str(row['Cod_Centro_Mat']).strip()
            elif 'Cod_Centro_Len' in row and str(row['Cod_Centro_Len']).strip() != "" and str(row['Cod_Centro_Len']).lower() != "nan":
                cod = str(row['Cod_Centro_Len']).strip()
                
        if nom == "" or nom.lower() == "nan":
            if 'Nom_Centro_Mat' in row and str(row['Nom_Centro_Mat']).strip() != "" and str(row['Nom_Centro_Mat']).lower() != "nan":
                nom = str(row['Nom_Centro_Mat']).strip()
            elif 'Nom_Centro_Len' in row and str(row['Nom_Centro_Len']).strip() != "" and str(row['Nom_Centro_Len']).lower() != "nan":
                nom = str(row['Nom_Centro_Len']).strip()
        
        return pd.Series([cod, nom])

    df_final[['Código Centro', 'Nombre Centro']] = df_final.apply(rellenar_centro, axis=1)
    
    # Eliminar las columnas espejo de respaldo
    cols_drop = ['Cod_Centro_Mat', 'Nom_Centro_Mat', 'Cod_Centro_Len', 'Nom_Centro_Len']
    df_final = df_final.drop(columns=[c for c in cols_drop if c in df_final.columns])

    if 'Puntaje mat' in df_final.columns: df_final['Puntaje mat'] = df_final['Puntaje mat'].round(2)
    if 'Puntaje len' in df_final.columns: df_final['Puntaje len'] = df_final['Puntaje len'].round(2)
    df_final = df_final.fillna("")

    # Ordenamiento estructural
    df_final = df_final.sort_values(by=['Código Centro', 'Grado', 'Nombre Completo'])

    # Reordenar las columnas del reporte de forma lógica
    columnas_reporte = ['NIE', 'Nombre Completo', 'Código Centro', 'Nombre Centro', 'Grado', 'Sección', 'Puntaje mat', 'Fecha de aplicación Mat', 'Puntaje len', 'Fecha de aplicación Len']
    columnas_reporte = [c for c in columnas_reporte if c in df_final.columns]
    df_final = df_final[columnas_reporte]

    # ---------------------------------------------------------
    # PASO 4: Exportación y Diseño del Excel
    # ---------------------------------------------------------
    ruta_salida = os.path.join(DIR_SALIDA, "Reporte_Estudiantes_Busqueda_Expandida.xlsx")
    
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="Estudiantes")
    
    # Carga de la hoja para aplicar anchos y alineaciones dinámicas
    wb = load_workbook(ruta_salida)
    ws = wb["Estudiantes"]
    
    # Configuración de anchos ajustados a las nuevas columnas
    ws.column_dimensions['A'].width = 15 # NIE
    ws.column_dimensions['B'].width = 45 # Nombre Completo
    ws.column_dimensions['C'].width = 15 # Código Centro
    ws.column_dimensions['D'].width = 45 # Nombre Centro
    ws.column_dimensions['E'].width = 10 # Grado
    ws.column_dimensions['F'].width = 20 # Sección
    ws.column_dimensions['G'].width = 15 # Puntaje Mat
    ws.column_dimensions['H'].width = 25 # Fecha Mat
    ws.column_dimensions['I'].width = 15 # Puntaje Len
    ws.column_dimensions['J'].width = 25 # Fecha Len
    
    # Formato Header
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    # Formato de alineaciones por fila
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].alignment = Alignment(horizontal="center") # NIE
        row[2].alignment = Alignment(horizontal="center") # Código Centro
        row[4].alignment = Alignment(horizontal="center") # Grado
        if len(row) > 6: row[6].alignment = Alignment(horizontal="center")  # Ptos Mat
        if len(row) > 7: row[7].alignment = Alignment(horizontal="center")  # Fecha Mat
        if len(row) > 8: row[8].alignment = Alignment(horizontal="center")  # Ptos Len
        if len(row) > 9: row[9].alignment = Alignment(horizontal="center")  # Fecha Len

    wb.save(ruta_salida)
    print(f"\n[OK] ¡Reporte generado con la información de centros educativos incorporada!")
    print(f"      --> Guardado en: {ruta_salida}\n")

if __name__ == "__main__":
    generar_reporte_nominal()