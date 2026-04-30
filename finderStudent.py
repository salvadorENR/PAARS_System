import os
import glob
import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS Y CÓDIGO
# ==========================================
TARGET_SCHOOL = "11532"

PATH_MATRICULA = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes2.csv"
PATH_RESULTADOS = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"
DIR_SALIDA = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Final_Reports"

os.makedirs(DIR_SALIDA, exist_ok=True)

# ==========================================
# 2. FUNCIONES DE LIMPIEZA
# ==========================================
def normalizar_y_formatear_grado(grado_str):
    """Extrae el número del grado y le agrega el símbolo °. Ignora Tercer Año."""
    if pd.isna(grado_str): return None
    texto = str(grado_str).lower().strip()
    
    # Ignorar Tercer año
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
    """Concatena los nombres y apellidos omitiendo los valores nulos"""
    partes = [ape1, ape2, nom1, nom2]
    partes_limpias = [str(p).strip() for p in partes if pd.notna(p) and str(p).strip() != ""]
    return " ".join(partes_limpias)

# ==========================================
# 3. MOTOR PRINCIPAL
# ==========================================
def generar_reporte_nominal():
    print(f"\n[*] INICIANDO CRUCE NOMINAL DE ESTUDIANTES PARA LA ESCUELA {TARGET_SCHOOL}...")

    # ---------------------------------------------------------
    # PASO 1: Cargar Universo de Estudiantes (Matrícula)
    # ---------------------------------------------------------
    print("   -> Construyendo universo desde la Matrícula Oficial...")
    if not os.path.exists(PATH_MATRICULA):
        print(f"   [!] ERROR: No se encontró el archivo de matrícula en:\n       {PATH_MATRICULA}")
        sys.exit()

    df_mat = pd.read_csv(PATH_MATRICULA, sep=';', dtype=str, encoding_errors='ignore')
    
    # Identificar columnas
    col_cod = next((c for c in df_mat.columns if 'codigo' in c.lower() or 'código' in c.lower()), None)
    col_nie = next((c for c in df_mat.columns if 'nie' in c.lower()), None)
    col_grado = next((c for c in df_mat.columns if 'grado' in c.lower()), None)
    col_sec_nom = next((c for c in df_mat.columns if 'nombre_secc' in c.lower()), None)
    col_sec_cod = next((c for c in df_mat.columns if 'código_secc' in c.lower() or 'codigo_secc' in c.lower()), None)
    
    # Nombres
    col_a1 = next((c for c in df_mat.columns if 'primer_apellido' in c.lower()), None)
    col_a2 = next((c for c in df_mat.columns if 'segundo_apellido' in c.lower()), None)
    col_n1 = next((c for c in df_mat.columns if 'primer_nombre' in c.lower()), None)
    col_n2 = next((c for c in df_mat.columns if 'segundo_nombre' in c.lower()), None)

    # Filtrar por escuela
    df_base = df_mat[df_mat[col_cod].astype(str).str.strip().str.replace('.0', '', regex=False) == TARGET_SCHOOL].copy()
    
    if df_base.empty:
        print(f"   [!] ADVERTENCIA: La escuela {TARGET_SCHOOL} no aparece en la Matrícula.")
        return

    # Construir columnas requeridas
    df_base['NIE'] = df_base[col_nie].astype(str).str.strip()
    df_base['Grado'] = df_base[col_grado].apply(normalizar_y_formatear_grado)
    
    # Descartar los que son None (Tercer año o grados no válidos)
    df_base = df_base.dropna(subset=['Grado'])
    
    # Nombre Completo
    df_base['Nombre Completo'] = df_base.apply(lambda row: limpiar_nombre(
        row[col_a1] if col_a1 else "", 
        row[col_a2] if col_a2 else "", 
        row[col_n1] if col_n1 else "", 
        row[col_n2] if col_n2 else ""
    ), axis=1)
    
    # Sección
    def armar_seccion(nom, cod):
        n = str(nom).strip() if pd.notna(nom) else ""
        c = str(cod).strip().replace('.0', '') if pd.notna(cod) else ""
        if n and c: return f"{n} ({c})"
        elif n: return n
        elif c: return f"({c})"
        return "Sin Sección"
        
    df_base['Sección'] = df_base.apply(lambda row: armar_seccion(
        row[col_sec_nom] if col_sec_nom else "", 
        row[col_sec_cod] if col_sec_cod else ""
    ), axis=1)

    # Dejar solo las columnas base y quitar duplicados por si acaso
    df_base = df_base[['NIE', 'Nombre Completo', 'Sección', 'Grado']].drop_duplicates(subset=['NIE'])

    # ---------------------------------------------------------
    # PASO 2: Extraer Resultados de Geiser (Matemática y Lengua)
    # ---------------------------------------------------------
    print("   -> Buscando puntajes y fechas en los archivos Geiser...")
    archivos_csv = glob.glob(os.path.join(PATH_RESULTADOS, "*.csv"))
    
    lista_mat = []
    lista_len = []

    for f in archivos_csv:
        try:
            df_res = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            
            c_cod = next((c for c in df_res.columns if 'nro de centro' in str(c).lower() or 'código' in str(c).lower() or 'centro' in str(c).lower()), None)
            c_nie = next((c for c in df_res.columns if 'documento' in str(c).lower() or 'nie' in str(c).lower()), None)
            c_theta = next((c for c in df_res.columns if 'escala 0-100' in str(c).lower()), 'theta.global (escala 0-100)')
            
            c_fecha = next((c for c in df_res.columns if 'fecha' in str(c).lower() and 'inicio' in str(c).lower()), None)
            if not c_fecha: c_fecha = next((c for c in df_res.columns if 'inicio' in str(c).lower()), None)
            
            if not c_cod or not c_nie or c_theta not in df_res.columns: continue
            
            df_res[c_cod] = df_res[c_cod].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_school = df_res[df_res[c_cod] == TARGET_SCHOOL].copy()
            
            if df_school.empty: continue
            
            df_school['NIE'] = df_school[c_nie].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_school['Puntaje'] = pd.to_numeric(df_school[c_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df_school['Fecha'] = df_school[c_fecha] if c_fecha else ""
            
            materia = 'Mat' if 'MAT' in os.path.basename(f).upper() else 'Len'
            
            # Quitar duplicados priorizando a los que sí tienen nota
            df_school = df_school.sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
            
            if materia == 'Mat':
                lista_mat.append(df_school[['NIE', 'Puntaje', 'Fecha']])
            else:
                lista_len.append(df_school[['NIE', 'Puntaje', 'Fecha']])
                
        except Exception: pass

    # Consolidar DataFrames de Resultados
    if lista_mat:
        df_mat = pd.concat(lista_mat).sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
        df_mat.columns = ['NIE', 'Puntaje mat', 'Fecha de aplicación Mat']
    else:
        df_mat = pd.DataFrame(columns=['NIE', 'Puntaje mat', 'Fecha de aplicación Mat'])

    if lista_len:
        df_len = pd.concat(lista_len).sort_values(by='Puntaje', ascending=False).drop_duplicates(subset=['NIE'])
        df_len.columns = ['NIE', 'Puntaje len', 'Fecha de aplicación Len']
    else:
        df_len = pd.DataFrame(columns=['NIE', 'Puntaje len', 'Fecha de aplicación Len'])

    # ---------------------------------------------------------
    # PASO 3: Cruce (Merge) de los datos
    # ---------------------------------------------------------
    print("   -> Cruzando Matrícula con Resultados...")
    
    # Left Join: Mantenemos a TODOS los de matrícula. Si no están en Geiser, quedan en NaN
    df_final = pd.merge(df_base, df_mat, on='NIE', how='left')
    df_final = pd.merge(df_final, df_len, on='NIE', how='left')
    
    # Redondear puntajes para mejor visualización y llenar vacíos con strings vacíos
    if 'Puntaje mat' in df_final.columns: df_final['Puntaje mat'] = df_final['Puntaje mat'].round(2)
    if 'Puntaje len' in df_final.columns: df_final['Puntaje len'] = df_final['Puntaje len'].round(2)
    df_final = df_final.fillna("")

    # Ordenar lógicamente por Grado y Sección
    df_final = df_final.sort_values(by=['Grado', 'Sección', 'Nombre Completo'])

    # ---------------------------------------------------------
    # PASO 4: Exportación al Excel
    # ---------------------------------------------------------
    ruta_salida = os.path.join(DIR_SALIDA, f"Reporte_Nominal_Escuela_{TARGET_SCHOOL}.xlsx")
    
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="Estudiantes")
    
    # Aplicar Formato Visual
    wb = load_workbook(ruta_salida)
    ws = wb["Estudiantes"]
    
    ws.column_dimensions['A'].width = 15 # NIE
    ws.column_dimensions['B'].width = 45 # Nombre
    ws.column_dimensions['C'].width = 25 # Seccion
    ws.column_dimensions['D'].width = 10 # Grado
    ws.column_dimensions['E'].width = 15 # Puntaje Mat
    ws.column_dimensions['F'].width = 25 # Fecha Mat
    ws.column_dimensions['G'].width = 15 # Puntaje Len
    ws.column_dimensions['H'].width = 25 # Fecha Len
    
    # Dar color a la cabecera
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    # Centrar columnas de datos técnicos
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].alignment = Alignment(horizontal="center") # NIE
        row[3].alignment = Alignment(horizontal="center") # Grado
        row[4].alignment = Alignment(horizontal="center") # Ptos Mat
        row[5].alignment = Alignment(horizontal="center") # Fecha Mat
        row[6].alignment = Alignment(horizontal="center") # Ptos Len
        row[7].alignment = Alignment(horizontal="center") # Fecha Len

    wb.save(ruta_salida)
    print(f"\n[OK] ¡Archivo creado exitosamente! Los que no hicieron la prueba tendrán celdas vacías.")
    print(f"     --> {ruta_salida}\n")

if __name__ == "__main__":
    generar_reporte_nominal()