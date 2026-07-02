# src/generador_consolidado_estudiantes.py
import os
import sys
import pandas as pd
import glob
import re

# Conexión con tu sistema de configuración principal
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def clasificar_puntaje(val):
    if pd.isna(val) or val == '': return "Sin Nota"
    try:
        val = float(val)
        if val <= 35: return "Crítico"
        elif val <= 45: return "Bajo"
        elif val <= 55: return "Medio"
        elif val <= 65: return "Bueno"
        else: return "Excelente"
    except:
        return "Sin Nota"

def cargar_y_limpiar_estudiantes(path, mes_etiqueta):
    """Extrae las notas de todos los estudiantes desde una carpeta específica."""
    print(f"  -> Extrayendo estudiantes de: Mes {mes_etiqueta} ...")
    archivos = glob.glob(os.path.join(path, "*.csv"))
    
    lista_dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            # Identificar materia
            nombre_archivo = os.path.basename(f).upper()
            if 'MAT' in nombre_archivo: materia = 'Matemática'
            elif 'LEC' in nombre_archivo or 'LENG' in nombre_archivo: materia = 'Lectura'
            else: continue

            # Encontrar columnas clave
            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            col_grado = next((c for c in df.columns if 'grado' in c.lower()), None)
            col_nie = next((c for c in df.columns if 'documento' in c.lower()), None)
            col_centro = next((c for c in df.columns if 'nro de centro' in c.lower() or 'centro' in c.lower()), None)
            
            # Lógica robusta para encontrar la columna de Grupo/Sección
            col_grupos = [c for c in df.columns if 'rup' in str(c).lower() or 'ecc' in str(c).lower()]
            grupo_col_final = None
            if col_grupos:
                grupo_col_final = col_grupos[0]
                for c in col_grupos:
                    if df[c].astype(str).str.contains(r'\(\d+\)', regex=True).any():
                        grupo_col_final = c
                        break

            if not all([col_theta, col_grado, col_nie, col_centro]):
                continue

            # Filtro de anulados (Solo estudiantes con pruebas válidas)
            if col_anular:
                df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')].copy()
            
            if df.empty: continue

            # Limpieza y conversión de NIE, Grado, Grupo y Escuela
            df['NIE (Documento)'] = df[col_nie].str.replace(r'\.0$', '', regex=True).str.strip()
            df['Grado'] = df[col_grado].str.strip()
            df['Código de Infraestructura'] = df[col_centro].str.replace(r'\.0$', '', regex=True).str.strip()
            
            if grupo_col_final:
                df['Grupo'] = df[grupo_col_final].astype(str).str.strip()
            else:
                df['Grupo'] = "A" # Valor por defecto si no existe la columna
            
            # --- FILTRO: IGNORAR CENTRO VIRTUAL (99999) ---
            df = df[df['Código de Infraestructura'] != '99999']
            if df.empty: continue
            
            # Convertir nota y clasificar nivel
            df[f'Nota Mes {mes_etiqueta}'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df.dropna(subset=[f'Nota Mes {mes_etiqueta}'], inplace=True)
            df[f'Nivel {mes_etiqueta}'] = df[f'Nota Mes {mes_etiqueta}'].apply(clasificar_puntaje)
            df['Materia'] = materia
            
            # Seleccionamos y ordenamos las columnas base
            df_final = df[['NIE (Documento)', 'Grado', 'Grupo', 'Código de Infraestructura', f'Nota Mes {mes_etiqueta}', f'Nivel {mes_etiqueta}', 'Materia']]
            lista_dfs.append(df_final)

        except Exception as e:
            pass
            
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()

def obtener_rutas_historicas():
    """Descubre dinámicamente cuántos meses anteriores existen hasta llegar al mes seleccionado."""
    base_month_dir = os.path.dirname(config.PATH_INTERIM)  
    base_year_dir = os.path.dirname(base_month_dir)        
    
    current_month_folder = os.path.basename(base_month_dir) 
    
    # Diferenciar si el usuario está en PROGRESO o en RESULTADOS
    exam_keyword = "PROGRESO" if "PROGRESO" in current_month_folder.upper() else "RESULTADOS"
    
    try:
        all_dirs = [d for d in os.listdir(base_year_dir) if os.path.isdir(os.path.join(base_year_dir, d))]
    except:
        return []
        
    # Filtrar solo las carpetas que sean del mismo examen y que empiecen por números (ej: 01_, 02_)
    month_dirs = [d for d in all_dirs if exam_keyword in d.upper() and re.match(r'^\d{2}_', d)]
    month_dirs.sort() # Garantiza orden cronológico
    
    rutas_validas = []
    mes_num = 1
    for m_dir in month_dirs:
        # Construye la ruta apuntando específicamente a la carpeta Resultados
        path_resultados = os.path.join(base_year_dir, m_dir, "Interim_CSVs", "Resultados")
        if os.path.exists(path_resultados):
            rutas_validas.append((mes_num, path_resultados))
        mes_num += 1
        
        # Detener la búsqueda cuando lleguemos al mes que el usuario está procesando actualmente
        if m_dir == current_month_folder:
            break
            
    return rutas_validas

def generar_excel_estudiantes():
    print("\n[*] Iniciando módulo Histórico de Consolidación de Estudiantes...")
    
    rutas_meses = obtener_rutas_historicas()
    
    if not rutas_meses:
        print("[!] Error: No se encontraron carpetas históricas en Interim_CSVs/Resultados.")
        return
        
    print(f"[*] Escaneo completado: Se detectaron {len(rutas_meses)} meses históricos para unificar.")

    df_master_mat = None
    df_master_lec = None
    
    # Llaves maestras para no duplicar filas al hacer los cruces
    llaves_cruce = ['NIE (Documento)', 'Grado', 'Grupo', 'Código de Infraestructura']
    
    for mes_num, path_resultados in rutas_meses:
        df_mes = cargar_y_limpiar_estudiantes(path_resultados, str(mes_num))
        if df_mes.empty:
            continue
            
        df_mat = df_mes[df_mes['Materia'] == 'Matemática'].drop(columns=['Materia'])
        df_lec = df_mes[df_mes['Materia'] == 'Lectura'].drop(columns=['Materia'])
        
        # Cruce Matemática (Outer Join para no perder a los alumnos que faltaron a un mes)
        if df_master_mat is None:
            df_master_mat = df_mat
        else:
            df_master_mat = pd.merge(df_master_mat, df_mat, on=llaves_cruce, how='outer')
            
        # Cruce Lectura
        if df_master_lec is None:
            df_master_lec = df_lec
        else:
            df_master_lec = pd.merge(df_master_lec, df_lec, on=llaves_cruce, how='outer')

    if df_master_mat is None and df_master_lec is None:
        print("[!] Error: Los archivos CSV estaban vacíos o no tenían notas válidas.")
        return
        
    # Preparar el guardado automático en Final_Reports
    os.makedirs(config.PATH_REPORTS, exist_ok=True)
    nombre_archivo = f"Consolidado_Estudiantes_Seguimiento_Mes1_al_Mes{len(rutas_meses)}.xlsx"
    ruta_excel = os.path.join(config.PATH_REPORTS, nombre_archivo)
    
    with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
        for materia, df_final in [("Est_Matemática", df_master_mat), ("Est_Lengua", df_master_lec)]:
            if df_final is not None and not df_final.empty:
                
                # --- NUEVA LÓGICA DE ORDENAMIENTO COMPLETA ---
                # 1. Extraer el número real del grado en una columna invisible temporal
                df_final['Grado_Num'] = df_final['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
                
                # 2. Ordenar por Grado (2 al 11) -> Código de Infraestructura -> Grupo -> NIE
                df_final = df_final.sort_values(by=['Grado_Num', 'Código de Infraestructura', 'Grupo', 'NIE (Documento)'])
                
                # 3. Borrar la columna numérica temporal
                df_final = df_final.drop(columns=['Grado_Num'])
                # --------------------------------------------------------

                # Redondear notas para que se vea limpio
                cols_notas = [c for c in df_final.columns if 'Nota' in c]
                for col in cols_notas:
                    df_final[col] = df_final[col].round(2)
                
                # Si un estudiante no hizo la prueba un mes, rellenar sus huecos
                cols_nivel = [c for c in df_final.columns if 'Nivel' in c]
                df_final[cols_nivel] = df_final[cols_nivel].fillna('No Evaluado')
                
                df_final.to_excel(writer, sheet_name=materia, index=False)
            
    print(f"\n[OK] ¡Excel de seguimiento de estudiantes creado con éxito!")
    print(f" -> Creado en: {ruta_excel}")

if __name__ == "__main__":
    generar_excel_estudiantes()