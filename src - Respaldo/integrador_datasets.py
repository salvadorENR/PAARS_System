import os
import glob
import pandas as pd

def unificar_datasets_crudos(path_raw, path_reports):
    print("\n=======================================================")
    print("  INTEGRANDO Y UNIFICANDO ARCHIVOS CRUDOS (MAT + LEC)")
    print("=======================================================\n")
    
    # Buscar archivos Excel o CSV en la carpeta Raw_Data
    archivos = glob.glob(os.path.join(path_raw, "*.xlsx")) + glob.glob(os.path.join(path_raw, "*.csv"))
    
    if not archivos:
        print(f"[!] No se encontraron archivos en: {path_raw}")
        return False

    df_mat_list = []
    df_lec_list = []

    print("[*] Leyendo e identificando archivos...")
    for archivo in archivos:
        nombre_base = os.path.basename(archivo).lower()
        
        # Omitir archivos de leyenda (legend_bloque, etc.)
        if "legend" in nombre_base:
            continue
            
        es_mat = 'mat' in nombre_base
        es_lec = 'lec' in nombre_base or 'leng' in nombre_base
        
        if not es_mat and not es_lec:
            continue

        print(f"  - Procesando: {os.path.basename(archivo)}")
        
        # Leer el archivo (soporta CSV y Excel con múltiples hojas)
        if archivo.endswith('.csv'):
            df = pd.read_csv(archivo, dtype=str, encoding_errors='ignore')
            if es_mat: df_mat_list.append(df)
            if es_lec: df_lec_list.append(df)
        else:
            try:
                xls = pd.ExcelFile(archivo)
                # Leer todas las hojas (grados) dentro del Excel
                for sheet in xls.sheet_names:
                    if "legend" in sheet.lower(): continue
                    df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                    if es_mat: df_mat_list.append(df)
                    if es_lec: df_lec_list.append(df)
            except Exception as e:
                print(f"  [!] Error leyendo {nombre_base}: {e}")

    # Concatenar todos los grados verticalmente (Dataframes originales consolidados)
    df_mat_master = pd.concat(df_mat_list, ignore_index=True) if df_mat_list else pd.DataFrame()
    df_lec_master = pd.concat(df_lec_list, ignore_index=True) if df_lec_list else pd.DataFrame()

    print("\n[*] Unificando Matemática y Lengua (Outer Join)...")
    # Identificar columnas comunes de metadatos del estudiante para cruzar (elimina duplicados)
    claves_comunes = ['Departamento', 'Nro de centro', 'Centro', 'Grado', 'Grupo', 
                      'Documento', 'Nombre', 'Apellido']
    
    if not df_mat_master.empty and not df_lec_master.empty:
        # Asegurar que las claves existan en ambos datasets para poder cruzarlos
        claves_merge = [k for k in claves_comunes if k in df_mat_master.columns and k in df_lec_master.columns]
        
        # Realizar Outer Join: Conservamos estudiantes aunque solo tengan una materia
        df_final = pd.merge(df_mat_master, df_lec_master, on=claves_merge, how='outer', suffixes=('_MAT', '_LEC'))
    elif not df_mat_master.empty:
        df_final = df_mat_master
    elif not df_lec_master.empty:
        df_final = df_lec_master
    else:
        print("[!] No se encontraron datos válidos para procesar.")
        return False

    # Eliminar cualquier columna completamente duplicada que pudiera sobrevivir en la hoja integrada
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]

    # Exportar el resultado a la carpeta Final_Reports (path_reports)
    os.makedirs(path_reports, exist_ok=True)
    archivo_salida = os.path.join(path_reports, "0_Base_Maestra_Integrada.xlsx")
    
    print(f"[*] Guardando Base Maestra Única con 3 hojas en el archivo Excel...")
    
    # NUEVO: Escribir múltiples hojas en el mismo archivo
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        # Hoja 1: Todo el cruce integrado (Merge)
        df_final.to_excel(writer, sheet_name='Base_Integrada', index=False)
        
        # Hoja 2: Solo estructura y datos puros de Matemática
        if not df_mat_master.empty:
            df_mat_master.to_excel(writer, sheet_name='Matematica', index=False)
            
        # Hoja 3: Solo estructura y datos puros de Lengua
        if not df_lec_master.empty:
            df_lec_master.to_excel(writer, sheet_name='Lengua', index=False)
    
    print(f"\n[OK] ¡Integración exitosa! Archivo maestro guardado en:\n  -> {archivo_salida}")
    return True