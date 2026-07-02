# src/conversor_excel_csv.py
import os
import glob
import pandas as pd
import config

def convertir_todos_excel_a_csv():
    print("======================================================")
    print("   MOTOR DE CONVERSIÓN: EXCEL A CSV (INTERIM)")
    print("======================================================\n")
    
    # Intenta deducir la carpeta raíz (ej. PAARS_Warehouse\2026) subiendo de nivel desde la configuración
    try:
        if hasattr(config, 'PATH_INTERIM'):
            exam_dir = os.path.dirname(config.PATH_INTERIM)
            root_dir = os.path.dirname(exam_dir)
        else:
            # Fallback a la ruta absoluta por seguridad
            root_dir = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026"
    except:
        root_dir = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026"

    print(f"[*] Escaneando exámenes en: {root_dir}")
    
    # Busca todas las carpetas Raw_Data dentro del año
    patron_raw = os.path.join(root_dir, "*", "Raw_Data")
    carpetas_raw = glob.glob(patron_raw)
    
    if not carpetas_raw:
        print("[!] No se encontraron carpetas Raw_Data.")
        return

    archivos_convertidos = 0

    for raw_folder in carpetas_raw:
        exam_folder = os.path.dirname(raw_folder)
        exam_name = os.path.basename(exam_folder)
        interim_folder = os.path.join(exam_folder, "Interim_CSVs")
        
        # Se asegura de que la carpeta Interim exista
        os.makedirs(interim_folder, exist_ok=True)
        
        # Busca los Excels, ignorando archivos temporales de Windows (los que empiezan con ~)
        excel_files = [f for f in glob.glob(os.path.join(raw_folder, "*.xlsx")) if not os.path.basename(f).startswith('~')]
        
        if excel_files:
            print(f"  -> Revisando: {exam_name}")
            
        for excel_file in excel_files:
            excel_name = os.path.basename(excel_file)
            try:
                xls = pd.ExcelFile(excel_file)
                for sheet_name in xls.sheet_names:
                    # Formato deseado: LEC-resultados.xlsx - 3er.csv
                    csv_filename = f"{excel_name} - {sheet_name}.csv"
                    csv_path = os.path.join(interim_folder, csv_filename)
                    
                    # Solo convierte si el archivo no existe, ahorrando muchísimo tiempo
                    if not os.path.exists(csv_path):
                        print(f"     [+] Extrayendo hoja '{sheet_name}' de {excel_name}...")
                        df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
                        # Guarda en utf-8-sig para evitar problemas con tildes y caracteres especiales
                        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                        archivos_convertidos += 1
            except Exception as e:
                print(f"     [!] Error procesando {excel_name}: {e}")

    if archivos_convertidos == 0:
        print("\n[OK] Todos los archivos CSV ya estaban actualizados. No se requirió conversión.")
    else:
        print(f"\n[OK] Conversión finalizada. {archivos_convertidos} nuevas hojas extraídas a CSV.")