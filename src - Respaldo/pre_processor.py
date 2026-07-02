# src/pre_processor.py
import os
import pandas as pd
import sys

# Añadimos la carpeta principal al sistema para que encuentre a config.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config 

def ejecutar_conversion():
    # Creamos la carpeta de destino si no existe
    os.makedirs(config.PATH_INTERIM, exist_ok=True)
    
    # Buscamos los Excel en la carpeta Raw_Data que definiste en config.py
    archivos_excel = [f for f in os.listdir(config.PATH_RAW) if f.endswith(('.xlsx', '.xls'))]
    
    if not archivos_excel:
        print(f"⚠️ No se encontraron archivos Excel en: {config.PATH_RAW}")
        return

    print(f"--- Iniciando conversión...")

    for nombre_archivo in archivos_excel:
        ruta_completa = os.path.join(config.PATH_RAW, nombre_archivo)
        try:
            # Leemos todas las pestañas del Excel
            contenido_excel = pd.read_excel(ruta_completa, sheet_name=None)
            
            for nombre_pestaña, df in contenido_excel.items():
                # Creamos un nombre único para el CSV
                nombre_csv = f"{nombre_archivo} - {nombre_pestaña}.csv"
                ruta_guardado = os.path.join(config.PATH_INTERIM, nombre_csv)
                
                # Guardamos como CSV
                df.to_csv(ruta_guardado, index=False, encoding='utf-8-sig')
                print(f"  [OK] Generado: {nombre_csv}")
                
        except Exception as e:
            print(f"  [ERROR] Error procesando {nombre_archivo}: {e}")

    print("\n Conversión terminada. Revisa la carpeta Interim_CSVs en tu Drive.")

if __name__ == "__main__":
    ejecutar_conversion()