# src/calcular_idesp.py
import os
import sys
import pandas as pd
import glob

# Conexión con el mapa PAARS
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import config

def ejecutar_calculo_idesp():
    print("\n[*] Iniciando Motor de Cálculo IDESP Institucional...")
    
    # Usamos las rutas dinámicas establecidas en config.py
    PATH_INTERIM = config.PATH_INTERIM
    PATH_REPORTS = config.PATH_REPORTS
    
    archivos_csv = glob.glob(os.path.join(PATH_INTERIM, "*.csv"))
    if not archivos_csv:
        print("  [!] No se encontraron archivos en Interim_CSVs para calcular el IDESP.")
        return
        
    lista_dfs = []
    print(f"  -> Leyendo {len(archivos_csv)} archivos de la evaluación actual...")
    
    for f in archivos_csv:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            # Limpiamos anulados
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            if col_anular:
                df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')]
                
            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_centro = next((c for c in df.columns if 'centro' in c.lower()), None)
            
            if col_theta and col_centro:
                df['Puntaje'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
                df['Centro'] = df[col_centro].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df = df.dropna(subset=['Puntaje', 'Centro'])
                
                if not df.empty:
                    materia = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
                    df['Materia'] = materia
                    lista_dfs.append(df[['Centro', 'Puntaje', 'Materia']])
        except Exception as e:
            print(f"  [!] Error leyendo {os.path.basename(f)}: {e}")
            
    if not lista_dfs:
        print("  [!] No se extrajeron datos válidos para el IDESP.")
        return
        
    df_master = pd.concat(lista_dfs, ignore_index=True)
    
    # Cálculos del IDESP (Promedio por escuela y materia)
    idesp_escuelas = df_master.groupby(['Centro', 'Materia'])['Puntaje'].mean().unstack(fill_value=0)
    
    # Calcular el Índice Global (Promedio simple entre Matemática y Lengua)
    idesp_escuelas['IDESP_Global'] = idesp_escuelas.mean(axis=1)
    idesp_escuelas = idesp_escuelas.round(2).reset_index()
    idesp_escuelas = idesp_escuelas.sort_values(by='IDESP_Global', ascending=False)
    
    # Guardar reporte
    out_dir = os.path.join(PATH_REPORTS, "00_Calculo_IDESP")
    os.makedirs(out_dir, exist_ok=True)
    ruta_salida = os.path.join(out_dir, f"Reporte_IDESP_{config.MONTH_FOLDER}.xlsx")
    
    idesp_escuelas.to_excel(ruta_salida, index=False)
    print(f"  [OK] Reporte IDESP generado exitosamente en:\n       {ruta_salida}")

if __name__ == "__main__":
    ejecutar_calculo_idesp()