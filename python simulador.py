# simulador.py
import os
import pandas as pd
import numpy as np

# --- RUTA ACTUALIZADA A TU GOOGLE DRIVE ---
base_dir = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026"

# Simulación ULTRA CONSERVADORA (Cuentagotas)
carpetas_simulacion = [
    {"nombre": "02_PROGRESO_Abril",      "pts_mat": 0.8,  "pts_lec": 1.2},
    {"nombre": "03_PROGRESO_Mayo",       "pts_mat": 1.6,  "pts_lec": 2.4},
    {"nombre": "04_PROGRESO_Junio",      "pts_mat": 2.4,  "pts_lec": 3.6},
    {"nombre": "05_PROGRESO_Julio",      "pts_mat": 3.2,  "pts_lec": 4.8},
    {"nombre": "06_PROGRESO_Agosto",     "pts_mat": 4.0,  "pts_lec": 6.0},
    {"nombre": "07_PROGRESO_Septiembre", "pts_mat": 4.8,  "pts_lec": 7.2},
    {"nombre": "08_PROGRESO_Octubre",    "pts_mat": 5.6,  "pts_lec": 8.4}
]

print("====================================================")
print("   INICIANDO LA MÁQUINA DEL TIEMPO (MODO REALISTA)")
print("====================================================\n")

for mes in carpetas_simulacion:
    ruta_interim = os.path.join(base_dir, mes["nombre"], "Interim_CSVs")
    
    if not os.path.exists(ruta_interim):
        print(f"[!] No encontré la carpeta: {ruta_interim}")
        continue
        
    # Ignora los archivos que tengan "compl" o que empiecen con "legend"
    archivos = [
        f for f in os.listdir(ruta_interim) 
        if f.endswith('.csv') 
        and 'compl' not in f.lower() 
        and not f.startswith('legend')
    ]
    
    modificados = 0
    
    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_interim, archivo)
        try:
            df = pd.read_csv(ruta_archivo, encoding_errors='ignore')
            col_exacta = "theta.global (escala 0-100)"
            
            if col_exacta in df.columns:
                if df[col_exacta].dtype == object:
                    df[col_exacta] = df[col_exacta].astype(str).str.replace('"', '').str.replace(',', '.')
                
                scores = pd.to_numeric(df[col_exacta], errors='coerce')
                
                if 'MAT' in archivo.upper():
                    puntos_sumar = mes["pts_mat"]
                elif 'LEC' in archivo.upper():
                    puntos_sumar = mes["pts_lec"]
                else:
                    puntos_sumar = 0
                
                nuevos_scores = scores + puntos_sumar
                nuevos_scores = nuevos_scores.clip(upper=100)
                
                df[col_exacta] = nuevos_scores.round(2)
                df.to_csv(ruta_archivo, index=False)
                
                print(f"  [+] {archivo} -> Añadidos {puntos_sumar} puntos.")
                modificados += 1
            else:
                print(f"  [X] {archivo} -> No se encontró la columna '{col_exacta}'.")
                
        except Exception as e:
            print(f"  [X] Error procesando {archivo}: {e}")
            
    if modificados > 0:
        print(f"  [OK] -> ¡{mes['nombre']} simulado con éxito!\n")
    else:
        print(f"  [!] -> CUIDADO: No se modificó ningún archivo en {mes['nombre']}.\n")

print("====================================================")
print("  ¡SIMULACIÓN REALISTA COMPLETADA!")
print("====================================================")