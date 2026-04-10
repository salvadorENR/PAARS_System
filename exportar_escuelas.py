# cruzar_mis_escuelas.py
import os
import pandas as pd
import sys
import io

# Conexión a tu configuración
try:
    import config
except ImportError:
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    import config

# --- AQUÍ ESTÁN TUS DATOS EXACTOS ---
DATOS_USUARIO = """Nro de Centro	Valor
10057	218
10084	175
10096	96
10123	155
10160	453
10209	56
11015	97
11307	483
11308	248
11314	47
11317	140
11321	211
11327	190
11330	471
11331	168
11337	78
11342	264
11349	458
11353	374
11365	271
11394	200
11395	414
11406	142
11409	168
11411	322
11420	318
11422	193
11426	212
11433	497
11436	332
11437	133
11439	237
11440	86
11443	313
11471	500
11474	109
11482	141
11483	309
11495	523
11499	137
11501	270
11507	353
11531	207
11532	482
11580	241
11594	291
11597	141
11598	176
11631	126
11640	140
11668	10
11680	166
11695	457
11711	185
11724	45
11941	63
11963	77
12037	36
13713	386
14806	544
60247	204
99999	59
"""

def cruzar_nombres_escuelas():
    print("[*] Leyendo tu lista de códigos...")
    # Leer los datos del texto de arriba como si fuera un archivo
    df_usuario = pd.read_csv(io.StringIO(DATOS_USUARIO), sep='\t')
    df_usuario['Nro de Centro'] = df_usuario['Nro de Centro'].astype(str).str.strip()
    
    ruta_csvs = config.PATH_INTERIM
    diccionario_escuelas = {}
    
    if not os.path.exists(ruta_csvs):
        print(f"[!] No se encontró la carpeta de Geiser en: {ruta_csvs}")
        return
        
    print("[*] Escaneando archivos Geiser para extraer el diccionario de nombres...")
    
    for f in os.listdir(ruta_csvs):
        if not f.endswith('.csv') or 'legend' in f.lower():
            continue
            
        try:
            path = os.path.join(ruta_csvs, f)
            df_g = pd.read_csv(path, dtype=str, encoding='utf-8-sig', on_bad_lines='skip')
            
            col_cod = next((c for c in df_g.columns if 'nro de centro' in str(c).lower() or 'código' in str(c).lower()), None)
            col_nom = next((c for c in df_g.columns if 'centro' in str(c).lower() and 'nro' not in str(c).lower()), None)
            
            if col_cod and col_nom:
                df_filtro = df_g[[col_cod, col_nom]].dropna().drop_duplicates()
                for _, row in df_filtro.iterrows():
                    cod = str(row[col_cod]).strip()
                    nom = str(row[col_nom]).strip()
                    if cod != 'nan' and nom != 'nan':
                        diccionario_escuelas[cod] = nom
        except Exception as e:
            pass
            
    print("[*] Cruzando tu lista con la base de datos...")
    # Asignar el nombre usando el diccionario. Si no existe, poner un aviso.
    df_usuario['Nombre del Centro'] = df_usuario['Nro de Centro'].map(diccionario_escuelas).fillna("⚠️ No encontrado en Geiser")
    
    # Reordenar las columnas para que el nombre quede en medio
    df_usuario = df_usuario[['Nro de Centro', 'Nombre del Centro', 'Valor']]
    
    # --- RUTA DE SALIDA FIJADA EXACTAMENTE DONDE LA PEDISTE ---
    carpeta_destino = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo"
    
    # Asegurarnos de que la carpeta exista, por si acaso
    os.makedirs(carpeta_destino, exist_ok=True)
    
    ruta_salida = os.path.join(carpeta_destino, "Mis_Escuelas_Buscadas.xlsx")
    
    df_usuario.to_excel(ruta_salida, index=False)
    
    print(f"\n[OK] ¡Éxito! Archivo Excel generado con tus {len(df_usuario)} escuelas.")
    print(f"📁 Puedes encontrar tu archivo aquí: {ruta_salida}")

if __name__ == "__main__":
    cruzar_nombres_escuelas()