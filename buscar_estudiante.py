import os
import pandas as pd

# --- CONFIGURACIÓN DE RUTAS (Ajusta si es necesario) ---
DIR_ABRIL = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"
META_FILE = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\CentroVirtual_Progreso2.csv"

# --- DATOS DE BÚSQUEDA ---
TARGET_CENTRO = "11386"
# Lista de estudiantes a auditar
ESTUDIANTES_A_BUSCAR = [
    {"nie": "10465984", "n1": "joaquin", "n2": "telenchana"},
    {"nie": "10597415", "n1": "ethan", "n2": "escobar"}
    # Puedes agregar el tercer estudiante aquí si consigues el NIE/Nombre
]

def buscar_en_metadata(estudiante):
    """Busca al estudiante en el Centro Virtual para obtener IDs adicionales."""
    print(f"▶️ PASO 1: Buscando a {estudiante['n1'].upper()} en Metadata...")
    if not os.path.exists(META_FILE):
        return []
    
    ids_extraidos = []
    try:
        df = pd.read_csv(META_FILE, dtype=str, encoding='utf-8-sig')
        df.columns = df.columns.str.strip()
        df_clean = df.fillna('').astype(str).replace(r'\.0$', '', regex=True)
        df_clean['FILA_TXT'] = df_clean.apply(lambda row: ' '.join(row).lower(), axis=1)
        
        mask_nie = df_clean['FILA_TXT'].str.contains(estudiante['nie'], regex=False)
        mask_nombre = df_clean['FILA_TXT'].str.contains(estudiante['n1'], regex=False) & \
                      df_clean['FILA_TXT'].str.contains(estudiante['n2'], regex=False)
        
        match = df[mask_nie | mask_nombre]
        
        if not match.empty:
            print(f"✅ Encontrado en Metadata.")
            col_autogen = next((c for c in df.columns if 'autogenerado' in c.lower()), None)
            if col_autogen:
                for val in match[col_autogen].dropna():
                    ids_extraidos.append(str(val).replace('.0', '').strip())
    except Exception as e:
        print(f"⚠️ Error en Metadata: {e}")
    
    return list(set(ids_extraidos))

def escanear_resultados(filepath, estudiante, ids_auto):
    """Busca al estudiante en un archivo de resultados específico."""
    filename = os.path.basename(filepath)
    try:
        df = pd.read_csv(filepath, dtype=str, encoding='utf-8-sig')
        df.columns = df.columns.str.strip()
        df_clean = df.fillna('').astype(str).replace(r'\.0$', '', regex=True)
        df_clean['FILA_TXT'] = df_clean.apply(lambda row: ' '.join(row).lower(), axis=1)
        
        # Búsqueda por NIE o Nombre
        mask_nie = df_clean['FILA_TXT'].str.contains(estudiante['nie'], regex=False)
        mask_nombre = df_clean['FILA_TXT'].str.contains(estudiante['n1'], regex=False) & \
                      df_clean['FILA_TXT'].str.contains(estudiante['n2'], regex=False)
        
        # Búsqueda por ID Autogenerado si existe
        mask_auto = pd.Series([False] * len(df))
        id_col = next((col for col in ['Documento', 'NIE', 'nie'] if col in df.columns), None)
        if id_col and ids_auto:
            mask_auto = df[id_col].str.replace(r'\.0$', '', regex=True).isin(ids_auto)
            
        match = df[mask_nie | mask_nombre | mask_auto]
        
        if not match.empty:
            print(f"\n🎯 ¡HALLAZGO en {filename}!")
            for _, row in match.iterrows():
                print(f"   > NIE: {row.get('NIE' if 'NIE' in df.columns else id_col, 'N/A')}")
                print(f"   > Nombre: {estudiante['n1']} {estudiante['n2']}")
                print(f"   > Grado/Sección: {row.get('Grado', 'N/A')} {row.get('Sección', 'N/A')}")
                print("-" * 30)
            return True
    except:
        pass
    return False

def iniciar_auditoria():
    print("=" * 80)
    print(f"🔍 AUDITORÍA CE: {TARGET_CENTRO} - ESTUDIANTES MISSING")
    print("=" * 80)

    archivos_csv = []
    if os.path.exists(DIR_ABRIL):
        archivos_csv = [f for f in os.listdir(DIR_ABRIL) if f.lower().endswith('.csv')]

    for est in ESTUDIANTES_A_BUSCAR:
        print(f"\n--- ANALIZANDO A: {est['n1'].upper()} (NIE: {est['nie']}) ---")
        ids_extra = buscar_en_metadata(est)
        hallado = False
        
        for p en archivos_csv:
            if escanear_resultados(os.path.join(DIR_ABRIL, p), est, ids_extra):
                hallado = True
        
        if not hallado:
            print(f"❌ El estudiante {est['n1'].upper()} NO fue encontrado en ningún archivo de resultados.")

    # Resumen del Centro
    print("\n" + "=" * 80)
    print(f"INFO ADICIONAL: Se recomienda verificar si el CE {TARGET_CENTRO} subió el archivo de 2B correctamente.")
    print("=" * 80)

if __name__ == "__main__":
    iniciar_auditoria()