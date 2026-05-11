import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# --- 1. CONFIGURATION ---
DIRECTORIES = {
    "FEBRERO": r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_Resultados_Febrero\Interim_CSVs\Resultados",
    "MARZO": r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados",
    "ABRIL": r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"
}

# --- DIRECTORIO DE SALIDA PARA EL EXCEL Y GRÁFICOS ---
OUTPUT_DIR = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\Requerimientos adicionales"

TARGET_SCHOOL_CODE = "88125"

NIVELES_CONFIG = [
    {'limite': 35, 'nombre': 'Crítico',   'color': '#991b1b', 'texto': 'white'},
    {'limite': 45, 'nombre': 'Bajo',      'color': '#ff8c2e', 'texto': 'white'},
    {'limite': 55, 'nombre': 'Medio',     'color': '#facc15', 'texto': '#333333'},
    {'limite': 65, 'nombre': 'Bueno',     'color': '#84cc16', 'texto': '#333333'},
    {'limite': 100,'nombre': 'Excelente', 'color': '#065f46', 'texto': 'white'}
]

def obtener_clasificacion(score):
    if pd.isna(score): return None
    for n in NIVELES_CONFIG:
        if score <= n['limite']: return n['nombre']
    return NIVELES_CONFIG[-1]['nombre']

def recopilar_datos_escuela():
    """Escanea los directorios, extrae datos y rastrea dónde se encontraron."""
    print("=" * 80)
    print(f"🔍 EXTRAYENDO DATOS DEL CE {TARGET_SCHOOL_CODE} DE TODAS LAS CARPETAS...")
    print("=" * 80)

    lista_dfs = []
    reporte_directorios = {mes: [] for mes in DIRECTORIES.keys()}

    for mes, dir_path in DIRECTORIES.items():
        if not os.path.exists(dir_path):
            reporte_directorios[mes] = None 
            continue

        archivos_csv = [f for f in os.listdir(dir_path) if f.lower().endswith('.csv')]

        for filename in archivos_csv:
            filepath = os.path.join(dir_path, filename)
            try:
                df = pd.read_csv(filepath, dtype=str, encoding='utf-8-sig')
            except:
                try: 
                    df = pd.read_csv(filepath, dtype=str, encoding='latin1')
                except:
                    continue 
            
            columnas_originales = df.columns.tolist()
            df.columns = df.columns.str.strip().str.lower()
            col_centro = next((col for col in df.columns if 'nro de centro' in col), None)
            
            if col_centro:
                df['codigo_limpio'] = df[col_centro].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                match = df[df['codigo_limpio'] == TARGET_SCHOOL_CODE].copy()
                
                if not match.empty:
                    reporte_directorios[mes].append(filename)
                    match.columns = columnas_originales + ['codigo_limpio']
                    
                    if 'Materia' not in match.columns and 'Area temática' not in match.columns:
                        match['Materia'] = 'Matemática' if 'MAT' in filename.upper() else 'Lengua'
                    elif 'Area temática' in match.columns:
                        match.rename(columns={'Area temática': 'Materia'}, inplace=True)

                    lista_dfs.append(match)

    # --- IMPRIMIR RESUMEN DE DIRECTORIOS ---
    print("\n" + "=" * 80)
    print("📁 RESUMEN DE BÚSQUEDA POR DIRECTORIO Y EXAMEN")
    print("=" * 80)
    for mes, archivos in reporte_directorios.items():
        if archivos is None:
            print(f"⚠️ {mes}: El directorio no existe en tu computadora.")
        elif len(archivos) > 0:
            print(f"✅ {mes}: Estudiantes encontrados en {len(archivos)} examen(es):")
            for arch in archivos:
                print(f"   -> {arch}")
        else:
            print(f"❌ {mes}: NO se encontraron estudiantes en esta carpeta.")
    print("=" * 80 + "\n")

    if not lista_dfs:
        print(f"❌ Abortando: No se encontró ningún dato para la escuela {TARGET_SCHOOL_CODE}.")
        return pd.DataFrame()

    df_final = pd.concat(lista_dfs, ignore_index=True)
    return df_final

def exportar_a_excel(df_resultados):
    """Filtra las columnas, redondea los puntajes y exporta a Excel."""
    if df_resultados.empty:
        return
        
    print(f"💾 Guardando archivo Excel en: {OUTPUT_DIR}...")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    def encontrar_columna(df, palabras_clave):
        for col in df.columns:
            col_limpia = str(col).lower()
            if any(kw in col_limpia for kw in palabras_clave):
                return col
        return None

    # Búsquedas de columnas generales
    col_nro = encontrar_columna(df_resultados, ['nro de centro', 'código de centro'])
    col_centro = encontrar_columna(df_resultados, ['centro', 'institucion'])
    col_grado = encontrar_columna(df_resultados, ['grado'])
    col_grupo = encontrar_columna(df_resultados, ['grupo', 'sección', 'seccion'])
    col_doc = encontrar_columna(df_resultados, ['documento', 'nie'])
    
    # Búsqueda ESTRICTA para la columna de puntaje específico
    col_theta_estricto = next((c for c in df_resultados.columns if 'theta.global (escala 0-100)' in str(c).lower()), None)

    if not col_theta_estricto:
        print("⚠️ Advertencia Excel: No se encontró la columna exacta 'theta.global (escala 0-100)'.")

    # Armado del DataFrame para Excel
    df_excel = pd.DataFrame()
    df_excel['Nro de centro'] = df_resultados[col_nro] if col_nro else TARGET_SCHOOL_CODE
    df_excel['Centro'] = df_resultados[col_centro] if col_centro else 'Desconocido'
    df_excel['Grado'] = df_resultados[col_grado] if col_grado else 'N/D'
    df_excel['Grupo'] = df_resultados[col_grupo] if col_grupo else 'N/D'
    df_excel['Documento'] = df_resultados[col_doc] if col_doc else 'N/D'
    
    # Nueva columna de Materia (MAT o LEC)
    df_excel['Materia'] = df_resultados['Materia'].apply(lambda x: 'MAT' if 'MAT' in str(x).upper() else 'LEC')
    
    # --- REDONDEO A 2 DECIMALES ---
    if col_theta_estricto:
        # Convertir a texto, limpiar comillas y comas por puntos
        theta_limpio = df_resultados[col_theta_estricto].astype(str).str.replace('"', '').str.replace(',', '.')
        # Convertir a número real (floats) y redondear a 2
        df_excel['theta.global (escala 0-100)'] = pd.to_numeric(theta_limpio, errors='coerce').round(2)
    else:
        df_excel['theta.global (escala 0-100)'] = 'N/D'

    nombre_archivo = f"Resultados_CE_{TARGET_SCHOOL_CODE}.xlsx"
    ruta_completa = os.path.join(OUTPUT_DIR, nombre_archivo)
    
    try:
        df_excel.to_excel(ruta_completa, index=False)
        print(f"✅ ¡Archivo Excel creado exitosamente!\n   Ruta: {ruta_completa}\n")
    except Exception as e:
        print(f"❌ Error al crear el Excel: {e}\n")

def generar_reportes_por_materia(df_resultados):
    """Filtra nulos, omite 3er grado, grafica y exporta las imágenes a PNG."""
    if df_resultados.empty:
        return
        
    print("=" * 80)
    print("📊 PROCESANDO DATOS PARA GRÁFICOS Y EXPORTANDO IMÁGENES")
    print("=" * 80)

    # Asegurar que el directorio existe para guardar las imágenes
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Búsqueda ESTRICTA para la columna de puntaje específico
    col_theta = next((c for c in df_resultados.columns if 'theta.global (escala 0-100)' in str(c).lower()), None)
    col_grado = next((c for c in df_resultados.columns if 'grado' in str(c).lower()), None)

    if not col_theta:
        print("❌ Error Crítico: No se encontró la columna 'theta.global (escala 0-100)' para graficar.")
        return
        
    if not col_grado:
        print("❌ Error Crítico: No se encontró la columna 'Grado' para graficar.")
        return

    df = df_resultados.copy()
    
    # Limpieza matemática y redondeo a 2 decimales para la lógica de gráficas
    df[col_theta] = df[col_theta].astype(str).str.replace('"', '').str.replace(',', '.')
    df[col_theta] = pd.to_numeric(df[col_theta], errors='coerce').round(2)
    
    df['Materia_Limpia'] = df['Materia'].apply(lambda x: 'Matemática' if 'MAT' in str(x).upper() else 'Lengua')

    # Filtrar vacíos y quitar tercer grado
    df = df.dropna(subset=[col_theta, col_grado, 'Materia_Limpia'])
    df = df[~df[col_grado].astype(str).str.contains('3', regex=True)]
    
    # Aplicar la clasificación de colores
    df['Categoria'] = df[col_theta].apply(obtener_clasificacion)
    
    materias_a_evaluar = ['Lengua', 'Matemática']
    
    for materia in materias_a_evaluar:
        df_mat = df[df['Materia_Limpia'] == materia]
        
        if df_mat.empty:
            continue
            
        print(f"\n{'-'*50}")
        print(f"📋 ESTUDIANTES VÁLIDOS EN GRÁFICO: {materia.upper()}")
        print(f"{'-'*50}")
        
        conteos_por_grado = df_mat[col_grado].value_counts().sort_index()
        total_materia = conteos_por_grado.sum()
        
        for grado, cantidad in conteos_por_grado.items():
            print(f"   > {grado}: {cantidad} estudiantes")
        print(f"   > TOTAL: {total_materia} estudiantes\n")
        
        conteo_cat = df_mat.groupby([col_grado, 'Categoria']).size().unstack(fill_value=0)
        porcentajes = conteo_cat.div(conteo_cat.sum(axis=1), axis=0) * 100
        
        categorias_ordenadas = [n['nombre'] for n in NIVELES_CONFIG]
        for cat in categorias_ordenadas:
            if cat not in porcentajes.columns:
                porcentajes[cat] = 0.0
                
        porcentajes = porcentajes[categorias_ordenadas]
        
        fig, ax = plt.subplots(figsize=(10, 6))
        colores = [n['color'] for n in NIVELES_CONFIG]
        text_colors = {n['nombre']: n['texto'] for n in NIVELES_CONFIG}
        
        porcentajes.plot(kind='barh', stacked=True, color=colores, ax=ax, width=0.8, edgecolor='white')
        
        ax.set_title(f'Distribución de Niveles por Grado - {materia} (CE {TARGET_SCHOOL_CODE})', fontsize=16, fontweight='bold', color='#2c3e50', pad=20)
        ax.set_xlabel('Porcentaje de Estudiantes (%)', fontsize=12)
        ax.set_ylabel('Grado', fontsize=12)
        ax.set_xlim(0, 100)
        ax.xaxis.set_major_formatter(lambda x, pos: f"{int(x)}%")
        ax.invert_yaxis() 
        
        for c in ax.containers:
            cat_name = c.get_label()
            t_color = text_colors.get(cat_name, 'black')
            labels = [f'{w:.1f}%' if w > 3 else '' for w in c.datavalues]
            ax.bar_label(c, labels=labels, label_type='center', color=t_color, fontsize=10, fontweight='bold')

        ax.legend(title='Niveles de Desempeño', bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        plt.tight_layout()

        # --- GUARDAR GRÁFICO EN PNG ---
        nombre_imagen = f"Grafico_Niveles_{materia}_CE_{TARGET_SCHOOL_CODE}.png"
        ruta_imagen = os.path.join(OUTPUT_DIR, nombre_imagen)
        plt.savefig(ruta_imagen, dpi=300, bbox_inches='tight')
        print(f"🖼️ ¡Imagen PNG guardada exitosamente!\n   Ruta: {ruta_imagen}")

        # Mostrar en pantalla (opcional)
        plt.show()

if __name__ == "__main__":
    df_escuela = recopilar_datos_escuela()
    
    if not df_escuela.empty:
        exportar_a_excel(df_escuela)
        generar_reportes_por_materia(df_escuela)