import os
import sys
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go
import unicodedata
import numpy as np

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
# Cambia SOLO estas dos líneas cada mes nuevo:
DIRECTORIO_MES_ACTUAL = (
    r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados",
    "Mes 3 (Mayo)"
)

PATH_METADATA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_REQ027   = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\REQ_027_1.xlsx"

# Carpeta de salida para gráficas y Excel finales
PATH_OUTPUT_PLOTS = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 3"
os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)

# Carpeta fija de cachés — NO cambiar entre meses
PATH_CACHE = os.path.join(
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual",
    "_Cache_Datos_Por_Grafica"
)
os.makedirs(PATH_CACHE, exist_ok=True)

# Archivos de caché (uno por tipo de gráfica)
CACHE_SCORES        = os.path.join(PATH_CACHE, "cache_scores_individuales.xlsx")
CACHE_SANKEY        = os.path.join(PATH_CACHE, "cache_sankey.xlsx")
CACHE_BARRAS_GRUPOS = os.path.join(PATH_CACHE, "cache_barras_grupos_grados.xlsx")
CACHE_BARRAS_TODOS  = os.path.join(PATH_CACHE, "cache_barras_todos_grados.xlsx")
CACHE_CONTEOS       = os.path.join(PATH_CACHE, "cache_conteos.xlsx")
# Archivo auxiliar que guarda el orden cronológico de los meses
CACHE_ORDEN_MESES   = os.path.join(PATH_CACHE, "cache_orden_meses.txt")

# ==========================================
# 2. COLORES Y CONSTANTES
# ==========================================
COL_SCORE_EXACTA = 'theta.global (escala 0-100)'

NIVELES          = ['Excelente', 'Bueno', 'Medio', 'Bajo', 'Crítico']
ORDEN_CATEGORIAS = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
COLOR_CATEGORIAS = {
    "Crítico":   "#991b1b",
    "Bajo":      "#ff8c2e",
    "Medio":     "#facc15",
    "Bueno":     "#84cc16",
    "Excelente": "#065f46"
}

# Paleta de colores para meses (se extiende automáticamente si hay más de 5)
_PALETA_MESES_BASE = ["#80b1d3", "#fb8072", "#b3de69", "#c9b2e8", "#f4a261",
                       "#8dd3c7", "#fdb462", "#bc80bd", "#ccebc5", "#ffed6f"]

COLORES_NODO = ['#065f46', '#84cc16', '#facc15', '#ff8c2e', '#991b1b']
COLORES_LINK = [
    'rgba(6, 95, 70, 0.4)',
    'rgba(132, 204, 22, 0.4)',
    'rgba(250, 204, 21, 0.4)',
    'rgba(255, 140, 46, 0.4)',
    'rgba(153, 27, 27, 0.4)'
]

GRUPOS_GRADOS = {
    "2° y 3° Grado":     [2, 3],
    "4°, 5° y 6° Grado": [4, 5, 6],
    "7°, 8° y 9° Grado": [7, 8, 9],
    "10° y 11° Grado":   [10, 11],
}


def color_mes(idx):
    """Devuelve un color consistente dado el índice 0-based del mes."""
    return _PALETA_MESES_BASE[idx % len(_PALETA_MESES_BASE)]


# ==========================================
# 3. CLASIFICACIÓN DE PUNTAJE
# ==========================================
def clasificar_puntaje(val):
    if pd.isna(val): return ''
    val = float(val)
    if val <= 35:   return 'Crítico'
    elif val <= 45: return 'Bajo'
    elif val <= 55: return 'Medio'
    elif val <= 65: return 'Bueno'
    else:           return 'Excelente'

def extraer_numero_grado(grado_str):
    if pd.isna(grado_str): return None
    match = re.search(r'\d+', str(grado_str).lower())
    return int(match.group()) if match else None


# ==========================================
# 4. UTILIDADES DE CACHÉ
# ==========================================
def leer_orden_meses():
    """Lee el orden cronológico de meses guardado en el archivo de texto."""
    if os.path.exists(CACHE_ORDEN_MESES):
        with open(CACHE_ORDEN_MESES, 'r', encoding='utf-8') as f:
            return [l.strip() for l in f if l.strip()]
    return []

def guardar_orden_meses(orden):
    """Persiste el orden cronológico de meses."""
    with open(CACHE_ORDEN_MESES, 'w', encoding='utf-8') as f:
        f.write('\n'.join(orden))

def _leer_cache(path):
    if os.path.exists(path):
        try:
            return pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"  [!] No se pudo leer caché {os.path.basename(path)}: {e}")
    return pd.DataFrame()

def _guardar_cache(df, path):
    try:
        df.to_excel(path, index=False)
        print(f"  [OK] Caché actualizado: {os.path.basename(path)}")
    except Exception as e:
        print(f"  [!] Error guardando caché {os.path.basename(path)}: {e}")


# ==========================================
# 5. LECTURA DE CSVs (solo mes actual)
# ==========================================
def preparar_entorno():
    sns.set_theme(style="whitegrid")
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14

def cargar_matricula():
    print(f"\n[*] Cargando Matrícula desde: {PATH_METADATA}")
    if not os.path.exists(PATH_METADATA):
        print("  [!] ERROR: No se encontró el archivo de matrícula.")
        return pd.DataFrame()
    try:
        df_mat = pd.read_csv(PATH_METADATA, sep=';', dtype=str,
                             encoding_errors='ignore')
        df_mat.columns = df_mat.columns.str.strip().str.upper()
        if 'NIE' not in df_mat.columns:
            print("  [!] ERROR: columna 'NIE' no encontrada.")
            return pd.DataFrame()
        df_mat['NIE_Limpio'] = (df_mat['NIE'].astype(str)
                                .str.replace(r'\.0$', '', regex=True)
                                .str.strip())
        df_mat['GRUPO_Limpio'] = (
            df_mat['GRUPO'].astype(str).str.strip().str.upper()
            if 'GRUPO' in df_mat.columns else 'DESCONOCIDO')
        return df_mat
    except Exception as e:
        print(f"  [!] Error leyendo matrícula: {e}")
        return pd.DataFrame()

def cargar_req027():
    print(f"\n[*] Cargando REQ_027_1 desde: {PATH_REQ027}")
    if not os.path.exists(PATH_REQ027):
        print("  [!] ADVERTENCIA: REQ_027_1.xlsx no encontrado — "
              "CÓDIGO_SECCIÓN quedará vacío.")
        return pd.Series(dtype=str)
    try:
        df = pd.read_excel(PATH_REQ027, dtype=str)
    except Exception as e:
        print(f"  [!] Error leyendo REQ_027_1.xlsx: {e}")
        return pd.Series(dtype=str)
    df.columns = df.columns.str.strip().str.upper()
    if 'GROUP_ID' not in df.columns:
        return pd.Series(dtype=str)
    nie_col = next(
        (c for c in df.columns
         if c in ('NIE', 'DOCUMENTO', 'ID_ESTUDIANTE', 'STUDENT_ID', 'DOC', 'NID')),
        None)
    if nie_col is None:
        candidates = [c for c in df.columns if c != 'GROUP_ID']
        nie_col = candidates[0] if candidates else None
    if nie_col is None:
        return pd.Series(dtype=str)
    df['_KEY'] = (df[nie_col].astype(str)
                  .str.replace(r'\.0$', '', regex=True)
                  .str.replace('"', '', regex=False)
                  .str.replace("'", '', regex=False)
                  .str.strip())
    df['GROUP_ID'] = (df['GROUP_ID'].astype(str)
                      .str.replace('"', '', regex=False).str.strip())
    df = df[df['_KEY'].str.len() > 0]
    df = df[df['GROUP_ID'].str.len() > 0]
    df = df[df['GROUP_ID'].str.upper() != 'NAN']
    return df.drop_duplicates(subset='_KEY', keep='last').set_index('_KEY')['GROUP_ID']

def _aplicar_group_id(df_out, lookup_req027):
    if lookup_req027.empty:
        return df_out
    nie_col = next(
        (c for c in df_out.columns if c.upper() in ('NIE', 'NIE_LIMPIO')), None)
    if nie_col is None:
        return df_out
    nie_key = (df_out[nie_col].astype(str)
               .str.replace(r'\.0$', '', regex=True).str.strip())
    df_out['CÓDIGO_SECCIÓN'] = nie_key.map(lookup_req027).fillna('').values
    return df_out

def extraer_csvs_puntajes_mes(ruta, etiqueta_mes):
    """Lee los CSVs de UN mes y devuelve el DataFrame de puntajes individuales."""
    lista_dfs = []
    if not os.path.exists(ruta):
        print(f"  [!] Directorio no encontrado: {ruta}")
        return pd.DataFrame()
    for f in glob.glob(os.path.join(ruta, "*.csv")):
        if 'legend' in os.path.basename(f).lower():
            continue
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            materia = ('Matemática' if 'MAT' in os.path.basename(f).upper()
                       else 'Lengua')
            if COL_SCORE_EXACTA not in df.columns:
                continue
            col_anular = next((c for c in df.columns if 'anular'   in c.lower()), None)
            col_grado  = next((c for c in df.columns if 'grado'    in c.lower()), None)
            col_codigo = next(
                (c for c in df.columns
                 if 'nro de centro' in c.lower() or 'código' in c.lower()
                 or 'codigo' in c.lower()
                 or ('centro' in c.lower() and 'id' in c.lower())), None)
            col_centro = next(
                (c for c in df.columns
                 if c.lower() in ('centro', 'nombre centro', 'institución')), None)
            if not col_centro:
                col_centro = next(
                    (c for c in df.columns
                     if 'centro' in c.lower()
                     and 'nro' not in c.lower()
                     and 'código' not in c.lower()), None)
            col_grupo = next(
                (c for c in df.columns
                 if 'grupo' in c.lower() or 'sección' in c.lower()
                 or 'seccion' in c.lower()), None)
            col_doc = next(
                (c for c in df.columns
                 if 'documento' in c.lower() or 'nie' in c.lower()
                 or c.lower() == 'id'), None)
            if not (col_grado and col_codigo and col_doc):
                continue
            if col_anular:
                def excluir_por_tiempo(val):
                    if pd.isna(val): return False
                    v = ''.join(c for c in unicodedata.normalize('NFD', str(val).lower())
                                if unicodedata.category(c) != 'Mn')
                    return '5 min' in v or 'demora' in v or 'duracion' in v
                df = df[~df[col_anular].apply(excluir_por_tiempo)].copy()
            df[COL_SCORE_EXACTA] = pd.to_numeric(
                df[COL_SCORE_EXACTA].astype(str).str.replace(',', '.'),
                errors='coerce')
            df = df.dropna(subset=[COL_SCORE_EXACTA])
            if df.empty:
                continue
            df['Grado_Num']     = df[col_grado].apply(extraer_numero_grado)
            df['Grado_Str']     = df[col_grado].astype(str).str.strip()
            df['Nro de Centro'] = (df[col_codigo].astype(str)
                                   .str.replace(r'\.0$', '', regex=True).str.strip())
            df['Centro']        = (df[col_centro].astype(str).str.strip()
                                   if col_centro else "Desconocido")
            df['Grupo']         = (df[col_grupo].astype(str).str.strip()
                                   if col_grupo  else "Desconocido")
            df['Documento']     = (df[col_doc].astype(str)
                                   .str.replace(r'\.0$', '', regex=True).str.strip())
            df['Nivel_Logro']   = df[COL_SCORE_EXACTA].apply(clasificar_puntaje)
            df_final = df[[COL_SCORE_EXACTA, 'Grado_Num', 'Grado_Str', 'Nivel_Logro',
                           'Nro de Centro', 'Centro', 'Grupo', 'Documento']].copy()
            df_final['Materia'] = materia
            df_final['Mes']     = etiqueta_mes
            df_final = df_final[df_final['Nro de Centro'] != '99999']
            df_final = df_final[
                ~df_final['Centro'].str.contains('Centro Virtual', case=False, na=False)]
            lista_dfs.append(df_final)
        except Exception:
            pass
    return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()


# ==========================================
# 6. ACTUALIZACIÓN DE CACHÉS
# ==========================================
def _actualizar_cache_generico(df_nuevo, mes_label, path_cache, tipo_float_cols=None):
    """
    Lee el caché existente, elimina filas del mes actual (si ya existían),
    agrega los datos nuevos, guarda y devuelve el DataFrame acumulado con
    tipos numéricos restaurados.
    """
    df_viejo = _leer_cache(path_cache)
    if not df_viejo.empty and 'Mes' in df_viejo.columns:
        df_viejo = df_viejo[df_viejo['Mes'] != mes_label]
    df_acum = pd.concat([df_viejo, df_nuevo], ignore_index=True)
    _guardar_cache(df_acum, path_cache)
    if tipo_float_cols:
        for col in tipo_float_cols:
            if col in df_acum.columns:
                df_acum[col] = pd.to_numeric(df_acum[col], errors='coerce')
    return df_acum

def actualizar_cache_scores(df_mes, mes_label):
    print(f"\n[*] Actualizando caché de scores individuales...")
    cols = [COL_SCORE_EXACTA, 'Grado_Num', 'Grado_Str', 'Nivel_Logro',
            'Nro de Centro', 'Centro', 'Grupo', 'Documento',
            'Materia', 'Mes', 'GRUPO_Limpio']
    df_nuevo = df_mes[[c for c in cols if c in df_mes.columns]].copy()
    df_nuevo['Grado_Num'] = df_nuevo['Grado_Num'].astype(str)
    return _actualizar_cache_generico(
        df_nuevo, mes_label, CACHE_SCORES,
        tipo_float_cols=[COL_SCORE_EXACTA, 'Grado_Num'])

def actualizar_cache_sankey(df_mes, mes_label):
    """Guarda promedios de puntaje por centro/materia para el Sankey."""
    print(f"\n[*] Actualizando caché Sankey...")
    registros = []
    for materia in df_mes['Materia'].unique():
        df_b1 = df_mes[(df_mes['Materia'] == materia) &
                       (df_mes['GRUPO_Limpio'] == 'B1')]
        if df_b1.empty:
            continue
        prom = (df_b1.groupby(['Nro de Centro', 'Centro'])
                [COL_SCORE_EXACTA].mean().reset_index())
        prom['Nivel']   = prom[COL_SCORE_EXACTA].apply(clasificar_puntaje)
        prom['Materia'] = materia
        prom['Mes']     = mes_label
        registros.append(prom)
    if not registros:
        print("  [!] Sin datos B1 para el caché Sankey.")
        df_cache = _leer_cache(CACHE_SANKEY)
        df_cache[COL_SCORE_EXACTA] = pd.to_numeric(
            df_cache.get(COL_SCORE_EXACTA, pd.Series()), errors='coerce')
        return df_cache
    df_nuevo = pd.concat(registros, ignore_index=True)
    return _actualizar_cache_generico(
        df_nuevo, mes_label, CACHE_SANKEY,
        tipo_float_cols=[COL_SCORE_EXACTA])

def actualizar_cache_barras_grupos(df_mes, mes_label):
    """Porcentaje de cada nivel por grado/materia/grupo."""
    print(f"\n[*] Actualizando caché de barras por grupos de grados...")
    registros = []
    for grupo in ['B1', 'B2']:
        df_g = df_mes[df_mes['GRUPO_Limpio'] == grupo].copy()
        if df_g.empty:
            continue
        for materia in df_g['Materia'].unique():
            df_m = df_g[df_g['Materia'] == materia].copy()
            df_m['Grado_Str'] = df_m['Grado_Num'].astype(str) + "° Grado"
            ct = (pd.crosstab(df_m['Grado_Str'], df_m['Nivel_Logro'],
                              normalize='index') * 100).reset_index()
            ct['Grado_Num']    = ct['Grado_Str'].str.extract(r'(\d+)').astype(float)
            ct['Materia']      = materia
            ct['Mes']          = mes_label
            ct['GRUPO_Limpio'] = grupo
            registros.append(ct)
    if not registros:
        return _leer_cache(CACHE_BARRAS_GRUPOS)
    df_nuevo = pd.concat(registros, ignore_index=True)
    df_acum  = _actualizar_cache_generico(
        df_nuevo, mes_label, CACHE_BARRAS_GRUPOS,
        tipo_float_cols=['Grado_Num'])
    for cat in ORDEN_CATEGORIAS:
        if cat not in df_acum.columns:
            df_acum[cat] = 0.0
        df_acum[cat] = pd.to_numeric(df_acum[cat], errors='coerce').fillna(0)
    return df_acum

def actualizar_cache_barras_todos(df_mes, mes_label):
    """Porcentaje de cada nivel por grado (todos) + fila Global."""
    print(f"\n[*] Actualizando caché de barras todos los grados...")
    registros = []
    for grupo in ['B1', 'B2']:
        df_g = df_mes[df_mes['GRUPO_Limpio'] == grupo].copy()
        if df_g.empty:
            continue
        for materia in df_g['Materia'].unique():
            df_m = df_g[df_g['Materia'] == materia].copy()
            df_m['Grado_Str'] = df_m['Grado_Num'].astype(str) + "° Grado"
            ct_g = (pd.crosstab(df_m['Grado_Str'], df_m['Nivel_Logro'],
                                normalize='index') * 100).reset_index()
            ct_g['Grado_Num'] = ct_g['Grado_Str'].str.extract(r'(\d+)').astype(float)
            ct_glob = (df_m['Nivel_Logro'].value_counts(normalize=True)
                       .to_frame().T * 100)
            ct_glob.index       = ["Global (Todos)"]
            ct_glob             = ct_glob.reset_index().rename(
                columns={'index': 'Grado_Str'})
            ct_glob['Grado_Num'] = 999
            ct = pd.concat([ct_g, ct_glob], ignore_index=True)
            ct['Materia']      = materia
            ct['Mes']          = mes_label
            ct['GRUPO_Limpio'] = grupo
            registros.append(ct)
    if not registros:
        return _leer_cache(CACHE_BARRAS_TODOS)
    df_nuevo = pd.concat(registros, ignore_index=True)
    df_acum  = _actualizar_cache_generico(
        df_nuevo, mes_label, CACHE_BARRAS_TODOS,
        tipo_float_cols=['Grado_Num'])
    for cat in ORDEN_CATEGORIAS:
        if cat not in df_acum.columns:
            df_acum[cat] = 0.0
        df_acum[cat] = pd.to_numeric(df_acum[cat], errors='coerce').fillna(0)
    return df_acum

def actualizar_cache_conteos(df_mes, mes_label):
    """Conteo de estudiantes únicos por grado/materia/grupo."""
    print(f"\n[*] Actualizando caché de conteos...")
    grados_esp = list(range(2, 12))
    registros  = []
    for grupo in ['B1', 'B2', 'TODOS']:
        df_g = df_mes if grupo == 'TODOS' else df_mes[df_mes['GRUPO_Limpio'] == grupo]
        for materia in ['Matemática', 'Lengua']:
            df_m = df_g[df_g['Materia'] == materia]
            if df_m.empty:
                continue
            cnt = (df_m.groupby('Grado_Num')['Documento']
                   .nunique().reset_index().rename(columns={'Documento': 'Conteo'}))
            for g in grados_esp:
                if g not in cnt['Grado_Num'].values:
                    cnt = pd.concat(
                        [cnt, pd.DataFrame({'Grado_Num': [g], 'Conteo': [0]})],
                        ignore_index=True)
            cnt['Materia']      = materia
            cnt['Mes']          = mes_label
            cnt['GRUPO_Limpio'] = grupo
            registros.append(cnt)
    if not registros:
        return _leer_cache(CACHE_CONTEOS)
    df_nuevo = pd.concat(registros, ignore_index=True)
    df_acum  = _actualizar_cache_generico(
        df_nuevo, mes_label, CACHE_CONTEOS,
        tipo_float_cols=['Grado_Num', 'Conteo'])
    df_acum['Conteo'] = df_acum['Conteo'].fillna(0).astype(int)
    return df_acum


# ==========================================
# 7. CONTEOS Y ALERTAS
# ==========================================
def generar_reportes_conteos_separados(df_conteos, meses_ordenados):
    print("\n[*] Generando Reportes de Totales de Estudiantes...")
    grados_esp = list(range(2, 12))

    def crear_hoja(grupo, materia):
        df = df_conteos[(df_conteos['GRUPO_Limpio'] == grupo) &
                        (df_conteos['Materia'] == materia)]
        if df.empty:
            return pd.DataFrame()
        pivot = df.pivot_table(index='Mes', columns='Grado_Num',
                               values='Conteo', aggfunc='sum', fill_value=0)
        pivot = pivot.reindex(columns=grados_esp, fill_value=0)
        pivot.columns = [f"{int(c)}°" for c in pivot.columns]
        pivot = pivot.reindex(meses_ordenados).fillna(0)
        pivot['Total por prueba'] = pivot.sum(axis=1)
        return pivot

    def exportar(grupo, filename):
        p = os.path.join(PATH_OUTPUT_PLOTS, filename)
        with pd.ExcelWriter(p, engine='openpyxl') as writer:
            for mat in ['Matemática', 'Lengua']:
                h = crear_hoja(grupo, mat)
                if not h.empty:
                    h.to_excel(writer, sheet_name=mat)
                else:
                    pd.DataFrame({'Aviso': [f'Sin datos para {mat}']}).to_excel(
                        writer, sheet_name=mat, index=False)
        print(f"  [OK] Guardado: {filename}")

    exportar('TODOS', "Conteo_Estudiantes_B1_y_B2.xlsx")
    exportar('B1',    "Conteo_Estudiantes_Solo_B1.xlsx")
    exportar('B2',    "Conteo_Estudiantes_Solo_B2.xlsx")

def alertar_b2_mes1_mes2(df_scores, meses_ordenados):
    print("\n[*] Verificando escuelas B2 con participación en meses previos...")
    if len(meses_ordenados) < 2:
        print("  [-] Menos de 2 meses acumulados — sin alerta.")
        return
    previos = meses_ordenados[:-1]
    df_alerta = df_scores[
        (df_scores['GRUPO_Limpio'] == 'B2') &
        (df_scores['Mes'].isin(previos))]
    if not df_alerta.empty:
        print("  [!] ATENCIÓN: Escuelas B2 con participación en meses previos:")
        cnt = (df_alerta.groupby(['Nro de Centro', 'Centro', 'Materia', 'Mes'])
               ['Documento'].nunique().reset_index())
        for _, r in cnt.iterrows():
            print(f"      - {r['Nro de Centro']} - {r['Centro']} "
                  f"| {r['Materia']} | {r['Mes']}: {r['Documento']} estudiante(s)")
    else:
        print("  [OK] Ninguna escuela B2 fue evaluada en meses previos.")


# ==========================================
# 8. SANKEY  (desde caché)
# ==========================================
def calc_y_centers(df_mes):
    """Calcula posiciones Y de los nodos según el conteo de centros por nivel."""
    totals   = [len(df_mes[df_mes['Nivel'] == n]) for n in NIVELES]
    pad_frac = 0.04
    usable   = 1.0 - pad_frac * (len(NIVELES) - 1)
    grand    = sum(totals) or 1
    heights  = [max((v / grand) * usable, 0.01) for v in totals]
    total_h  = sum(heights)
    if total_h > usable:
        heights = [h * usable / total_h for h in heights]
    y_centers, cursor = [], 0.0
    for h in heights:
        y_centers.append(round(cursor + h / 2, 4))
        cursor += h + pad_frac
    return y_centers

def dibujar_sankey_multimes(df_sankey_cache, materia, meses_ordenados):
    """
    Dibuja el diagrama Sankey de trayectorias por nivel (centros educativos, B1).
    df_sankey_cache: DataFrame acumulado con columnas
        [Nro de Centro, Centro, theta.global, Nivel, Materia, Mes]
    meses_ordenados: lista de etiquetas en orden cronológico
    """
    print(f"  -> Dibujando Sankey para {materia}...")
    df_mat = df_sankey_cache[df_sankey_cache['Materia'] == materia].copy()
    if df_mat.empty:
        print(f"     [!] Sin datos en caché Sankey para {materia}")
        return

    # Construir lista de DataFrames por mes (solo los que tienen datos)
    meses_validos, nombres_validos = [], []
    for mes in meses_ordenados:
        df_m = df_mat[df_mat['Mes'] == mes].copy()
        if df_m.empty:
            continue
        # Renombrar para uso interno
        df_m = df_m[['Nro de Centro', 'Centro', 'Nivel']].rename(
            columns={'Nro de Centro': 'Código'})
        meses_validos.append(df_m)
        nombres_validos.append(mes)

    if len(meses_validos) < 2:
        print(f"     [!] Se necesitan al menos 2 meses para el Sankey (hay {len(meses_validos)}).")
        return

    # Estadísticas inicio→fin
    df_stats = pd.merge(
        meses_validos[0][['Código', 'Nivel']],
        meses_validos[-1][['Código', 'Nivel']],
        on='Código', suffixes=('_Ini', '_Fin'), how='inner')
    n_centros = len(df_stats)
    mantienen = suben = bajan = 0
    for _, row in df_stats.iterrows():
        if row['Nivel_Ini'] not in NIVELES or row['Nivel_Fin'] not in NIVELES:
            continue
        d = NIVELES.index(row['Nivel_Fin']) - NIVELES.index(row['Nivel_Ini'])
        if d == 0:  mantienen += 1
        elif d < 0: suben     += 1
        else:       bajan     += 1
    subtitulo = (
        f"<span style='font-size:13px;color:#888'>"
        f"Población base (Grupo B1): {n_centros} centros<br>"
        f"Mantienen: {mantienen} ({mantienen/max(n_centros,1)*100:.1f}%) · "
        f"Suben: {suben} ({suben/max(n_centros,1)*100:.1f}%) · "
        f"Bajan: {bajan} ({bajan/max(n_centros,1)*100:.1f}%)</span>"
    ) if n_centros > 0 else ""

    # Construir nodos
    labels_nodos, colores_nodos, node_x, node_y = [], [], [], []
    for i, (nombre, df_mes) in enumerate(zip(nombres_validos, meses_validos)):
        x_val = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        ys    = calc_y_centers(df_mes)
        for j, nivel in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {nivel}")
            colores_nodos.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(ys[j])

    # Construir links
    sources, targets, values, link_colors = [], [], [], []

    # Links invisibles para forzar estructura (garantizan que todos los nodos aparecen)
    for i in range(len(meses_validos) - 1):
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(oo + j); targets.append(od + j)
            values.append(1e-9);    link_colors.append('rgba(0,0,0,0)')

    # Links reales de transición
    for i in range(len(meses_validos) - 1):
        df_cruce = pd.merge(
            meses_validos[i][['Código', 'Centro', 'Nivel']],
            meses_validos[i + 1][['Código', 'Nivel']],
            on='Código', suffixes=('_A', '_S'), how='inner')
        if df_cruce.empty:
            continue

        # Alerta Bueno→Bajo en Matemática primer paso
        if materia == "Matemática" and i == 0:
            alerta = df_cruce[
                (df_cruce['Nivel_A'] == 'Bueno') &
                (df_cruce['Nivel_S'] == 'Bajo')]
            if not alerta.empty:
                print(f"\n      [ALERTA] Escuelas Bueno→Bajo Matemática "
                      f"({nombres_validos[i]}→{nombres_validos[i+1]}):")
                for _, r in alerta.iterrows():
                    print(f"         - {r['Código']} | {r['Centro']}")

        flujos = (df_cruce.groupby(['Nivel_A', 'Nivel_S'])
                  .size().reset_index(name='Cantidad'))
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        mo = {n: j + oo for j, n in enumerate(NIVELES)}
        md = {n: j + od for j, n in enumerate(NIVELES)}
        for _, row in flujos.iterrows():
            if row['Cantidad'] == 0: continue
            si = mo[row['Nivel_A']]; ti = md[row['Nivel_S']]
            sources.append(si); targets.append(ti)
            values.append(row['Cantidad'])
            link_colors.append(COLORES_LINK[si % len(NIVELES)])

    if not sources:
        return

    fig = go.Figure(data=[go.Sankey(
        arrangement='fixed',
        node=dict(pad=20, thickness=24,
                  line=dict(color="white", width=0.5),
                  label=labels_nodos,
                  color=colores_nodos,
                  x=node_x, y=node_y),
        link=dict(source=sources, target=targets,
                  value=values, color=link_colors)
    )])
    fig.update_layout(
        title=dict(
            text=(f"<span style='font-size:22px;color:#1a2b4c;"
                  f"font-family:Arial,sans-serif;font-weight:bold;'>"
                  f"Trayectorias de niveles — {materia}</span><br>{subtitulo}"),
            x=0.01, y=0.95),
        font_size=12,
        width=max(900, 800 + len(meses_validos) * 150),
        height=600,
        margin=dict(l=20, r=20, t=100, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    nombre_mat = materia.replace('á', 'a').replace('é', 'e')
    fig.write_html(os.path.join(PATH_OUTPUT_PLOTS,
                                f"Sankey_Evolutivo_{nombre_mat}.html"))
    try:
        fig.write_image(
            os.path.join(PATH_OUTPUT_PLOTS,
                         f"Sankey_Evolutivo_{nombre_mat}.png"), scale=2)
    except Exception:
        pass
    print(f"     [OK] Sankey {materia} guardado.")


# ==========================================
# 9. BARRAS APILADAS  (desde caché)
# ==========================================
def generar_barras_por_grupo_grados(df_cache, meses_ordenados, sufijo="B1"):
    print(f"\n[*] Generando Barras por Grupo de Grados ({sufijo})...")
    df_g = df_cache[df_cache['GRUPO_Limpio'] == sufijo].copy()
    if df_g.empty:
        return

    for nombre_grupo, grados in GRUPOS_GRADOS.items():
        df_grp = df_g[df_g['Grado_Num'].isin(grados)]
        if df_grp.empty:
            continue
        grados_orden = [f"{g}° Grado"
                        for g in sorted(grados, reverse=True)
                        if g in df_grp['Grado_Num'].values]

        for materia in df_grp['Materia'].unique():
            df_mat = df_grp[df_grp['Materia'] == materia]
            meses_datos = [m for m in meses_ordenados
                           if not df_mat[df_mat['Mes'] == m].empty]
            if not meses_datos: continue

            n = len(meses_datos)
            fig, axes = plt.subplots(1, n, figsize=(max(9, n * 9), 7), sharey=True)
            if n == 1: axes = [axes]

            for i, mes in enumerate(meses_datos):
                ax    = axes[i]
                df_ms = df_mat[df_mat['Mes'] == mes].set_index('Grado_Str')
                if df_ms.empty:
                    ax.set_visible(False); continue
                ct = df_ms.reindex(columns=ORDEN_CATEGORIAS).fillna(0)
                ct = ct.reindex(grados_orden)
                ct.plot(kind='barh', stacked=True, ax=ax,
                        color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS],
                        legend=False, edgecolor='white', width=0.6)
                prueba_num = meses_ordenados.index(mes) + 1
                mes_nombre = mes.split('(')[-1].replace(')', '').strip()
                ax.set_title(f"Prueba {prueba_num}\n({mes_nombre})",
                             fontsize=16, color='#005288', pad=14, fontweight='bold')
                ax.set_xlim(0, 100)
                ax.set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)
                ax.set_ylabel("Grado" if i == 0 else "")
                if i == n - 1:
                    h, l = ax.get_legend_handles_labels()
                    ax.legend(h, l, title="Nivel de Logro",
                              bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)
                for p in ax.patches:
                    w = p.get_width()
                    if w > 4:
                        ax.text(p.get_x() + w / 2, p.get_y() + p.get_height() / 2,
                                f"{w:.1f}%", ha='center', va='center',
                                color='white', fontsize=9, fontweight='bold')

            plt.suptitle(f"Niveles de Logro — {sufijo} — {materia}\n{nombre_grupo}",
                         fontsize=16, fontweight='bold', color='#1a2b4c')
            plt.tight_layout(rect=[0, 0, 0.88, 1])
            seg = nombre_grupo.replace('°', '').replace(' ', '_').replace(',', '')
            fname = f"Barras_{sufijo}_{materia}_{seg}.png"
            plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, fname),
                        dpi=300, bbox_inches='tight')
            plt.close()
            print(f"    [OK] {fname}")

def generar_barras_todos_los_grados(df_cache, meses_ordenados, sufijo="B1"):
    print(f"\n[*] Generando Barras (Todos los Grados) ({sufijo})...")
    df_g = df_cache[df_cache['GRUPO_Limpio'] == sufijo].copy()
    if df_g.empty:
        return

    for materia in df_g['Materia'].unique():
        df_mat = df_g[df_g['Materia'] == materia]
        meses_datos = [m for m in meses_ordenados
                       if not df_mat[df_mat['Mes'] == m].empty]
        if not meses_datos: continue

        grados_num  = sorted(df_mat['Grado_Num'].dropna().unique(), reverse=True)
        grados_ord  = [f"{int(g)}° Grado" for g in grados_num if g != 999]
        grados_ord.append("Global (Todos)")

        n = len(meses_datos)
        fig, axes = plt.subplots(1, n, figsize=(max(9, n * 9), 9), sharey=True)
        if n == 1: axes = [axes]

        for i, mes in enumerate(meses_datos):
            ax    = axes[i]
            df_ms = df_mat[df_mat['Mes'] == mes].set_index('Grado_Str')
            if df_ms.empty:
                ax.set_visible(False); continue
            ct = df_ms.reindex(columns=ORDEN_CATEGORIAS).fillna(0)
            orden_actual = [g for g in grados_ord if g in ct.index]
            ct = ct.reindex(orden_actual)
            ct.plot(kind='barh', stacked=True, ax=ax,
                    color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS],
                    legend=False, edgecolor='white', width=0.7)
            prueba_num = meses_ordenados.index(mes) + 1
            mes_nombre = mes.split('(')[-1].replace(')', '').strip()
            ax.set_title(f"Prueba {prueba_num}\n({mes_nombre})",
                         fontsize=16, color='#005288', pad=14, fontweight='bold')
            ax.set_xlim(0, 100)
            ax.set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)
            ax.set_ylabel("Grado" if i == 0 else "")
            if i == n - 1:
                h, l = ax.get_legend_handles_labels()
                ax.legend(h, l, title="Nivel de Logro",
                          bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)
            for p in ax.patches:
                w = p.get_width()
                if w > 4:
                    ax.text(p.get_x() + w / 2, p.get_y() + p.get_height() / 2,
                            f"{w:.1f}%", ha='center', va='center',
                            color='white', fontsize=9, fontweight='bold')

        plt.suptitle(f"Niveles de Logro — {sufijo} — {materia}\n"
                     f"(Vista Completa Todos los Grados)",
                     fontsize=16, fontweight='bold', color='#1a2b4c')
        plt.tight_layout(rect=[0, 0, 0.88, 1])
        fname = f"Barras_{sufijo}_{materia}_TodosLosGrados.png"
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, fname),
                    dpi=300, bbox_inches='tight')
        plt.close()
        print(f"    [OK] {fname}")


# ==========================================
# 10. KDE + BOXPLOT  (desde caché)
# ==========================================
def generar_kde_y_boxplot(df_scores, meses_ordenados, sufijo="B1",
                          color_list=None):
    """
    KDE y Boxplot+Jitter.
    - df_scores: caché completo de scores individuales
    - meses_ordenados: lista de etiquetas en orden cronológico
    - sufijo: 'B1' o 'B2'
    - color_list: si None, se asignan colores automáticamente por posición en
      meses_ordenados (consistente entre runs)
    """
    print(f"\n[*] Generando KDE y Boxplot+Jitter ({sufijo})...")
    df_g = df_scores[df_scores['GRUPO_Limpio'] == sufijo]

    # Asignar color a cada mes según su posición global (consistente entre runs)
    if color_list is not None:
        # Modo legado: lista fija (ej. B2 usa un solo color)
        palette = {mes: color_list[i % len(color_list)]
                   for i, mes in enumerate(meses_ordenados)}
    else:
        palette = {mes: color_mes(i) for i, mes in enumerate(meses_ordenados)}

    for materia in df_g['Materia'].unique():
        df_sub = df_g[df_g['Materia'] == materia].copy()
        print(f"    - {materia}...")

        meses_con_datos = [m for m in meses_ordenados
                           if not df_sub[df_sub['Mes'] == m].empty]
        if not meses_con_datos:
            continue

        # ── KDE ──────────────────────────────────────────────────────────────
        plt.figure(figsize=(12, 7))
        sns.histplot(
            data=df_sub[df_sub['Mes'].isin(meses_con_datos)],
            x=COL_SCORE_EXACTA, hue='Mes',
            hue_order=meses_con_datos,
            kde=True, element="step", stat="probability",
            common_norm=False,
            palette={m: palette[m] for m in meses_con_datos},
            alpha=0.6)
        titulo = " vs ".join(
            [m.split(' ')[-1].replace('(', '').replace(')', '')
             for m in meses_con_datos])
        plt.title(f"Evolución de Puntajes Theta: {materia}\n"
                  f"({titulo} — {sufijo})")
        plt.xlim(0, 100)
        nombre_mat = materia.replace('á', 'a').replace('é', 'e')
        plt.savefig(
            os.path.join(PATH_OUTPUT_PLOTS,
                         f"Distribucion_Evolutiva_{sufijo}_{nombre_mat}.png"),
            dpi=300, bbox_inches='tight')
        plt.close()

        # ── Boxplot + Jitter ──────────────────────────────────────────────────
        n_meses   = len(meses_con_datos)
        fig, axes = plt.subplots(n_meses, 1,
                                 figsize=(12, 4 * n_meses), sharex=True)
        if n_meses == 1:
            axes = [axes]

        for i, mes in enumerate(meses_con_datos):
            ax        = axes[i]
            color     = palette[mes]
            datos_box = df_sub[df_sub['Mes'] == mes][COL_SCORE_EXACTA].dropna()
            if datos_box.empty:
                ax.set_visible(False); continue

            ax.boxplot(
                datos_box, vert=False, patch_artist=True, widths=0.45,
                boxprops=dict(facecolor=color, alpha=0.45, linewidth=1.5),
                medianprops=dict(color='#1a2b4c', linewidth=2.5),
                whiskerprops=dict(linewidth=1.5, linestyle='--', color='#555555'),
                capprops=dict(linewidth=2, color='#555555'),
                flierprops=dict(marker='', linestyle='none'),
                showmeans=True,
                meanprops=dict(marker='D', markerfacecolor='white',
                               markeredgecolor='#1a2b4c', markersize=8))
            np.random.seed(42)
            jitter = np.random.uniform(-0.18, 0.18, size=len(datos_box))
            ax.scatter(datos_box, 1 + jitter, alpha=0.20, s=10,
                       color=color, edgecolors='none', zorder=3)

            prueba_num = meses_ordenados.index(mes) + 1
            mes_nombre = mes.split('(')[-1].replace(')', '').strip()
            ax.set_ylabel(f"Prueba {prueba_num}\n({mes_nombre})",
                          fontsize=11, fontweight='bold',
                          color='#005288', labelpad=10)
            ax.set_xlim(0, 100)
            ax.set_yticks([])
            ax.grid(axis='x', linestyle='--', alpha=0.4)
            ax.spines[['top', 'right', 'left']].set_visible(False)
            ax.text(0.99, 0.82,
                    f"n={len(datos_box):,}  |  "
                    f"Mediana={datos_box.median():.1f}  |  "
                    f"Media={datos_box.mean():.1f}",
                    transform=ax.transAxes, ha='right', va='top',
                    fontsize=10, color='#333333',
                    bbox=dict(boxstyle='round,pad=0.35', facecolor='white',
                              edgecolor='#cccccc', alpha=0.85))

        axes[-1].set_xlabel("Puntaje (0–100)", fontsize=12)
        fig.suptitle(
            f"Distribución por Prueba — {sufijo} — {materia}\n(todos los grados)",
            fontsize=15, fontweight='bold', color='#1a2b4c', y=1.01)
        plt.tight_layout()
        plt.savefig(
            os.path.join(PATH_OUTPUT_PLOTS,
                         f"Boxplot_Jitter_{sufijo}_{nombre_mat}.png"),
            dpi=300, bbox_inches='tight')
        plt.close()
        print(f"    [OK] Boxplot_{sufijo}_{nombre_mat}.png")


# ==========================================
# 11. EXCELS DE RESULTADOS
# ==========================================
def exportar_base_consolidada_b1(df_scores, df_matricula, lookup_req027,
                                 meses_ordenados):
    print("\n[*] Generando Base de Datos Consolidada (Grupo B1)...")
    try:
        df_b1 = df_scores[df_scores['GRUPO_Limpio'] == 'B1'].copy()
        if df_b1.empty:
            print("  [!] Sin datos B1."); return
        df_pivot = df_b1.pivot_table(
            index=['Nro de Centro', 'Centro', 'Grupo', 'Documento', 'Grado_Str'],
            columns=['Materia', 'Mes'],
            values=COL_SCORE_EXACTA, aggfunc='first')
        score_cols_raw = sorted(
            df_pivot.columns.tolist(),
            key=lambda c: (0 if 'Matemática' in c[0] else 1,
                           next((i for i, m in enumerate(meses_ordenados)
                                 if m in c[1]), 99)))
        flat_map = {c: (f"{COL_SCORE_EXACTA} (para {c[0].lower()} {c[1].lower()})")
                    for c in df_pivot.columns}
        df_pivot.columns = [flat_map[c] for c in df_pivot.columns]
        df_pivot = df_pivot.reset_index().rename(columns={'Grado_Str': 'Grado'})
        interleaved = []
        for rc in score_cols_raw:
            sc = flat_map[rc]; nc = f"Nivel {rc[0]} ({rc[1]})"
            if sc in df_pivot.columns:
                df_pivot[nc] = df_pivot[sc].apply(clasificar_puntaje)
                interleaved += [sc, nc]
        id_cols = [c for c in ['Nro de Centro', 'Centro', 'Grupo', 'Documento', 'Grado']
                   if c in df_pivot.columns]
        rest = [c for c in df_pivot.columns if c not in set(id_cols + interleaved)]
        df_pivot = df_pivot[id_cols + interleaved + rest]
        cols_sol   = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                      'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                      'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO']
        cols_exist = [c for c in cols_sol
                      if c in df_matricula.columns and c != 'CÓDIGO_SECCIÓN']
        df_out = pd.merge(df_matricula[cols_exist + ['NIE_Limpio']],
                          df_pivot, left_on='NIE_Limpio', right_on='Documento',
                          how='inner')
        df_out = df_out.drop(columns=['NIE_Limpio'], errors='ignore')
        df_out = _aplicar_group_id(df_out, lookup_req027)
        header = [c for c in cols_sol if c in df_out.columns]
        rest   = [c for c in df_out.columns if c not in set(header)]
        df_out[header + rest].to_excel(
            os.path.join(PATH_OUTPUT_PLOTS, "Base_Estudiantes_B1_Consolidada.xlsx"),
            index=False)
        print("  [OK] Excel consolidado B1 guardado.")
    except Exception as e:
        print(f"  [!] Error al exportar base B1: {e}")

def generar_excel_y_barras_b2(df_b2_mes, df_matricula, lookup_req027):
    if df_b2_mes.empty:
        print("\n  [!] Sin estudiantes B2 en el mes actual."); return
    print("\n[*] Generando Excel para Grupo B2...")
    df_pivot = df_b2_mes.pivot_table(
        index='Documento', columns='Materia',
        values=COL_SCORE_EXACTA, aggfunc='first')
    df_pivot.columns = [f"Puntaje {c}" for c in df_pivot.columns]
    df_pivot = df_pivot.reset_index()
    for bc in ['Puntaje Matemática', 'Puntaje Lengua']:
        if bc not in df_pivot.columns: df_pivot[bc] = pd.NA
        df_pivot[bc.replace('Puntaje', 'Nivel')] = df_pivot[bc].apply(clasificar_puntaje)
    cols_sol   = ['CODIGO', 'NOMBRE', 'GRADO', 'NIE',
                  'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                  'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO']
    cols_exist = [c for c in cols_sol if c in df_matricula.columns]
    df_out = pd.merge(df_matricula[cols_exist + ['NIE_Limpio']],
                      df_pivot, left_on='NIE_Limpio', right_on='Documento',
                      how='inner')
    df_out = df_out.drop(columns=['NIE_Limpio', 'Documento'], errors='ignore')
    df_out = _aplicar_group_id(df_out, lookup_req027)
    priority = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE', 'PRIMER_APELLIDO',
                'SEGUNDO_APELLIDO', 'Puntaje Matemática', 'Nivel Matemática',
                'Puntaje Lengua', 'Nivel Lengua']
    final = [c for c in priority if c in df_out.columns]
    df_out[final].to_excel(
        os.path.join(PATH_OUTPUT_PLOTS, "Resultados_MesActual_Grupo_B2.xlsx"),
        index=False)
    print("  [OK] Excel Grupo B2 guardado.")

def construir_excel_unico(df_scores, df_matricula, meses_ordenados, lookup_req027):
    print("\n[*] Construyendo Excel Maestro (hoja única, todos los meses)...")
    df_b1 = df_scores[df_scores['GRUPO_Limpio'] == 'B1'].copy()
    df_b2 = df_scores[df_scores['GRUPO_Limpio'] == 'B2'].copy()
    if len(meses_ordenados) >= 2:
        base = (df_b1[df_b1['Mes'].isin(meses_ordenados[:2])]
                [['Documento', 'Materia']].drop_duplicates())
        df_b1 = pd.merge(df_b1, base, on=['Documento', 'Materia'], how='inner')
    df_b2 = df_b2[df_b2['Mes'] == meses_ordenados[-1]]
    df_b1['Grupo_Paars'] = 'B1'; df_b2['Grupo_Paars'] = 'B2'
    df_all = pd.concat([df_b1, df_b2], ignore_index=True)
    if df_all.empty:
        print("  [!] Sin datos para el Excel Maestro."); return None
    df_pivot = df_all.pivot_table(
        index=['Documento', 'Nro de Centro', 'Centro',
               'Grado_Str', 'GRUPO_Limpio', 'Grupo_Paars'],
        columns=['Materia', 'Mes'],
        values=COL_SCORE_EXACTA, aggfunc='first')
    score_cols_raw = sorted(
        df_pivot.columns.tolist(),
        key=lambda c: (0 if 'Matemática' in c[0] else 1,
                       next((i for i, m in enumerate(meses_ordenados)
                             if m in c[1]), 99)))
    flat_map = {c: f"Puntaje {c[0]} ({c[1]})" for c in df_pivot.columns}
    df_pivot.columns = [flat_map[c] for c in df_pivot.columns]
    df_pivot = df_pivot.reset_index()
    interleaved = []
    for rc in score_cols_raw:
        sc = flat_map[rc]; nc = f"Nivel {rc[0]} ({rc[1]})"
        if sc in df_pivot.columns:
            df_pivot[nc] = df_pivot[sc].apply(clasificar_puntaje)
            interleaved += [sc, nc]
    cols_sol   = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                  'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                  'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO']
    cols_exist = [c for c in cols_sol
                  if c in df_matricula.columns and c != 'CÓDIGO_SECCIÓN']
    df_out = pd.merge(df_matricula[cols_exist + ['NIE_Limpio']],
                      df_pivot, left_on='NIE_Limpio', right_on='Documento',
                      how='inner')
    df_out = df_out.drop(columns=['NIE_Limpio'], errors='ignore')
    df_out = _aplicar_group_id(df_out, lookup_req027)
    filter_cols = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                   'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE', 'PRIMER_APELLIDO',
                   'SEGUNDO_APELLIDO', 'Grupo_Paars', 'Nro de Centro',
                   'Centro', 'Grado_Str', 'GRUPO_Limpio']
    header = [c for c in filter_cols if c in df_out.columns]
    placed = set(header + interleaved)
    rest   = [c for c in df_out.columns if c not in placed]
    df_out = df_out[[c for c in header + interleaved + rest if c in df_out.columns]]
    out = os.path.join(PATH_OUTPUT_PLOTS, "Excel_Maestro_HojaUnica.xlsx")
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name="Todos los Datos", index=False)
        ws = writer.sheets["Todos los Datos"]
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes    = "A2"
    print(f"  [OK] Excel Maestro → {out}  ({len(df_out):,} filas, "
          f"{os.path.getsize(out)/1024/1024:.1f} MB)")
    return out


# ==========================================
# 12. EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print(f"\n{'='*60}")
    print("  GENERADOR DE REPORTE EVOLUTIVO — GRUPOS B1 Y B2")
    print(f"{'='*60}")

    preparar_entorno()
    ruta_mes, etiqueta_mes = DIRECTORIO_MES_ACTUAL
    print(f"\n[*] Mes a procesar : {etiqueta_mes}")
    print(f"    Directorio     : {ruta_mes}")

    # ── 1. Matrícula ──────────────────────────────────────────────
    df_matricula = cargar_matricula()
    if df_matricula.empty:
        print("\n[!] Sin matrícula — proceso detenido."); sys.exit()
    if 'CODIGO' not in df_matricula.columns:
        print("  [!] ERROR: columna 'CODIGO' no existe."); sys.exit()

    mapa_grupos = (df_matricula[['CODIGO', 'GRUPO_Limpio']]
                   .dropna().drop_duplicates(subset=['CODIGO'])
                   .set_index('CODIGO')['GRUPO_Limpio'])
    lookup_req027 = cargar_req027()

    # ── 2. Leer CSVs del mes actual ───────────────────────────────
    print(f"\n[*] Leyendo CSVs del mes actual ({etiqueta_mes})...")
    df_scores_mes = extraer_csvs_puntajes_mes(ruta_mes, etiqueta_mes)
    if df_scores_mes.empty:
        print("\n[!] Sin puntajes válidos — proceso detenido."); sys.exit()

    df_scores_mes = pd.merge(
        df_scores_mes, df_matricula[['NIE_Limpio']],
        left_on='Documento', right_on='NIE_Limpio', how='inner')
    df_scores_mes['GRUPO_Limpio'] = (
        df_scores_mes['Nro de Centro'].map(mapa_grupos).fillna('DESCONOCIDO'))
    print(f"  [OK] {len(df_scores_mes):,} registros válidos para {etiqueta_mes}")

    # ── 3. Actualizar orden cronológico de meses ──────────────────
    orden_previo = leer_orden_meses()
    if etiqueta_mes not in orden_previo:
        orden_previo.append(etiqueta_mes)
    meses_ordenados = orden_previo          # lista con orden de inserción real
    guardar_orden_meses(meses_ordenados)
    print(f"\n[*] Meses acumulados (orden cronológico): {meses_ordenados}")

    # ── 4. Actualizar todos los cachés ────────────────────────────
    df_scores_acum       = actualizar_cache_scores(df_scores_mes, etiqueta_mes)
    df_sankey_acum       = actualizar_cache_sankey(df_scores_mes, etiqueta_mes)
    df_barras_grupos_acum = actualizar_cache_barras_grupos(df_scores_mes, etiqueta_mes)
    df_barras_todos_acum  = actualizar_cache_barras_todos(df_scores_mes, etiqueta_mes)
    df_conteos_acum      = actualizar_cache_conteos(df_scores_mes, etiqueta_mes)

    # ── 5. Conteos y alertas ──────────────────────────────────────
    generar_reportes_conteos_separados(df_conteos_acum, meses_ordenados)
    alertar_b2_mes1_mes2(df_scores_acum, meses_ordenados)

    # ── 6. Gráficas B1 ───────────────────────────────────────────
    print("\n[*] Generando gráficas para Grupo B1...")
    df_b1_acum = df_scores_acum[df_scores_acum['GRUPO_Limpio'] == 'B1'].copy()

    if df_b1_acum.empty:
        print("  [!] Sin datos B1 en el caché.")
    else:
        # Filtro de trayectoria: sólo estudiantes presentes en los 2 primeros meses
        if len(meses_ordenados) >= 2:
            base_b1 = (df_b1_acum[df_b1_acum['Mes'].isin(meses_ordenados[:2])]
                       [['Documento', 'Materia']].drop_duplicates())
            df_b1_acum = pd.merge(df_b1_acum, base_b1,
                                  on=['Documento', 'Materia'], how='inner')
        print(f"  [OK] Estudiantes válidos B1: "
              f"{len(df_b1_acum['Documento'].unique())}")

        exportar_base_consolidada_b1(
            df_b1_acum, df_matricula, lookup_req027, meses_ordenados)

        # Sankey (trabaja con el caché de promedios por centro, no con scores individuales)
        dibujar_sankey_multimes(df_sankey_acum, "Matemática", meses_ordenados)
        dibujar_sankey_multimes(df_sankey_acum, "Lengua",     meses_ordenados)

        # Boxplot y KDE (puntajes individuales, colores automáticos por posición)
        generar_kde_y_boxplot(df_scores_acum, meses_ordenados, sufijo="B1")

        # Barras apiladas
        generar_barras_por_grupo_grados(
            df_barras_grupos_acum, meses_ordenados, sufijo="B1")
        generar_barras_todos_los_grados(
            df_barras_todos_acum, meses_ordenados, sufijo="B1")

    # ── 7. Gráficas B2 ───────────────────────────────────────────
    print("\n[*] Generando gráficas para Grupo B2...")
    df_b2_mes = df_scores_mes[df_scores_mes['GRUPO_Limpio'] == 'B2'].copy()

    if df_b2_mes.empty:
        print("  [!] Sin datos B2 en el mes actual.")
    else:
        generar_excel_y_barras_b2(df_b2_mes, df_matricula, lookup_req027)
        # B2: solo el mes actual en el boxplot (un solo panel)
        generar_kde_y_boxplot(
            df_scores_acum, [etiqueta_mes],
            sufijo="B2",
            color_list=["#b3b3ff"])
        generar_barras_todos_los_grados(
            df_barras_todos_acum, [etiqueta_mes], sufijo="B2")

    # ── 8. Excel Maestro ──────────────────────────────────────────
    construir_excel_unico(
        df_scores_acum, df_matricula, meses_ordenados, lookup_req027)

    print(f"\n{'='*60}")
    print(f"  [OK] ¡Proceso completado!")
    print(f"       Gráficas y Excel : {PATH_OUTPUT_PLOTS}")
    print(f"       Cachés           : {PATH_CACHE}")
    print(f"{'='*60}\n")