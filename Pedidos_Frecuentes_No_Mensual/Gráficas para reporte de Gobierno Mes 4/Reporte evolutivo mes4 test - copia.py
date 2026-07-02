"""
Reporte_Evolutivo_Mes4_TEST.py
==============================
Script de verificación para Mes 4.

Propósito:
  Confirmar que el sistema de caché funciona correctamente: lee los datos
  de Mes 1, 2 y 3 desde el Excel de caché creado por el script de Mes 3,
  genera datos sintéticos de Mes 4 (para prueba), y produce todos los
  plots con 4 paneles.

Cuando tengas los CSVs reales de Mes 4:
  - Pon la ruta real en DIRECTORIOS_MESES (descomenta la línea real,
    comenta o elimina la sección "# DATOS SINTÉTICOS").
  - Todo lo demás queda igual.
"""

import os
import sys
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import unicodedata
import numpy as np

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
DIRECTORIOS_MESES = [
    # Meses anteriores — se leerán desde el caché, no de CSVs
    (None, "Mes 1 (Marzo)"),
    (None, "Mes 2 (Abril)"),
    (None, "Mes 3 (Mayo)"),
    # Mes 4 — pon aquí la ruta real cuando tengas los CSVs:
    # (r"H:\...\04_PROGRESO_Junio\Interim_CSVs\Resultados", "Mes 4 (Junio)"),
    (None, "Mes 4 (Junio)"),   # None = se usarán datos sintéticos de prueba
]

PATH_METADATA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_REQ027   = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\REQ_027_1.xlsx"

# Carpeta de salida de Mes 4
PATH_OUTPUT_PLOTS = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4"

# Caché creado por el script de Mes 3 — ruta fija, no cambia
PATH_CACHE_EXCEL = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\_Cache_Datos_Por_Grafica\cache_datos_graficas.xlsx"

os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)
os.makedirs(os.path.dirname(PATH_CACHE_EXCEL), exist_ok=True)

# ─── Exclusiones puntuales ────────────────────────────────────────────────────
# Escuelas B2 que participaron por error en meses que no les correspondían.
# Formato: {(codigo_centro, etiqueta_mes): "motivo"}
# Los datos de estas escuelas se eliminan SOLO para el mes indicado;
# en los demás meses se conservan normalmente.
EXCLUSIONES_B2 = {
    ("11432", "Mes 1 (Marzo)"): "CENTRO ESCOLAR COLONIA SAN MAURICIO participó por error en Marzo",
}


# ==========================================
# 2. COLORES Y CONSTANTES  (idénticos al script base)
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

COLORES_MESES = ["#80b1d3", "#fb8072", "#b3de69", "#c9b2e8", "#f4a261",
                 "#8dd3c7", "#fdb462", "#bc80bd", "#ccebc5", "#ffed6f"]
COLORES_NODO  = ['#065f46', '#84cc16', '#facc15', '#ff8c2e', '#991b1b']
COLORES_LINK  = [
    'rgba(6, 95, 70, 0.4)',
    'rgba(132, 204, 22, 0.4)',
    'rgba(250, 204, 21, 0.4)',
    'rgba(255, 140, 46, 0.4)',
    'rgba(153, 27, 27, 0.4)'
]

GRUPOS_GRADOS = {
    "2° y 3° Grado":      [2, 3],
    "4°, 5° y 6° Grado":  [4, 5, 6],
    "7°, 8° y 9° Grado":  [7, 8, 9],
    "10° y 11° Grado":    [10, 11],
}

COLS_CACHE = [
    COL_SCORE_EXACTA, 'Grado_Num', 'Grado_Str', 'Nivel_Logro',
    'Nro de Centro', 'Centro', 'Grupo', 'Documento',
    'Materia', 'Mes', 'GRUPO_Limpio'
]

# ==========================================
# 3. FUNCIONES DE APOYO
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

def preparar_entorno():
    plt.style.use('seaborn-v0_8-whitegrid')
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
                                .str.replace(r'\.0$', '', regex=True).str.strip())
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
        print("  [!] ADVERTENCIA: REQ_027_1.xlsx no encontrado.")
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
                  .str.replace("'", '', regex=False).str.strip())
    df['GROUP_ID'] = df['GROUP_ID'].astype(str).str.replace('"', '', regex=False).str.strip()
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
    nie_key = df_out[nie_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_out['CÓDIGO_SECCIÓN'] = nie_key.map(lookup_req027).fillna('').values
    return df_out

# ==========================================
# 4. CACHÉ
# ==========================================
def leer_cache():
    if not os.path.exists(PATH_CACHE_EXCEL):
        print(f"  [!] Caché no encontrado en: {PATH_CACHE_EXCEL}")
        return {}
    print(f"  [*] Leyendo caché desde: {PATH_CACHE_EXCEL}")
    try:
        xl = pd.ExcelFile(PATH_CACHE_EXCEL)
        cache = {}
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, dtype=str)
            df[COL_SCORE_EXACTA] = pd.to_numeric(df[COL_SCORE_EXACTA], errors='coerce')
            df['Grado_Num']      = pd.to_numeric(df['Grado_Num'],       errors='coerce')
            cache[sheet] = df
            print(f"    -> Hoja '{sheet}': {len(df):,} filas cargadas del caché")
        return cache
    except Exception as e:
        print(f"  [!] Error leyendo caché: {e}")
        return {}

def guardar_cache(cache_dict):
    print(f"\n[*] Guardando caché actualizado en: {PATH_CACHE_EXCEL}")
    try:
        with pd.ExcelWriter(PATH_CACHE_EXCEL, engine='openpyxl') as writer:
            for mes_label, df in cache_dict.items():
                sheet_name = mes_label[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"    -> Hoja '{sheet_name}': {len(df):,} filas guardadas")
        print("  [OK] Caché guardado.")
    except Exception as e:
        print(f"  [!] Error guardando caché: {e}")


def aplicar_exclusiones_b2(df):
    """
    Elimina filas de escuelas B2 que participaron por error en un mes específico.
    Sólo afecta a las combinaciones (codigo_centro, mes) definidas en EXCLUSIONES_B2.
    """
    if not EXCLUSIONES_B2:
        return df
    mask = pd.Series(True, index=df.index)
    for (codigo, mes_label), motivo in EXCLUSIONES_B2.items():
        excluir = (df['Nro de Centro'] == codigo) & (df['Mes'] == mes_label)
        n = excluir.sum()
        if n > 0:
            print(f"  [EXCLUSION] {n:,} registros eliminados — código {codigo} en {mes_label}")
            print(f"              Motivo: {motivo}")
        mask = mask & ~excluir
    return df[mask].copy()

# ==========================================
# 5. DATOS SINTÉTICOS DE MES 4  (solo para prueba)
# ==========================================
def generar_datos_sinteticos_mes4(df_meses_previos, etiqueta_mes4):
    """
    Genera datos de prueba para Mes 4 basados en los centros y estudiantes
    que ya existen en el caché. Los puntajes tienen una leve mejora respecto
    a Mes 3 para que el Sankey muestre movimiento real.
    """
    print(f"\n[*] Generando datos SINTÉTICOS para {etiqueta_mes4}...")

    # Tomar la estructura de centros/estudiantes/grados del mes anterior (Mes 3)
    mes3 = "Mes 3 (Mayo)"
    if mes3 not in df_meses_previos:
        # Si no hay Mes 3 en caché, usar el último mes disponible
        mes3 = list(df_meses_previos.keys())[-1]

    df_base = df_meses_previos[mes3].copy()

    np.random.seed(2026)   # semilla fija para reproducibilidad

    # Simular una mejora media de ~3 puntos con varianza similar a los datos reales
    mejora = np.random.normal(loc=3.0, scale=8.0, size=len(df_base))
    nuevos_puntajes = (df_base[COL_SCORE_EXACTA] + mejora).clip(0, 100)

    df_mes4 = df_base.copy()
    df_mes4[COL_SCORE_EXACTA] = nuevos_puntajes
    df_mes4['Nivel_Logro']    = df_mes4[COL_SCORE_EXACTA].apply(clasificar_puntaje)
    df_mes4['Mes']            = etiqueta_mes4

    print(f"  [OK] {len(df_mes4):,} registros sintéticos generados para {etiqueta_mes4}")
    print(f"       Media Mes 3: {df_base[COL_SCORE_EXACTA].mean():.1f}  →  "
          f"Media Mes 4 (sintético): {df_mes4[COL_SCORE_EXACTA].mean():.1f}")

    return df_mes4[COLS_CACHE]

# ==========================================
# 6. CONTEOS Y ALERTAS
# ==========================================
def generar_reportes_conteos_separados(df_base_total):
    print("\n[*] Generando Reportes de Totales de Estudiantes...")
    grados_esperados = list(range(2, 12))
    meses_labels     = list(dict.fromkeys(df_base_total['Mes'].tolist()))

    def crear_hoja_materia(df, materia):
        df_sub = df[df['Materia'] == materia]
        if df_sub.empty:
            return pd.DataFrame()
        pivot = df_sub.pivot_table(
            index='Mes', columns='Grado_Num', values='Documento',
            aggfunc='nunique', fill_value=0)
        for col in grados_esperados:
            if col not in pivot.columns:
                pivot[col] = 0
        pivot = pivot[grados_esperados]
        pivot.columns = [f"{c}°" for c in pivot.columns]
        pivot['Total por prueba'] = pivot.sum(axis=1)
        return pivot

    def exportar_archivo(df_target, filename):
        dir_cnt = subdir("Conteos", "Excel")
        path_output = os.path.join(dir_cnt, filename)
        with pd.ExcelWriter(path_output, engine='openpyxl') as writer:
            for materia in ['Matemática', 'Lengua']:
                df_hoja = crear_hoja_materia(df_target, materia)
                if not df_hoja.empty:
                    df_hoja.to_excel(writer, sheet_name=materia)
                else:
                    pd.DataFrame({'Aviso': [f'Sin datos para {materia}']}).to_excel(
                        writer, sheet_name=materia, index=False)
        print(f"  [OK] Guardado: {filename}")

    df_b1 = df_base_total[df_base_total['GRUPO_Limpio'] == 'B1']
    df_b2 = df_base_total[df_base_total['GRUPO_Limpio'] == 'B2']
    exportar_archivo(df_base_total, "Conteo_Estudiantes_B1_y_B2.xlsx")
    exportar_archivo(df_b1,         "Conteo_Estudiantes_Solo_B1.xlsx")
    exportar_archivo(df_b2,         "Conteo_Estudiantes_Solo_B2.xlsx")

def alertar_b2_mes1_mes2(df_base_total, meses_labels):
    print("\n[*] Verificando escuelas B2 con estudiantes en meses previos...")
    if len(meses_labels) >= 2:
        meses_tempranos = meses_labels[:2]
        df_b2_temprano = df_base_total[
            (df_base_total['GRUPO_Limpio'] == 'B2') &
            (df_base_total['Mes'].isin(meses_tempranos))]
        if not df_b2_temprano.empty:
            print("  [!] ATENCIÓN: Escuelas B2 con participación temprana:")
            conteo = df_b2_temprano.groupby(
                ['Nro de Centro', 'Centro', 'Materia', 'Mes']
            )['Documento'].nunique().reset_index()
            for _, row in conteo.iterrows():
                print(f"      - {row['Nro de Centro']} | {row['Materia']} "
                      f"| {row['Mes']}: {row['Documento']} est.")
        else:
            print("  [OK] Ninguna escuela B2 fue evaluada en meses previos.")
    else:
        print("  [-] Menos de 2 meses — sin validación.")

# ==========================================
# 7. SANKEY
# ==========================================
def calc_y_centers(df_mes):
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

def dibujar_sankey_multimes(df_master, materia, meses_labels, sufijo_sankey="B1"):
    print(f"  -> Dibujando Sankey para {materia} ({sufijo_sankey})...")
    df_mat = df_master[df_master['Materia'] == materia].copy()
    if df_mat.empty: return

    meses_validos, nombres_validos = [], []
    for mes in meses_labels:
        df_m = df_mat[df_mat['Mes'] == mes]
        if not df_m.empty:
            promedios = (df_m.groupby(['Nro de Centro', 'Centro'])
                         [COL_SCORE_EXACTA].mean().reset_index())
            promedios['Nivel'] = promedios[COL_SCORE_EXACTA].apply(clasificar_puntaje)
            promedios = promedios.rename(columns={'Nro de Centro': 'Código'})
            meses_validos.append(promedios)
            nombres_validos.append(mes)
    if len(meses_validos) < 2: return

    labels_nodos, colores_nodos_lista, node_x, node_y = [], [], [], []
    sources, targets, values, link_colors = [], [], [], []

    df_stats = pd.merge(
        meses_validos[0][['Código', 'Nivel']],
        meses_validos[-1][['Código', 'Nivel']],
        on='Código', suffixes=('_Ini', '_Fin'), how='inner')
    n_centros = len(df_stats)
    mantienen = suben = bajan = 0
    if n_centros > 0:
        for _, row in df_stats.iterrows():
            if row['Nivel_Ini'] not in NIVELES or row['Nivel_Fin'] not in NIVELES: continue
            d = NIVELES.index(row['Nivel_Fin']) - NIVELES.index(row['Nivel_Ini'])
            if d == 0:  mantienen += 1
            elif d < 0: suben     += 1
            else:       bajan     += 1
        subtitulo_html = (
            f"<span style='font-size:13px;color:#888'>"
            f"Población base (Grupo B1): {n_centros} centros<br>"
            f"Mantienen: {mantienen} ({mantienen/n_centros*100:.1f}%) · "
            f"Suben: {suben} ({suben/n_centros*100:.1f}%) · "
            f"Bajan: {bajan} ({bajan/n_centros*100:.1f}%)</span>"
        )
    else:
        subtitulo_html = ""

    for i, (nombre, df_mes) in enumerate(zip(nombres_validos, meses_validos)):
        x_val = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        centros_y = calc_y_centers(df_mes)
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])

    for i in range(len(meses_validos) - 1):
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(oo + j); targets.append(od + j)
            values.append(1e-9);    link_colors.append('rgba(0,0,0,0)')

    for i in range(len(meses_validos) - 1):
        df_cruce = pd.merge(
            meses_validos[i][['Código', 'Centro', 'Nivel']],
            meses_validos[i + 1][['Código', 'Nivel']],
            on='Código', suffixes=('_A', '_S'), how='inner')
        if df_cruce.empty: continue
        if materia == "Matemática" and i == 0:
            alerta = df_cruce[(df_cruce['Nivel_A'] == 'Bueno') &
                              (df_cruce['Nivel_S'] == 'Bajo')]
            if not alerta.empty:
                print(f"\n      [ALERTA] Escuelas Bueno→Bajo ({nombres_validos[i]}→{nombres_validos[i+1]}):")
                for _, r in alerta.iterrows():
                    print(f"         - {r['Código']} | {r['Centro']}")
        flujos = df_cruce.groupby(['Nivel_A', 'Nivel_S']).size().reset_index(name='Cantidad')
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        mo = {n: j + oo for j, n in enumerate(NIVELES)}
        md = {n: j + od for j, n in enumerate(NIVELES)}
        for _, row in flujos.iterrows():
            if row['Cantidad'] == 0: continue
            si = mo[row['Nivel_A']]; ti = md[row['Nivel_S']]
            sources.append(si); targets.append(ti)
            values.append(row['Cantidad'])
            link_colors.append(COLORES_LINK[si % len(NIVELES)])

    if not sources: return

    fig = go.Figure(data=[go.Sankey(
        arrangement='fixed',
        node=dict(pad=20, thickness=24, line=dict(color="white", width=0.5),
                  label=labels_nodos, color=colores_nodos_lista,
                  x=node_x, y=node_y),
        link=dict(source=sources, target=targets, value=values, color=link_colors)
    )])
    fig.update_layout(
        title=dict(
            text=(f"<span style='font-size:22px;color:#1a2b4c;"
                  f"font-family:Arial,sans-serif;font-weight:bold;'>"
                  f"Trayectorias de niveles — {materia}</span><br>{subtitulo_html}"),
            x=0.01, y=0.95),
        font_size=12,
        width=max(900, 800 + len(meses_validos) * 150),
        height=600,
        margin=dict(l=20, r=20, t=100, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    dir_sank = subdir(sufijo_sankey, "Sankey")
    fig.write_html(os.path.join(dir_sank,
                                f"Sankey_Evolutivo_{sufijo_sankey}_{materia}.html"))
    try:
        fig.write_image(
            os.path.join(dir_sank,
                         f"Sankey_Evolutivo_{sufijo_sankey}_{materia}.png"), scale=2)
    except Exception: pass
    print(f"     [OK] {sufijo_sankey}/Sankey/Sankey_Evolutivo_{sufijo_sankey}_{materia}")
    print(f"     [OK] Sankey {materia} guardado.")


def subdir(sufijo, tipo):
    """
    Returns (and creates) the output subfolder for a given group + plot type.

    Folder layout under PATH_OUTPUT_PLOTS:
      B1/
        Boxplot/
        Sankey/
        Barras_Por_Grado/          ← one PNG per grade (all pruebas in one file)
        Barras_TodosGrados_Combo/  ← all grades × all pruebas in one file
        Barras_TodosGrados_Por_Prueba/ ← one PNG per prueba (all grades)
        Excel/
      B2/
        (same structure)
    """
    folder = os.path.join(PATH_OUTPUT_PLOTS, sufijo, tipo)
    os.makedirs(folder, exist_ok=True)
    return folder

# ==========================================
# 8. BARRAS APILADAS
# ==========================================
def generar_barras_por_grado(df_master, meses_labels, sufijo_archivo="B1"):
    """
    Genera un PNG por cada grado y materia.
    Cada PNG tiene una fila por Prueba (mes), con barras horizontales apiladas
    mostrando el % de estudiantes en cada nivel de logro.
    Esto facilita comparar la evolución de un grado a través de las pruebas.
    """
    print(f"\n[*] Generando Barras por Grado ({sufijo_archivo})...")
    grados_presentes = sorted(df_master['Grado_Num'].dropna().unique())

    for grado_num in grados_presentes:
        grado_str  = f"{int(grado_num)}° Grado"
        df_grado   = df_master[df_master['Grado_Num'] == grado_num].copy()
        if df_grado.empty:
            continue

        for materia in df_grado['Materia'].unique():
            df_mat = df_grado[df_grado['Materia'] == materia]
            meses_con_datos = [m for m in meses_labels
                               if not df_mat[df_mat['Mes'] == m].empty]
            if not meses_con_datos:
                continue

            n_pruebas = len(meses_con_datos)
            fig, axes = plt.subplots(
                n_pruebas, 1,
                figsize=(11, max(3, n_pruebas * 2.8)),
                sharex=True)
            if n_pruebas == 1:
                axes = [axes]

            for i, mes_key in enumerate(meses_con_datos):
                ax       = axes[i]
                df_mes   = df_mat[df_mat['Mes'] == mes_key]
                if df_mes.empty:
                    ax.set_visible(False)
                    continue

                # Porcentaje de cada nivel para este grado en esta prueba
                conteos = df_mes['Nivel_Logro'].value_counts()
                total   = conteos.sum() or 1
                pcts    = {cat: (conteos.get(cat, 0) / total * 100)
                           for cat in ORDEN_CATEGORIAS}

                # Dibujar barras apiladas horizontales (una sola fila por prueba)
                left = 0.0
                for cat in ORDEN_CATEGORIAS:
                    val = pcts[cat]
                    ax.barh(0, val, left=left, height=0.55,
                            color=COLOR_CATEGORIAS[cat],
                            edgecolor='white', linewidth=0.8)
                    if val > 4:
                        ax.text(left + val / 2, 0, f"{val:.1f}%",
                                ha='center', va='center',
                                color='white', fontsize=9, fontweight='bold')
                    left += val

                prueba_num = meses_labels.index(mes_key) + 1
                mes_nombre = mes_key.split('(')[-1].replace(')', '').strip()
                n_est      = len(df_mes['Documento'].unique())
                ax.set_ylabel(f"Prueba {prueba_num}\n({mes_nombre})\nn={n_est:,}",
                              fontsize=10, fontweight='bold',
                              color='#005288', labelpad=8)
                ax.set_xlim(0, 100)
                ax.set_yticks([])
                ax.grid(axis='x', linestyle='--', alpha=0.3)
                ax.spines[['top', 'right', 'left']].set_visible(False)

            # Leyenda en el último eje
            from matplotlib.patches import Patch
            legend_handles = [Patch(facecolor=COLOR_CATEGORIAS[c], label=c)
                              for c in ORDEN_CATEGORIAS]
            axes[-1].legend(handles=legend_handles, title="Nivel de Logro",
                            bbox_to_anchor=(1.01, 1), loc='upper left',
                            fontsize=9)
            axes[-1].set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)

            grado_seguro = grado_str.replace('°', '').replace(' ', '_')
            mat_segura   = materia.replace('á','a').replace('é','e')
            fname = f"Barras_{sufijo_archivo}_{mat_segura}_{grado_seguro}.png"
            fig.suptitle(
                f"Niveles de Logro por Prueba — {sufijo_archivo}\n"
                f"{grado_str} — {materia}",
                fontsize=14, fontweight='bold', color='#1a2b4c', y=1.01)
            plt.tight_layout()
            dir_pg = subdir(sufijo_archivo, "Barras_Por_Grado")
            plt.savefig(os.path.join(dir_pg, fname), dpi=300, bbox_inches='tight')
            plt.close()
            print(f"    [OK] {sufijo_archivo}/Barras_Por_Grado/{fname}")

def generar_barras_todos_los_grados(df_master, meses_labels, sufijo_archivo="B1"):
    """
    Genera dos versiones de barras horizontales apiladas (todos los grados):
      1. Combo   — todos los meses en un único PNG (un panel por prueba).
      2. Por Prueba — un PNG individual por cada prueba.
    """
    print(f"\n[*] Generando Barras (Todos los Grados) ({sufijo_archivo})...")
    dir_combo  = subdir(sufijo_archivo, "Barras_TodosGrados_Combo")
    dir_single = subdir(sufijo_archivo, "Barras_TodosGrados_Por_Prueba")

    for materia in df_master['Materia'].unique():
        df_mat = df_master[df_master['Materia'] == materia].copy()
        if df_mat.empty: continue
        meses_con_datos = [m for m in meses_labels
                           if not df_mat[df_mat['Mes'] == m].empty]
        if not meses_con_datos: continue

        mat_segura       = materia.replace('á','a').replace('é','e')
        grados_presentes = sorted(df_mat['Grado_Num'].dropna().unique(), reverse=True)
        grados_orden     = [f"{int(g)}° Grado" for g in grados_presentes]
        grados_orden.append("Global (Todos)")

        def _build_ct(df_mes_bar):
            df_mes_bar = df_mes_bar.copy()
            df_mes_bar['Grado_Str'] = df_mes_bar['Grado_Num'].astype(str) + "° Grado"
            ct_grados = pd.crosstab(df_mes_bar['Grado_Str'], df_mes_bar['Nivel_Logro'],
                                    normalize='index') * 100
            ct_global = (df_mes_bar['Nivel_Logro']
                         .value_counts(normalize=True).to_frame().T * 100)
            ct_global.index = ["Global (Todos)"]
            ct = pd.concat([ct_grados, ct_global])
            for cat in ORDEN_CATEGORIAS:
                if cat not in ct.columns: ct[cat] = 0
            orden_actual = [g for g in grados_orden if g in ct.index]
            return ct[ORDEN_CATEGORIAS].reindex(orden_actual)

        def _decorate_ax(ax, mes_key, is_last, show_ylabel, show_title=True):
            prueba_num = meses_labels.index(mes_key) + 1
            mes_nombre = mes_key.split("(")[-1].replace(")", "").strip()
            if show_title:
                ax.set_title(f"Prueba {prueba_num}\n({mes_nombre})",
                             fontsize=16, color="#005288", pad=14, fontweight="bold")
            ax.set_xlim(0, 100)
            ax.set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)
            ax.set_ylabel("Grado" if show_ylabel else "")
            if is_last:
                h, l = ax.get_legend_handles_labels()
                ax.legend(h, l, title="Nivel de Logro",
                          bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=9)
            for p in ax.patches:
                w = p.get_width()
                if w > 4:
                    ax.text(p.get_x() + w / 2, p.get_y() + p.get_height() / 2,
                            f"{w:.1f}%", ha="center", va="center",
                            color="white", fontsize=9, fontweight="bold")

        # ── 1. Combo: todos los meses en un archivo ───────────────────────────
        n = len(meses_con_datos)
        fig, axes = plt.subplots(1, n, figsize=(max(9, n * 9), 9), sharey=True)
        if n == 1: axes = [axes]
        for i, mes_key in enumerate(meses_con_datos):
            ax = axes[i]
            df_mes_bar = df_mat[df_mat['Mes'] == mes_key]
            if df_mes_bar.empty:
                ax.set_visible(False); continue
            ct = _build_ct(df_mes_bar)
            ct.plot(kind='barh', stacked=True, ax=ax,
                    color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS],
                    legend=False, edgecolor='white', width=0.7)
            _decorate_ax(ax, mes_key, is_last=(i == n - 1), show_ylabel=(i == 0))
        plt.suptitle(f"Niveles de Logro — {sufijo_archivo} — {materia}\n"
                     "(Todos los Grados — Todas las Pruebas)",
                     fontsize=16, fontweight='bold', color='#1a2b4c')
        plt.tight_layout(rect=[0, 0, 0.88, 1])
        fname_combo = f"Barras_{sufijo_archivo}_{mat_segura}_TodosGrados_Combo.png"
        plt.savefig(os.path.join(dir_combo, fname_combo), dpi=300, bbox_inches='tight')
        plt.close()
        print(f"    [OK] {sufijo_archivo}/Barras_TodosGrados_Combo/{fname_combo}")

        # ── 2. Individual: un archivo por prueba ──────────────────────────────
        for mes_key in meses_con_datos:
            df_mes_bar = df_mat[df_mat['Mes'] == mes_key]
            if df_mes_bar.empty: continue
            ct = _build_ct(df_mes_bar)
            fig_s, ax_s = plt.subplots(figsize=(9, max(5, len(ct) * 0.55 + 2)))
            ct.plot(kind='barh', stacked=True, ax=ax_s,
                    color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS],
                    legend=False, edgecolor='white', width=0.7)
            _decorate_ax(ax_s, mes_key, is_last=True, show_ylabel=True, show_title=False)
            prueba_num = meses_labels.index(mes_key) + 1
            mes_nombre = mes_key.split("(")[-1].replace(")", "").strip()
            fig_s.suptitle(
                f"Niveles de Logro — {sufijo_archivo} — {materia}\n"
                f"Todos los Grados — Prueba {prueba_num} ({mes_nombre})",
                fontsize=14, fontweight='bold', color='#1a2b4c')
            plt.tight_layout(rect=[0, 0, 0.88, 1])
            fname_s = (f"Barras_{sufijo_archivo}_{mat_segura}_TodosGrados"
                       f"_Prueba{prueba_num}_{mes_nombre}.png")
            plt.savefig(os.path.join(dir_single, fname_s), dpi=300, bbox_inches='tight')
            plt.close()
            print(f"    [OK] {sufijo_archivo}/Barras_TodosGrados_Por_Prueba/{fname_s}")

def generar_boxplot(df_master, meses_labels, sufijo_archivo="B1", color_list=None):
    if color_list is None:
        color_list = COLORES_MESES
    print(f"\n[*] Generando Boxplot+Jitter ({sufijo_archivo})...")
    for materia in df_master['Materia'].unique():
        df_sub = df_master[df_master['Materia'] == materia].copy()
        print(f"    - {materia}...")
        meses_con_datos = [m for m in meses_labels
                           if not df_sub[df_sub['Mes'] == m].empty]
        n_meses   = len(meses_con_datos)
        fig, axes = plt.subplots(n_meses, 1,
                                 figsize=(12, 4 * n_meses), sharex=True)
        if n_meses == 1: axes = [axes]
        for i, mes_key in enumerate(meses_con_datos):
            ax        = axes[i]
            color     = color_list[i % len(color_list)]
            datos_box = df_sub[df_sub['Mes'] == mes_key][COL_SCORE_EXACTA].dropna()
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
            prueba_num = meses_labels.index(mes_key) + 1
            mes_nombre = mes_key.split('(')[-1].replace(')', '').strip()
            ax.set_ylabel(f"Prueba {prueba_num}\n({mes_nombre})",
                          fontsize=11, fontweight='bold', color='#005288', labelpad=10)
            ax.set_xlim(0, 100); ax.set_yticks([])
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
        fig.suptitle(f"Distribución por Prueba — {sufijo_archivo} — {materia}\n"
                     "(todos los grados)",
                     fontsize=15, fontweight='bold', color='#1a2b4c', y=1.01)
        plt.tight_layout()
        dir_box = subdir(sufijo_archivo, "Boxplot")
        plt.savefig(
            os.path.join(dir_box,
                         f"Boxplot_Jitter_{sufijo_archivo}_{materia}.png"),
            dpi=300, bbox_inches='tight')
        plt.close()
        print(f"    [OK] {sufijo_archivo}/Boxplot/Boxplot_Jitter_{sufijo_archivo}_{materia}.png")
        print(f"    [OK] Boxplot_{sufijo_archivo}_{materia}.png")

# ==========================================
# 10. EXCELS DE RESULTADOS
# ==========================================
def exportar_base_consolidada_b1(df_master, df_matricula, lookup_req027, meses_labels):
    print("\n[*] Generando Base de Datos Consolidada (Grupo B1)...")
    try:
        df_pivot = df_master.pivot_table(
            index=['Nro de Centro', 'Centro', 'Grupo', 'Documento', 'Grado_Str'],
            columns=['Materia', 'Mes'], values=COL_SCORE_EXACTA, aggfunc='first')
        score_cols_raw = sorted(
            df_pivot.columns.tolist(),
            key=lambda col: (0 if 'Matemática' in col[0] else 1,
                             next((i for i, m in enumerate(meses_labels)
                                   if m in col[1]), 99)))
        flat_map = {col: (f"{COL_SCORE_EXACTA} "
                          f"(para {col[0].lower()} {col[1].lower()})")
                    for col in df_pivot.columns}
        df_pivot.columns = [flat_map[c] for c in df_pivot.columns]
        df_pivot = df_pivot.reset_index().rename(columns={'Grado_Str': 'Grado'})
        interleaved = []
        for rc in score_cols_raw:
            sc = flat_map[rc]; nc = f"Nivel {rc[0]} ({rc[1]})"
            if sc in df_pivot.columns:
                df_pivot[nc] = df_pivot[sc].apply(clasificar_puntaje)
                interleaved += [sc, nc]
        id_cols   = [c for c in ['Nro de Centro', 'Centro', 'Grupo',
                                  'Documento', 'Grado'] if c in df_pivot.columns]
        placed    = set(id_cols + interleaved)
        remaining = [c for c in df_pivot.columns if c not in placed]
        df_pivot  = df_pivot[id_cols + interleaved + remaining]
        cols_sol  = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
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
        dir_xls = subdir("B1", "Excel")
        df_out[header + rest].to_excel(
            os.path.join(dir_xls, "Base_Estudiantes_B1_Consolidada.xlsx"),
            index=False)
        print("  [OK] B1/Excel/Base_Estudiantes_B1_Consolidada.xlsx guardado.")
    except Exception as e:
        print(f"  [!] Error al exportar base B1: {e}")

def generar_excel_y_barras_b2(df_b2_mes, df_matricula, lookup_req027):
    if df_b2_mes.empty:
        print("\n  [!] Sin estudiantes B2."); return
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
    df_excel   = pd.merge(df_matricula[cols_exist + ['NIE_Limpio']],
                          df_pivot, left_on='NIE_Limpio', right_on='Documento',
                          how='inner')
    df_excel   = df_excel.drop(columns=['NIE_Limpio', 'Documento'], errors='ignore')
    df_excel   = _aplicar_group_id(df_excel, lookup_req027)
    priority   = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                  'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE', 'PRIMER_APELLIDO',
                  'SEGUNDO_APELLIDO', 'Puntaje Matemática', 'Nivel Matemática',
                  'Puntaje Lengua', 'Nivel Lengua']
    final_cols = [c for c in priority if c in df_excel.columns]
    dir_xls_b2 = subdir("B2", "Excel")
    df_excel[final_cols].to_excel(
        os.path.join(dir_xls_b2, "Resultados_B2_MesActual.xlsx"), index=False)
    print("  [OK] B2/Excel/Resultados_B2_MesActual.xlsx guardado.")
    # barras por grado generadas desde el bloque principal con todos los meses B2

def construir_excel_unico(df_merged, df_matricula, meses_labels, lookup_req027):
    print("\n[*] Construyendo Excel Maestro (hoja única)...")
    df_b1 = df_merged[df_merged['GRUPO_Limpio'] == 'B1'].copy()
    if len(meses_labels) >= 3:
        base  = (df_b1[df_b1['Mes'].isin(meses_labels[:2])]
                 [['Documento', 'Materia']].drop_duplicates())
        df_b1 = pd.merge(df_b1, base, on=['Documento', 'Materia'], how='inner')
    df_b1['Grupo_Paars'] = 'B1'
    df_b2 = df_merged[df_merged['GRUPO_Limpio'] == 'B2'].copy()
    if len(meses_labels) >= 3:
        df_b2 = df_b2[df_b2['Mes'] == meses_labels[-1]]
    df_b2['Grupo_Paars'] = 'B2'
    df_all = pd.concat([df_b1, df_b2], ignore_index=True)
    if df_all.empty:
        print("  [!] Sin datos para el Excel Maestro."); return None
    df_pivot = df_all.pivot_table(
        index=['Documento', 'Nro de Centro', 'Centro',
               'Grado_Str', 'GRUPO_Limpio', 'Grupo_Paars'],
        columns=['Materia', 'Mes'], values=COL_SCORE_EXACTA, aggfunc='first')
    score_cols_raw = sorted(
        df_pivot.columns.tolist(),
        key=lambda col: (0 if 'Matemática' in col[0] else 1,
                         next((i for i, m in enumerate(meses_labels)
                               if m in col[1]), 99)))
    flat_map = {col: f"Puntaje {col[0]} ({col[1]})" for col in df_pivot.columns}
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
    dir_xls_m = subdir("Maestro", "Excel")
    out_path = os.path.join(dir_xls_m, "Excel_Maestro_HojaUnica.xlsx")
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name="Todos los Datos", index=False)
        ws = writer.sheets["Todos los Datos"]
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes    = "A2"
    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"  [OK] Excel Maestro → {out_path} ({len(df_out):,} filas, {size_mb:.1f} MB)")
    return out_path


# ==========================================
# 11. EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print(f"\n{'='*60}")
    print("  REPORTE EVOLUTIVO MES 4 — VERIFICACIÓN CON CACHÉ")
    print(f"{'='*60}")

    preparar_entorno()
    meses_labels = [etiqueta for _, etiqueta in DIRECTORIOS_MESES]

    # ── 1. Matrícula ──────────────────────────────────────────────
    df_matricula = cargar_matricula()
    if df_matricula.empty:
        print("\n[!] Sin matrícula — proceso detenido."); sys.exit()
    if 'CODIGO' not in df_matricula.columns:
        print("  [!] ERROR: columna 'CODIGO' no existe."); sys.exit()

    mapa_grupos_escuela = (
        df_matricula[['CODIGO', 'GRUPO_Limpio']]
        .dropna().drop_duplicates(subset=['CODIGO'])
        .set_index('CODIGO')['GRUPO_Limpio'])
    lookup_req027 = cargar_req027()

    # ── 2. Leer caché de Mes 3 ────────────────────────────────────
    print("\n[*] Leyendo caché de meses anteriores...")
    cache = leer_cache()

    meses_en_cache = set(cache.keys())
    print(f"  [*] Meses en caché: {list(meses_en_cache)}")

    # ── 3. Obtener datos de Mes 4 ─────────────────────────────────
    etiqueta_mes4 = meses_labels[-1]   # "Mes 4 (Junio)"

    if etiqueta_mes4 not in cache:
        ruta_mes4 = DIRECTORIOS_MESES[-1][0]

        if ruta_mes4 is not None and os.path.exists(ruta_mes4):
            # ── Ruta real disponible: leer CSVs ───────────────────
            print(f"\n[*] Leyendo CSVs reales de {etiqueta_mes4}...")
            from extraer_csvs_mes_helper import extraer_csvs_mes   # si lo separas
            # (en este script está inline abajo)
            lista_dfs = []
            for f in glob.glob(os.path.join(ruta_mes4, "*.csv")):
                if 'legend' in os.path.basename(f).lower(): continue
                try:
                    df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
                    df.columns = df.columns.str.strip()
                    materia = ('Matemática' if 'MAT' in os.path.basename(f).upper()
                               else 'Lengua')
                    if COL_SCORE_EXACTA not in df.columns: continue
                    col_anular = next((c for c in df.columns if 'anular'   in c.lower()), None)
                    col_grado  = next((c for c in df.columns if 'grado'    in c.lower()), None)
                    col_codigo = next(
                        (c for c in df.columns
                         if 'nro de centro' in c.lower() or 'código' in c.lower()
                         or 'codigo' in c.lower()
                         or ('centro' in c.lower() and 'id' in c.lower())), None)
                    col_centro = next(
                        (c for c in df.columns
                         if c.lower() in ('centro','nombre centro','institución')), None)
                    if not col_centro:
                        col_centro = next(
                            (c for c in df.columns
                             if 'centro' in c.lower() and 'nro' not in c.lower()
                             and 'código' not in c.lower()), None)
                    col_grupo = next(
                        (c for c in df.columns
                         if 'grupo' in c.lower() or 'sección' in c.lower()
                         or 'seccion' in c.lower()), None)
                    col_doc = next(
                        (c for c in df.columns
                         if 'documento' in c.lower() or 'nie' in c.lower()
                         or c.lower() == 'id'), None)
                    if not (col_grado and col_codigo and col_doc): continue
                    if col_anular:
                        def excluir_por_tiempo(val):
                            if pd.isna(val): return False
                            v = ''.join(c for c in unicodedata.normalize('NFD', str(val).lower())
                                        if unicodedata.category(c) != 'Mn')
                            return '5 min' in v or 'demora' in v or 'duracion' in v
                        df = df[~df[col_anular].apply(excluir_por_tiempo)].copy()
                    df[COL_SCORE_EXACTA] = pd.to_numeric(
                        df[COL_SCORE_EXACTA].astype(str).str.replace(',','.'), errors='coerce')
                    df = df.dropna(subset=[COL_SCORE_EXACTA])
                    if df.empty: continue
                    df['Grado_Num']     = df[col_grado].apply(extraer_numero_grado)
                    df['Grado_Str']     = df[col_grado].astype(str).str.strip()
                    df['Nro de Centro'] = (df[col_codigo].astype(str)
                                           .str.replace(r'\.0$','',regex=True).str.strip())
                    df['Centro']        = (df[col_centro].astype(str).str.strip()
                                           if col_centro else "Desconocido")
                    df['Grupo']         = (df[col_grupo].astype(str).str.strip()
                                           if col_grupo  else "Desconocido")
                    df['Documento']     = (df[col_doc].astype(str)
                                           .str.replace(r'\.0$','',regex=True).str.strip())
                    df['Nivel_Logro']   = df[COL_SCORE_EXACTA].apply(clasificar_puntaje)
                    df_final = df[[COL_SCORE_EXACTA, 'Grado_Num', 'Grado_Str', 'Nivel_Logro',
                                   'Nro de Centro', 'Centro', 'Grupo', 'Documento']].copy()
                    df_final['Materia'] = materia
                    df_final['Mes']     = etiqueta_mes4
                    df_final = df_final[df_final['Nro de Centro'] != '99999']
                    df_final = df_final[
                        ~df_final['Centro'].str.contains('Centro Virtual', case=False, na=False)]
                    lista_dfs.append(df_final)
                except Exception: pass

            if lista_dfs:
                df_mes4 = pd.concat(lista_dfs, ignore_index=True)
                df_mes4 = pd.merge(df_mes4, df_matricula[['NIE_Limpio']],
                                   left_on='Documento', right_on='NIE_Limpio', how='inner')
                df_mes4['GRUPO_Limpio'] = (
                    df_mes4['Nro de Centro'].map(mapa_grupos_escuela).fillna('DESCONOCIDO'))
                cache[etiqueta_mes4] = df_mes4[COLS_CACHE]
            else:
                print(f"  [!] No se leyeron CSVs válidos — usando datos sintéticos.")
                cache[etiqueta_mes4] = generar_datos_sinteticos_mes4(cache, etiqueta_mes4)
        else:
            # ── Sin ruta real: usar datos sintéticos ──────────────
            if not cache:
                print("\n[!] Caché vacío y sin ruta de CSVs — proceso detenido.")
                sys.exit()
            cache[etiqueta_mes4] = generar_datos_sinteticos_mes4(cache, etiqueta_mes4)
    else:
        print(f"  [*] {etiqueta_mes4} ya está en caché — se omite la lectura.")

    # ── 4. Guardar caché actualizado (ahora incluye Mes 4) ────────
    guardar_cache({m: cache[m] for m in meses_labels if m in cache})

    # ── 5. Construir df_merged_full ───────────────────────────────
    meses_disponibles = [m for m in meses_labels if m in cache]
    print(f"\n[*] Meses disponibles para los plots: {meses_disponibles}")
    df_merged_full = pd.concat(
        [cache[m] for m in meses_disponibles], ignore_index=True)

    # Aplicar exclusiones puntuales de escuelas B2
    df_merged_full = aplicar_exclusiones_b2(df_merged_full)

    # ── 6. Conteos y alertas ──────────────────────────────────────
    generar_reportes_conteos_separados(df_merged_full)
    alertar_b2_mes1_mes2(df_merged_full, meses_labels)

    # ── 7. Filtros y gráficas B1 ──────────────────────────────────
    print("\n[*] Aplicando filtros de trayectoria para gráficas B1...")
    df_b1 = df_merged_full[df_merged_full['GRUPO_Limpio'] == 'B1'].copy()
    if df_b1.empty:
        print("  [!] Sin datos B1.")
    else:
        if len(meses_labels) >= 3:
            base  = (df_b1[df_b1['Mes'].isin(meses_labels[:2])]
                     [['Documento', 'Materia']].drop_duplicates())
            df_b1 = pd.merge(df_b1, base, on=['Documento', 'Materia'], how='inner')
        print(f"  [OK] Estudiantes válidos B1: {len(df_b1['Documento'].unique())}")

        exportar_base_consolidada_b1(df_b1, df_matricula, lookup_req027, meses_labels)
        dibujar_sankey_multimes(df_b1, "Matemática", meses_labels, sufijo_sankey="B1")
        dibujar_sankey_multimes(df_b1, "Lengua",     meses_labels, sufijo_sankey="B1")
        generar_boxplot(df_b1, meses_labels, sufijo_archivo="B1",
                        color_list=COLORES_MESES)
        generar_barras_por_grado(df_b1, meses_labels, sufijo_archivo="B1")
        generar_barras_todos_los_grados(df_b1, meses_labels, sufijo_archivo="B1")

    # ── 8. Gráficas B2 ────────────────────────────────────────────
    df_b2 = df_merged_full[df_merged_full['GRUPO_Limpio'] == 'B2'].copy()
    if not df_b2.empty:
        # Meses en los que B2 tiene datos (orden cronológico preservado)
        meses_b2 = [m for m in meses_labels if not df_b2[df_b2['Mes'] == m].empty]
        df_b2_ultimo_mes = df_b2[df_b2['Mes'] == meses_b2[-1]].copy()

        generar_excel_y_barras_b2(df_b2_ultimo_mes, df_matricula, lookup_req027)

        # Sankey B2: desde el primer mes en que B2 tiene datos
        if len(meses_b2) >= 2:
            dibujar_sankey_multimes(df_b2, "Matemática", meses_b2,
                                    sufijo_sankey="B2")
            dibujar_sankey_multimes(df_b2, "Lengua",     meses_b2,
                                    sufijo_sankey="B2")
        else:
            print("  [!] B2: solo un mes disponible — Sankey requiere al menos 2.")

        # Boxplot B2: todos los meses donde B2 tiene datos
        generar_boxplot(df_b2, meses_labels, sufijo_archivo="B2",
                        color_list=["#b3b3ff", "#9b9bff", "#7b7bdf", "#5b5bbf"])

        # Se pasa meses_labels (lista global) para que la numeración de Prueba
        # sea consistente con B1 (ej. Mayo = Prueba 3, no Prueba 1)
        generar_barras_por_grado(df_b2, meses_labels, sufijo_archivo="B2")
        generar_barras_todos_los_grados(df_b2, meses_labels, sufijo_archivo="B2")
    else:
        print("\n  [!] Sin datos B2 en ningún mes.")

    # ── 9. Excel Maestro ──────────────────────────────────────────
    construir_excel_unico(df_merged_full, df_matricula, meses_labels, lookup_req027)

    print(f"\n{'='*60}")
    print(f"  [OK] ¡Proceso completado!")
    print(f"       Gráficas y Excel : {PATH_OUTPUT_PLOTS}")
    print(f"       Caché actualizado: {PATH_CACHE_EXCEL}")
    print(f"{'='*60}\n")