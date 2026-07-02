"""
Sankey_Estudiantes_Evolutivo.py
================================
Genera Sankeys de trayectorias de nivel por ESTUDIANTE (no por escuela)
para Matemática y Lengua, separados por Grupo B1 y B2.

Filtro de inclusión:
  - B1 → solo estudiantes presentes en los 4 meses (Mes 1, 2, 3 y 4)
  - B2 → solo estudiantes presentes en los 2 últimos meses (Mes 3 y 4)

Salidas:
  - 4 archivos HTML + PNG de Sankey:
      B1/Sankey/Sankey_Estudiantes_B1_Matemática.html/.png
      B1/Sankey/Sankey_Estudiantes_B1_Lengua.html/.png
      B2/Sankey/Sankey_Estudiantes_B2_Matemática.html/.png
      B2/Sankey/Sankey_Estudiantes_B2_Lengua.html/.png
  - 1 Excel con dos hojas (Matemática / Lengua):
      Excel/Trayectorias_Estudiantes_B1_B2.xlsx
"""

import os
import sys
import glob
import re
import unicodedata

import numpy as np
import pandas as pd
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
DIRECTORIOS_MESES = [
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados",  "Mes 1 (Marzo)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados",  "Mes 2 (Abril)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados",   "Mes 3 (Mayo)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\04_PROGRESO_Junio\Interim_CSVs\Resultados",  "Mes 4 (Junio)"),
]

PATH_METADATA = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_OUTPUT   = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\12. Evolución_Estudiantes_Mes1_4"

# Meses requeridos por grupo
MESES_B1 = [etiq for _, etiq in DIRECTORIOS_MESES]       # 4 meses
MESES_B2 = [etiq for _, etiq in DIRECTORIOS_MESES[-2:]]  # últimos 2

COL_SCORE = "theta.global (escala 0-100)"

NIVELES   = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
NIVEL_ORD = {n: i for i, n in enumerate(NIVELES)}   # Crítico=0 … Excelente=4

COLOR_NODO = ["#991b1b", "#ff8c2e", "#facc15", "#84cc16", "#065f46"]
COLOR_LINK = [
    "rgba(153, 27,  27, 0.35)",
    "rgba(255,140,  46, 0.35)",
    "rgba(250,204,  21, 0.35)",
    "rgba(132,204,  22, 0.35)",
    "rgba(  6, 95,  70, 0.35)",
]

# Crear carpetas de salida
for sub in ("B1/Sankey", "B2/Sankey", "Excel"):
    os.makedirs(os.path.join(PATH_OUTPUT, sub), exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# 2. HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def clasificar(val):
    if pd.isna(val):
        return ""
    v = float(val)
    if v <= 35:  return "Crítico"
    if v <= 45:  return "Bajo"
    if v <= 55:  return "Medio"
    if v <= 65:  return "Bueno"
    return "Excelente"


def quitar_tildes(s):
    return "".join(
        c for c in unicodedata.normalize("NFD", str(s))
        if unicodedata.category(c) != "Mn"
    )


def etiqueta_trayectoria(niveles_seq):
    """
    Recibe lista de strings de nivel en orden cronológico.
    Devuelve una de las 6 etiquetas de trayectoria.
    """
    ords = [NIVEL_ORD[n] for n in niveles_seq if n in NIVEL_ORD]
    if len(ords) < 2:
        return "Sin cambio de nivel"

    cambios = [ords[i + 1] - ords[i] for i in range(len(ords) - 1)]
    hubo_subida = any(c > 0 for c in cambios)
    hubo_bajada = any(c < 0 for c in cambios)
    hubo_plano  = any(c == 0 for c in cambios)

    if not hubo_subida and not hubo_bajada:
        return "Sin cambio de nivel"
    if hubo_subida and hubo_bajada:
        return "Trayectoria mixta"
    if hubo_subida:   # solo subidas (con o sin plano)
        return "Ascenso con estabilización" if hubo_plano else "Ascenso constante"
    # solo bajadas (con o sin plano)
    return "Descenso con estabilización" if hubo_plano else "Descenso constante"


# ══════════════════════════════════════════════════════════════════════════════
# 3. LECTURA DE CSVs
# ══════════════════════════════════════════════════════════════════════════════
def inferir_materia(nombre_archivo):
    """Infiere la materia a partir del nombre del archivo."""
    n = nombre_archivo.upper()
    if n.startswith("MAT"):
        return "Matemática"
    if n.startswith("LEC") or n.startswith("LEN"):
        return "Lengua"
    return None


def excluir_por_tiempo(val):
    """Devuelve True si la prueba debe anularse por tiempo insuficiente."""
    if pd.isna(val):
        return False
    v = quitar_tildes(str(val).lower())
    return "5 min" in v or "demora" in v or "duracion" in v


def leer_csvs_mes(ruta, etiqueta_mes):
    """
    Lee todos los CSVs de una carpeta y devuelve un DataFrame unificado.
    Columnas fijas que se conocen de los archivos reales:
      Departamento, Nro de centro, Centro, Grado, Grupo, Documento,
      Nombre, Apellido, anular_prueba, theta.global (escala 0-100)
    La materia se infiere del nombre del archivo (MAT* → Matemática, LEC* → Lengua).
    """
    archivos = glob.glob(os.path.join(ruta, "*.csv"))
    if not archivos:
        print(f"  [!] Sin CSVs en: {ruta}")
        return pd.DataFrame()

    lista = []
    for f in archivos:
        nombre = os.path.basename(f)
        if "legend" in nombre.lower():
            continue

        materia = inferir_materia(nombre)
        if materia is None:
            print(f"  [!] No se pudo inferir materia de '{nombre}' — omitido.")
            continue

        try:
            df = pd.read_csv(f, dtype=str, encoding_errors="ignore")
            df.columns = df.columns.str.strip()
        except Exception as e:
            print(f"  [!] Error leyendo {nombre}: {e}")
            continue

        if COL_SCORE not in df.columns:
            print(f"  [!] '{COL_SCORE}' no encontrada en {nombre} — omitido.")
            continue

        # ── Filtro de anulación por tiempo ──────────────────────────────────
        if "anular_prueba" in df.columns:
            df = df[~df["anular_prueba"].apply(excluir_por_tiempo)].copy()

        # ── Puntaje numérico ─────────────────────────────────────────────────
        df[COL_SCORE] = pd.to_numeric(
            df[COL_SCORE].astype(str).str.replace(",", "."), errors="coerce")
        df = df.dropna(subset=[COL_SCORE])
        if df.empty:
            continue

        # ── Normalizar Documento (quitar decimales tipo "12345.0") ───────────
        df["Documento"] = (df["Documento"].astype(str)
                           .str.replace(r"\.0$", "", regex=True).str.strip())

        # ── Excluir Centro Virtual ───────────────────────────────────────────
        df["Nro de centro"] = (df["Nro de centro"].astype(str)
                               .str.replace(r"\.0$", "", regex=True).str.strip())
        df = df[df["Nro de centro"] != "99999"]
        df = df[~df["Centro"].astype(str).str.contains(
            "Centro Virtual", case=False, na=False)]

        df["Materia"] = materia
        df["Mes"]     = etiqueta_mes

        cols_keep = [
            "Departamento", "Nro de centro", "Centro", "Grado", "Grupo",
            "Documento", "Nombre", "Apellido", COL_SCORE, "Materia", "Mes"
        ]
        # Solo conservar columnas que existan en este CSV
        cols_keep = [c for c in cols_keep if c in df.columns]
        lista.append(df[cols_keep])

    if not lista:
        return pd.DataFrame()

    resultado = pd.concat(lista, ignore_index=True)
    print(f"  [OK] {etiqueta_mes}: {len(resultado):,} registros "
          f"({resultado['Materia'].value_counts().to_dict()})")
    return resultado


# ══════════════════════════════════════════════════════════════════════════════
# 4. CARGAR MATRÍCULA (para obtener GRUPO_Limpio B1/B2)
# ══════════════════════════════════════════════════════════════════════════════
def cargar_matricula():
    print(f"\n[*] Cargando matrícula: {PATH_METADATA}")
    if not os.path.exists(PATH_METADATA):
        print("  [!] Archivo de matrícula no encontrado — se intentará inferir "
              "el grupo desde el campo 'Grupo' del CSV.")
        return pd.DataFrame()
    try:
        df = pd.read_csv(PATH_METADATA, sep=";", dtype=str, encoding_errors="ignore")
        df.columns = df.columns.str.strip().str.upper()
        df["NIE_Limpio"] = (df["NIE"].astype(str)
                            .str.replace(r"\.0$", "", regex=True).str.strip())
        df["GRUPO_Limpio"] = (
            df["GRUPO"].astype(str).str.strip().str.upper()
            if "GRUPO" in df.columns else "DESCONOCIDO")
        df = df[["NIE_Limpio", "GRUPO_Limpio"]].drop_duplicates(subset="NIE_Limpio")
        print(f"  [OK] {len(df):,} estudiantes en matrícula")
        return df
    except Exception as e:
        print(f"  [!] Error leyendo matrícula: {e}")
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# 5. CONSTRUIR TABLA WIDE CON FILTRO DE COMPLETITUD
# ══════════════════════════════════════════════════════════════════════════════
def construir_tabla_wide(df_long, meses_requeridos):
    """
    Pivotea df_long → una fila por (Documento × Materia).
    Solo conserva estudiantes con puntaje en TODOS los meses requeridos.
    Añade columna 'Trayectoria'.
    """
    df = df_long.copy()
    df["Nivel"] = df[COL_SCORE].apply(clasificar)

    # Pivot puntajes y niveles
    pivot_score = df.pivot_table(
        index=["Documento", "Materia"], columns="Mes",
        values=COL_SCORE, aggfunc="first")
    pivot_nivel = df.pivot_table(
        index=["Documento", "Materia"], columns="Mes",
        values="Nivel", aggfunc="first")

    pivot_score.columns = [f"Puntaje_{m}" for m in pivot_score.columns]
    pivot_nivel.columns = [f"Nivel_{m}"   for m in pivot_nivel.columns]

    df_wide = pd.concat([pivot_score, pivot_nivel], axis=1).reset_index()

    # Filtrar completitud
    score_cols_req = [f"Puntaje_{m}" for m in meses_requeridos]
    # Verificar que todas las columnas requeridas existen
    cols_faltantes = [c for c in score_cols_req if c not in df_wide.columns]
    if cols_faltantes:
        print(f"  [!] Meses sin datos: {cols_faltantes} — "
              f"puede que esos meses no tengan CSVs válidos.")
        return pd.DataFrame(), meses_requeridos

    mask = df_wide[score_cols_req].notna().all(axis=1)
    df_wide = df_wide[mask].copy()

    # Adjuntar metadatos del estudiante (del último mes disponible)
    # "Mes" se incluye temporalmente para ordenar y luego se descarta
    meta_cols = ["Documento", "Departamento", "Nro de centro",
                 "Centro", "Grado", "Grupo", "Nombre", "Apellido", "Materia"]
    meta_cols = [c for c in meta_cols if c in df.columns]
    meta_cols_con_mes = meta_cols + ["Mes"] if "Mes" in df.columns else meta_cols
    df_meta = (df[meta_cols_con_mes]
               .sort_values("Mes")
               .drop_duplicates(subset=["Documento", "Materia"], keep="last")
               .drop(columns=["Mes"], errors="ignore"))
    df_wide = pd.merge(df_wide, df_meta, on=["Documento", "Materia"], how="left")

    # Calcular trayectoria
    nivel_cols_ord = [f"Nivel_{m}" for m in meses_requeridos]
    df_wide["Trayectoria"] = df_wide.apply(
        lambda row: etiqueta_trayectoria(
            [row[c] for c in nivel_cols_ord
             if c in row.index and pd.notna(row[c]) and row[c] != ""]
        ), axis=1)

    return df_wide, meses_requeridos


# ══════════════════════════════════════════════════════════════════════════════
# 6. SANKEY POR ESTUDIANTES
# ══════════════════════════════════════════════════════════════════════════════
def calc_y_centers(conteos_nivel):
    totals   = [conteos_nivel.get(n, 0) for n in NIVELES]
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


def dibujar_sankey_estudiantes(df_wide, meses_req, materia, grupo, sufijo):
    print(f"  -> Sankey — {grupo} — {materia}  ({len(df_wide):,} estudiantes totales)")

    df_mat = df_wide[df_wide["Materia"] == materia].copy()
    if df_mat.empty:
        print(f"     [!] Sin datos para {materia} — omitido.")
        return

    nivel_cols = [f"Nivel_{m}" for m in meses_req]
    n_est = len(df_mat)

    # ── Nodos ────────────────────────────────────────────────────────────────
    labels_nodos, color_nodos, node_x, node_y = [], [], [], []

    for i, mes in enumerate(meses_req):
        x_val    = 0.01 + i * (0.98 / max(1, len(meses_req) - 1))
        col_niv  = f"Nivel_{mes}"
        conteos  = df_mat[col_niv].value_counts().to_dict() if col_niv in df_mat.columns else {}
        centros_y = calc_y_centers(conteos)
        for j, nivel in enumerate(NIVELES):
            n = conteos.get(nivel, 0)
            labels_nodos.append(f"{mes}\n{nivel}\n(n={n:,})")
            color_nodos.append(COLOR_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])

    # ── Links ────────────────────────────────────────────────────────────────
    sources, targets, values, link_colors = [], [], [], []

    # Links invisibles para anclar posición de nodos vacíos
    for i in range(len(meses_req) - 1):
        oo = i * len(NIVELES)
        od = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(oo + j); targets.append(od + j)
            values.append(1e-9);    link_colors.append("rgba(0,0,0,0)")

    # Links reales
    for i in range(len(meses_req) - 1):
        col_a = f"Nivel_{meses_req[i]}"
        col_b = f"Nivel_{meses_req[i + 1]}"
        if col_a not in df_mat.columns or col_b not in df_mat.columns:
            continue
        flujos = (df_mat.groupby([col_a, col_b])
                  .size().reset_index(name="n"))
        oo = i * len(NIVELES)
        od = (i + 1) * len(NIVELES)
        for _, row in flujos.iterrows():
            if row["n"] == 0:
                continue
            if row[col_a] not in NIVEL_ORD or row[col_b] not in NIVEL_ORD:
                continue
            si = oo + NIVEL_ORD[row[col_a]]
            ti = od + NIVEL_ORD[row[col_b]]
            sources.append(si);          targets.append(ti)
            values.append(int(row["n"])); link_colors.append(COLOR_LINK[NIVEL_ORD[row[col_a]]])

    # ── Subtítulo con resumen de trayectorias ────────────────────────────────
    tray_counts = df_mat["Trayectoria"].value_counts()
    orden_etiq  = ["Sin cambio de nivel", "Ascenso constante",
                   "Ascenso con estabilización", "Descenso constante",
                   "Descenso con estabilización", "Trayectoria mixta"]
    tray_lineas = "  ·  ".join(
        f"{et}: {tray_counts.get(et, 0):,} ({tray_counts.get(et, 0) / n_est * 100:.1f}%)"
        for et in orden_etiq if tray_counts.get(et, 0) > 0
    )
    meses_txt = f"{len(meses_req)} meses" if grupo == "B1" else "últimos 2 meses"
    subtitulo = (
        f"<span style='font-size:12px;color:#555'>"
        f"Grupo {grupo} — {n_est:,} estudiantes con datos completos ({meses_txt})<br>"
        f"{tray_lineas}</span>"
    )

    # ── Figura ───────────────────────────────────────────────────────────────
    fig = go.Figure(data=[go.Sankey(
        arrangement="fixed",
        node=dict(pad=18, thickness=22,
                  line=dict(color="white", width=0.4),
                  label=labels_nodos, color=color_nodos,
                  x=node_x, y=node_y),
        link=dict(source=sources, target=targets,
                  value=values,   color=link_colors)
    )])
    fig.update_layout(
        title=dict(
            text=(f"<span style='font-size:21px;color:#1a2b4c;"
                  f"font-family:Arial,sans-serif;font-weight:bold'>"
                  f"Trayectorias de nivel — Estudiantes — {materia}</span>"
                  f"<br>{subtitulo}"),
            x=0.01, y=0.96),
        font_size=11,
        width=max(950, 820 + len(meses_req) * 160),
        height=620,
        margin=dict(l=20, r=20, t=110, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )

    dir_out = os.path.join(PATH_OUTPUT, sufijo, "Sankey")
    base    = f"Sankey_Estudiantes_{sufijo}_{materia}"
    fig.write_html(os.path.join(dir_out, base + ".html"))
    try:
        fig.write_image(os.path.join(dir_out, base + ".png"), scale=2)
    except Exception:
        pass
    print(f"     [OK] {sufijo}/Sankey/{base}.html/.png")


# ══════════════════════════════════════════════════════════════════════════════
# 7. EXCEL DE TRAYECTORIAS
# ══════════════════════════════════════════════════════════════════════════════
def exportar_excel(df_b1_wide, df_b2_wide, meses_b1, meses_b2):
    print("\n[*] Exportando Excel de trayectorias...")

    def preparar_hoja(df_wide, meses_req, grupo_paars):
        if df_wide.empty:
            return pd.DataFrame()
        df = df_wide.copy()
        df["Grupo PAARS"] = grupo_paars

        # Pares Puntaje+Nivel intercalados en orden cronológico
        score_nivel = []
        for m in meses_req:
            pc = f"Puntaje_{m}"
            nc = f"Nivel_{m}"
            if pc in df.columns:
                score_nivel += [pc, nc]

        id_cols = ["Departamento", "Nro de centro", "Centro", "Grado",
                   "Grupo", "Documento", "Nombre", "Apellido", "Grupo PAARS"]
        id_cols = [c for c in id_cols if c in df.columns]
        return df[id_cols + score_nivel + ["Trayectoria"]]

    # Construir hoja Matemática y Lengua combinando B1 + B2
    hojas = {}
    for materia in ["Matemática", "Lengua"]:
        df_b1_m = df_b1_wide[df_b1_wide["Materia"] == materia] if not df_b1_wide.empty else pd.DataFrame()
        df_b2_m = df_b2_wide[df_b2_wide["Materia"] == materia] if not df_b2_wide.empty else pd.DataFrame()

        parte_b1 = preparar_hoja(df_b1_m, meses_b1, "B1")
        parte_b2 = preparar_hoja(df_b2_m, meses_b2, "B2")

        # Alinear columnas antes de concatenar (B1 tiene más columnas de mes que B2)
        all_cols = list(dict.fromkeys(
            list(parte_b1.columns) + list(parte_b2.columns)))
        parte_b1 = parte_b1.reindex(columns=all_cols)
        parte_b2 = parte_b2.reindex(columns=all_cols)

        hojas[materia] = pd.concat([parte_b1, parte_b2], ignore_index=True)

    out_path = os.path.join(PATH_OUTPUT, "Excel",
                            "Trayectorias_Estudiantes_B1_B2.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for nombre_hoja, df_hoja in hojas.items():
            if df_hoja.empty:
                pd.DataFrame({"Aviso": [f"Sin datos para {nombre_hoja}"]}).to_excel(
                    writer, sheet_name=nombre_hoja, index=False)
                continue
            df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False)
            ws = writer.sheets[nombre_hoja]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = max(
                    (len(str(cell.value)) if cell.value else 0)
                    for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 42)

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"  [OK] Excel → {out_path}  ({size_mb:.1f} MB)")
    for mat, df_h in hojas.items():
        print(f"       {mat}: {len(df_h):,} estudiantes")


# ══════════════════════════════════════════════════════════════════════════════
# 8. EJECUCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print(f"\n{'='*65}")
    print("  SANKEY DE TRAYECTORIAS POR ESTUDIANTE — B1 (4 meses) / B2 (2 meses)")
    print(f"{'='*65}\n")

    # ── 1. Leer todos los CSVs ────────────────────────────────────────────────
    print("[*] Leyendo CSVs de los 4 meses...")
    dfs_por_mes = []
    for ruta, etiqueta in DIRECTORIOS_MESES:
        if not os.path.exists(ruta):
            print(f"  [!] Ruta no encontrada: {ruta}")
            continue
        df_mes = leer_csvs_mes(ruta, etiqueta)
        if not df_mes.empty:
            dfs_por_mes.append(df_mes)

    if not dfs_por_mes:
        print("\n[!] No se leyó ningún CSV válido — proceso detenido.")
        sys.exit(1)

    df_long = pd.concat(dfs_por_mes, ignore_index=True)
    meses_leidos = df_long["Mes"].unique().tolist()
    print(f"\n[*] Total registros: {len(df_long):,}  |  Meses leídos: {meses_leidos}")

    # ── 2. Asignar GRUPO_Limpio (B1/B2) ──────────────────────────────────────
    df_mat_cula = cargar_matricula()
    if not df_mat_cula.empty:
        df_long = pd.merge(
            df_long,
            df_mat_cula.rename(columns={"NIE_Limpio": "Documento"}),
            on="Documento", how="left")
        df_long["GRUPO_Limpio"] = df_long["GRUPO_Limpio"].fillna("DESCONOCIDO")
    else:
        print("  [!] Sin matrícula — infiriendo grupo del campo 'Grupo' del CSV.")
        df_long["GRUPO_Limpio"] = (
            df_long["Grupo"].astype(str).str.upper().str.strip()
            .apply(lambda x: "B1" if "B1" in x else ("B2" if "B2" in x else "DESCONOCIDO")))

    # ── 3. Separar B1 y B2 ───────────────────────────────────────────────────
    df_b1_long = df_long[df_long["GRUPO_Limpio"] == "B1"].copy()
    df_b2_long = df_long[df_long["GRUPO_Limpio"] == "B2"].copy()
    print(f"\n[*] Registros B1: {len(df_b1_long):,}  |  B2: {len(df_b2_long):,}")

    # ── 4. Tablas wide con filtro de completitud ──────────────────────────────
    print("\n[*] Filtrando estudiantes con datos completos en todos sus meses...")
    df_b1_wide, _ = construir_tabla_wide(df_b1_long, MESES_B1)
    df_b2_wide, _ = construir_tabla_wide(df_b2_long, MESES_B2)

    n_b1 = df_b1_wide["Documento"].nunique() if not df_b1_wide.empty else 0
    n_b2 = df_b2_wide["Documento"].nunique() if not df_b2_wide.empty else 0
    print(f"  [OK] Estudiantes B1 con los 4 meses completos : {n_b1:,}")
    print(f"  [OK] Estudiantes B2 con los 2 últimos meses   : {n_b2:,}")

    # ── 5. Sankeys B1 ────────────────────────────────────────────────────────
    print("\n[*] Generando Sankeys B1...")
    if not df_b1_wide.empty:
        for materia in ["Matemática", "Lengua"]:
            dibujar_sankey_estudiantes(df_b1_wide, MESES_B1, materia, "B1", "B1")
    else:
        print("  [!] Sin datos B1 con los 4 meses — Sankeys B1 omitidos.")

    # ── 6. Sankeys B2 ────────────────────────────────────────────────────────
    print("\n[*] Generando Sankeys B2...")
    if not df_b2_wide.empty and n_b2 > 0:
        for materia in ["Matemática", "Lengua"]:
            dibujar_sankey_estudiantes(df_b2_wide, MESES_B2, materia, "B2", "B2")
    else:
        print("  [!] Sin datos B2 suficientes — Sankeys B2 omitidos.")

    # ── 7. Excel ──────────────────────────────────────────────────────────────
    exportar_excel(df_b1_wide, df_b2_wide, MESES_B1, MESES_B2)

    print(f"\n{'='*65}")
    print(f"  [OK] ¡Proceso completado!")
    print(f"       Salida: {PATH_OUTPUT}")
    print(f"{'='*65}\n")