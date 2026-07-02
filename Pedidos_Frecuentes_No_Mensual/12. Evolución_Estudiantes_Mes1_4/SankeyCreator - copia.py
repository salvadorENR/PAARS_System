"""
Sankey_Estudiantes_Evolutivo.py
================================
Genera Sankeys de trayectorias de nivel por ESTUDIANTE (no por escuela)
para Matemática y Lengua, separados por Grupo B1 y B2.

Salidas:
  A) Sankey B1 — 4 meses (solo estudiantes con los 4 meses completos)
       B1/Sankey/Sankey_Estudiantes_B1_Matemática.html/.png
       B1/Sankey/Sankey_Estudiantes_B1_Lengua.html/.png

  B) Sankey B1 — Mes 3→4 (mismos estudiantes B1 de 4 meses, solo 2 nodos)
       B1/Sankey/Sankey_Estudiantes_B1_Mes3a4_Matemática.html/.png
       B1/Sankey/Sankey_Estudiantes_B1_Mes3a4_Lengua.html/.png

  C) Sankey B2 — Mes 3→4 (solo estudiantes con Mes 3 y Mes 4 completos)
       B2/Sankey/Sankey_Estudiantes_B2_Matemática.html/.png
       B2/Sankey/Sankey_Estudiantes_B2_Lengua.html/.png

  D) Sankey B1 Mes3+4 — todos los estudiantes B1 con Mes 3 y Mes 4,
     sin importar si tomaron Mes 1 o Mes 2
       B1_Mes3_4/Sankey/Sankey_Estudiantes_B1_TodosMes3a4_Matemática.html/.png
       B1_Mes3_4/Sankey/Sankey_Estudiantes_B1_TodosMes3a4_Lengua.html/.png

  E) Excel con 4 meses (B1) + 2 meses (B2):
       Excel/Trayectorias_Estudiantes_B1_B2.xlsx  (hojas: Matemática / Lengua)

  F) Excel B1 Mes 3→4 (todos los B1 con Mes 3 y 4, sin filtro de Mes 1/2):
       Excel/Trayectorias_Mes3a4_B1_SinFiltroMes1y2.xlsx  (hojas: Matemática / Lengua)
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
MESES_B1      = [etiq for _, etiq in DIRECTORIOS_MESES]        # 4 meses
MESES_B2      = [etiq for _, etiq in DIRECTORIOS_MESES[-2:]]   # Mes 3 y Mes 4
MESES_3_4     = [etiq for _, etiq in DIRECTORIOS_MESES[-2:]]   # Mes 3 y Mes 4

COL_SCORE = "theta.global (escala 0-100)"

NIVELES   = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
NIVEL_ORD = {n: i for i, n in enumerate(NIVELES)}

COLOR_NODO = ["#991b1b", "#ff8c2e", "#facc15", "#84cc16", "#065f46"]
COLOR_LINK = [
    "rgba(153, 27,  27, 0.35)",
    "rgba(255,140,  46, 0.35)",
    "rgba(250,204,  21, 0.35)",
    "rgba(132,204,  22, 0.35)",
    "rgba(  6, 95,  70, 0.35)",
]

# Crear carpetas de salida
for sub in ("B1/Sankey", "B2/Sankey", "B1_Mes3_4/Sankey", "Excel"):
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
    6 etiquetas para secuencias de 3+ puntos (B1 4 meses).
    """
    ords = [NIVEL_ORD[n] for n in niveles_seq if n in NIVEL_ORD]
    if len(ords) < 2:
        return "Sin cambio de nivel"
    cambios    = [ords[i + 1] - ords[i] for i in range(len(ords) - 1)]
    hubo_sub   = any(c > 0 for c in cambios)
    hubo_baj   = any(c < 0 for c in cambios)
    hubo_plano = any(c == 0 for c in cambios)
    if not hubo_sub and not hubo_baj:
        return "Sin cambio de nivel"
    if hubo_sub and hubo_baj:
        return "Trayectoria mixta"
    if hubo_sub:
        return "Ascenso con estabilización" if hubo_plano else "Ascenso constante"
    return "Descenso con estabilización" if hubo_plano else "Descenso constante"


def etiqueta_cambio_simple(nivel_a, nivel_b):
    """
    4 etiquetas para una sola transición Mes3 → Mes4.
    """
    if nivel_a not in NIVEL_ORD or nivel_b not in NIVEL_ORD:
        return "Sin información"
    d = NIVEL_ORD[nivel_b] - NIVEL_ORD[nivel_a]
    if d > 0:  return "Subió de nivel"
    if d < 0:  return "Bajó de nivel"
    return "Se mantuvo en el mismo nivel"


def excluir_por_tiempo(val):
    if pd.isna(val):
        return False
    v = quitar_tildes(str(val).lower())
    return "5 min" in v or "demora" in v or "duracion" in v


# ══════════════════════════════════════════════════════════════════════════════
# 3. LECTURA DE CSVs
# ══════════════════════════════════════════════════════════════════════════════
def inferir_materia(nombre_archivo):
    n = nombre_archivo.upper()
    if n.startswith("MAT"):              return "Matemática"
    if n.startswith("LEC") or n.startswith("LEN"): return "Lengua"
    return None


def leer_csvs_mes(ruta, etiqueta_mes):
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
        if "anular_prueba" in df.columns:
            df = df[~df["anular_prueba"].apply(excluir_por_tiempo)].copy()
        df[COL_SCORE] = pd.to_numeric(
            df[COL_SCORE].astype(str).str.replace(",", "."), errors="coerce")
        df = df.dropna(subset=[COL_SCORE])
        if df.empty:
            continue
        df["Documento"]     = (df["Documento"].astype(str)
                               .str.replace(r"\.0$", "", regex=True).str.strip())
        df["Nro de centro"] = (df["Nro de centro"].astype(str)
                               .str.replace(r"\.0$", "", regex=True).str.strip())
        df = df[df["Nro de centro"] != "99999"]
        df = df[~df["Centro"].astype(str).str.contains(
            "Centro Virtual", case=False, na=False)]
        df["Materia"] = materia
        df["Mes"]     = etiqueta_mes
        cols_keep = [c for c in [
            "Departamento", "Nro de centro", "Centro", "Grado", "Grupo",
            "Documento", "Nombre", "Apellido", COL_SCORE, "Materia", "Mes"
        ] if c in df.columns]
        lista.append(df[cols_keep])

    if not lista:
        return pd.DataFrame()
    resultado = pd.concat(lista, ignore_index=True)
    print(f"  [OK] {etiqueta_mes}: {len(resultado):,} registros "
          f"({resultado['Materia'].value_counts().to_dict()})")
    return resultado


# ══════════════════════════════════════════════════════════════════════════════
# 4. CARGAR MATRÍCULA
# ══════════════════════════════════════════════════════════════════════════════
def cargar_matricula():
    print(f"\n[*] Cargando matrícula: {PATH_METADATA}")
    if not os.path.exists(PATH_METADATA):
        print("  [!] Archivo de matrícula no encontrado.")
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
        print(f"  [!] Error: {e}")
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# 5. CONSTRUIR TABLA WIDE
# ══════════════════════════════════════════════════════════════════════════════
def construir_tabla_wide(df_long, meses_requeridos, etiqueta_tray_fn=None):
    """
    Pivotea df_long → una fila por (Documento × Materia).
    Solo conserva estudiantes con puntaje en TODOS los meses requeridos.
    etiqueta_tray_fn: función(niveles_seq) → string. Si None usa etiqueta_trayectoria.
    """
    if etiqueta_tray_fn is None:
        etiqueta_tray_fn = etiqueta_trayectoria

    df = df_long.copy()
    df["Nivel"] = df[COL_SCORE].apply(clasificar)

    pivot_score = df.pivot_table(
        index=["Documento", "Materia"], columns="Mes",
        values=COL_SCORE, aggfunc="first")
    pivot_nivel = df.pivot_table(
        index=["Documento", "Materia"], columns="Mes",
        values="Nivel", aggfunc="first")

    pivot_score.columns = [f"Puntaje_{m}" for m in pivot_score.columns]
    pivot_nivel.columns = [f"Nivel_{m}"   for m in pivot_nivel.columns]

    df_wide = pd.concat([pivot_score, pivot_nivel], axis=1).reset_index()

    score_cols_req = [f"Puntaje_{m}" for m in meses_requeridos]
    cols_faltantes = [c for c in score_cols_req if c not in df_wide.columns]
    if cols_faltantes:
        print(f"  [!] Meses sin datos: {cols_faltantes}")
        return pd.DataFrame(), meses_requeridos

    mask    = df_wide[score_cols_req].notna().all(axis=1)
    df_wide = df_wide[mask].copy()

    meta_cols = [c for c in [
        "Documento", "Departamento", "Nro de centro",
        "Centro", "Grado", "Grupo", "Nombre", "Apellido", "Materia"
    ] if c in df.columns]
    meta_cols_con_mes = meta_cols + ["Mes"] if "Mes" in df.columns else meta_cols
    df_meta = (df[meta_cols_con_mes]
               .sort_values("Mes")
               .drop_duplicates(subset=["Documento", "Materia"], keep="last")
               .drop(columns=["Mes"], errors="ignore"))
    df_wide = pd.merge(df_wide, df_meta, on=["Documento", "Materia"], how="left")

    nivel_cols_ord = [f"Nivel_{m}" for m in meses_requeridos]
    df_wide["Trayectoria"] = df_wide.apply(
        lambda row: etiqueta_tray_fn(
            [row[c] for c in nivel_cols_ord
             if c in row.index and pd.notna(row[c]) and row[c] != ""]
        ), axis=1)

    return df_wide, meses_requeridos


def construir_tabla_wide_cambio_simple(df_long, mes_a, mes_b):
    """
    Versión para exactamente 2 meses: usa etiqueta_cambio_simple.
    """
    def _tray(seq):
        if len(seq) < 2:
            return "Sin información"
        return etiqueta_cambio_simple(seq[0], seq[1])
    return construir_tabla_wide(df_long, [mes_a, mes_b], etiqueta_tray_fn=_tray)


# ══════════════════════════════════════════════════════════════════════════════
# 6. SANKEY
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


def _construir_subtitulo_trayectorias(tray_counts, n_est, orden_etiq, desc_poblacion):
    """Genera el HTML del subtítulo con conteos y porcentajes."""
    tray_lineas = "  ·  ".join(
        f"{et}: {tray_counts.get(et, 0):,} ({tray_counts.get(et, 0) / n_est * 100:.1f}%)"
        for et in orden_etiq if tray_counts.get(et, 0) > 0
    )
    return (
        f"<span style='font-size:12px;color:#555'>"
        f"{desc_poblacion}<br>"
        f"{tray_lineas}</span>"
    )


def dibujar_sankey(df_wide, meses_req, materia, sufijo, dir_sankey,
                   orden_etiq, desc_poblacion, titulo_extra=""):
    """
    Función genérica de Sankey. Usada tanto para 4 meses como para 2.
    """
    df_mat = df_wide[df_wide["Materia"] == materia].copy()
    n_est  = len(df_mat)
    print(f"  -> Sankey {sufijo} — {materia}  ({n_est:,} estudiantes)")
    if df_mat.empty:
        print(f"     [!] Sin datos — omitido.")
        return

    # ── Nodos ────────────────────────────────────────────────────────────────
    labels_nodos, color_nodos, node_x, node_y = [], [], [], []
    for i, mes in enumerate(meses_req):
        x_val     = 0.01 + i * (0.98 / max(1, len(meses_req) - 1))
        col_niv   = f"Nivel_{mes}"
        conteos   = df_mat[col_niv].value_counts().to_dict() if col_niv in df_mat.columns else {}
        centros_y = calc_y_centers(conteos)
        for j, nivel in enumerate(NIVELES):
            n = conteos.get(nivel, 0)
            labels_nodos.append(f"{mes}\n{nivel}\n(n={n:,})")
            color_nodos.append(COLOR_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])

    # ── Links ────────────────────────────────────────────────────────────────
    sources, targets, values, link_colors = [], [], [], []
    # invisibles para anclar nodos vacíos
    for i in range(len(meses_req) - 1):
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(oo + j); targets.append(od + j)
            values.append(1e-9);    link_colors.append("rgba(0,0,0,0)")
    # reales
    for i in range(len(meses_req) - 1):
        col_a = f"Nivel_{meses_req[i]}"
        col_b = f"Nivel_{meses_req[i + 1]}"
        if col_a not in df_mat.columns or col_b not in df_mat.columns:
            continue
        flujos = df_mat.groupby([col_a, col_b]).size().reset_index(name="n")
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        for _, row in flujos.iterrows():
            if row["n"] == 0: continue
            if row[col_a] not in NIVEL_ORD or row[col_b] not in NIVEL_ORD: continue
            si = oo + NIVEL_ORD[row[col_a]]
            ti = od + NIVEL_ORD[row[col_b]]
            sources.append(si); targets.append(ti)
            values.append(int(row["n"]))
            link_colors.append(COLOR_LINK[NIVEL_ORD[row[col_a]]])

    # ── Subtítulo ─────────────────────────────────────────────────────────
    tray_counts = df_mat["Trayectoria"].value_counts()
    subtitulo   = _construir_subtitulo_trayectorias(
        tray_counts, n_est, orden_etiq, desc_poblacion)

    titulo_completo = (
        f"<span style='font-size:21px;color:#1a2b4c;"
        f"font-family:Arial,sans-serif;font-weight:bold'>"
        f"Trayectorias de nivel — Estudiantes — {materia}"
        f"{(' — ' + titulo_extra) if titulo_extra else ''}</span>"
        f"<br>{subtitulo}"
    )

    fig = go.Figure(data=[go.Sankey(
        arrangement="fixed",
        node=dict(pad=18, thickness=22,
                  line=dict(color="white", width=0.4),
                  label=labels_nodos, color=color_nodos,
                  x=node_x, y=node_y),
        link=dict(source=sources, target=targets,
                  value=values, color=link_colors)
    )])
    fig.update_layout(
        title=dict(text=titulo_completo, x=0.01, y=0.96),
        font_size=11,
        width=max(950, 820 + len(meses_req) * 160),
        height=620,
        margin=dict(l=20, r=20, t=110, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )

    base = f"Sankey_Estudiantes_{sufijo}_{materia}"
    fig.write_html(os.path.join(dir_sankey, base + ".html"))
    try:
        fig.write_image(os.path.join(dir_sankey, base + ".png"), scale=2)
    except Exception:
        pass
    print(f"     [OK] {sufijo}/{base}.html/.png")


# ══════════════════════════════════════════════════════════════════════════════
# 7. EXCEL
# ══════════════════════════════════════════════════════════════════════════════
ORDEN_ETIQ_6 = [
    "Sin cambio de nivel", "Ascenso constante", "Ascenso con estabilización",
    "Descenso constante", "Descenso con estabilización", "Trayectoria mixta"
]
ORDEN_ETIQ_4 = [
    "Subió de nivel", "Se mantuvo en el mismo nivel",
    "Bajó de nivel", "Sin información"
]


def _preparar_hoja(df_wide, meses_req, grupo_paars):
    if df_wide.empty:
        return pd.DataFrame()
    df = df_wide.copy()
    df["Grupo PAARS"] = grupo_paars
    score_nivel = []
    for m in meses_req:
        pc = f"Puntaje_{m}"; nc = f"Nivel_{m}"
        if pc in df.columns:
            score_nivel += [pc, nc]
    id_cols = [c for c in [
        "Departamento", "Nro de centro", "Centro", "Grado",
        "Grupo", "Documento", "Nombre", "Apellido", "Grupo PAARS"
    ] if c in df.columns]
    return df[id_cols + score_nivel + ["Trayectoria"]]


def exportar_excel(df_b1_wide, df_b2_wide, meses_b1, meses_b2,
                   nombre_archivo="Trayectorias_Estudiantes_B1_B2.xlsx"):
    print(f"\n[*] Exportando {nombre_archivo}...")
    hojas = {}
    for materia in ["Matemática", "Lengua"]:
        df_b1_m = df_b1_wide[df_b1_wide["Materia"] == materia] if not df_b1_wide.empty else pd.DataFrame()
        df_b2_m = df_b2_wide[df_b2_wide["Materia"] == materia] if not df_b2_wide.empty else pd.DataFrame()
        parte_b1 = _preparar_hoja(df_b1_m, meses_b1, "B1")
        parte_b2 = _preparar_hoja(df_b2_m, meses_b2, "B2")
        all_cols = list(dict.fromkeys(list(parte_b1.columns) + list(parte_b2.columns)))
        hojas[materia] = pd.concat(
            [parte_b1.reindex(columns=all_cols),
             parte_b2.reindex(columns=all_cols)],
            ignore_index=True)

    out_path = os.path.join(PATH_OUTPUT, "Excel", nombre_archivo)
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
    print(f"  [OK] → {out_path}  ({size_mb:.1f} MB)")
    for mat, df_h in hojas.items():
        print(f"       {mat}: {len(df_h):,} estudiantes")


# ══════════════════════════════════════════════════════════════════════════════
# 8. EJECUCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print(f"\n{'='*65}")
    print("  SANKEY DE TRAYECTORIAS POR ESTUDIANTE")
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

    MES_3 = MESES_3_4[0]   # "Mes 3 (Mayo)"
    MES_4 = MESES_3_4[1]   # "Mes 4 (Junio)"

    # ══════════════════════════════════════════════════════════════════════════
    # A) B1 — 4 MESES COMPLETOS
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[*] (A) B1 — estudiantes con los 4 meses completos...")
    df_b1_wide, _ = construir_tabla_wide(df_b1_long, MESES_B1)
    n_b1 = df_b1_wide["Documento"].nunique() if not df_b1_wide.empty else 0
    print(f"  [OK] {n_b1:,} estudiantes B1 con 4 meses")

    dir_b1 = os.path.join(PATH_OUTPUT, "B1", "Sankey")
    if not df_b1_wide.empty:
        for materia in ["Matemática", "Lengua"]:
            dibujar_sankey(
                df_wide       = df_b1_wide,
                meses_req     = MESES_B1,
                materia       = materia,
                sufijo        = "B1",
                dir_sankey    = dir_b1,
                orden_etiq    = ORDEN_ETIQ_6,
                desc_poblacion= f"Grupo B1 — {n_b1:,} estudiantes con datos completos (4 meses)",
            )

    # ══════════════════════════════════════════════════════════════════════════
    # B) B1 — SOLO MES 3 → MES 4  (mismos estudiantes del filtro de 4 meses)
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[*] (B) B1 — Sankey Mes 3→4 (subconjunto de los 4 meses)...")
    if not df_b1_wide.empty:
        # Reutilizar df_b1_wide pero calcular Trayectoria con cambio simple
        df_b1_mes34 = df_b1_wide.copy()
        col_n3 = f"Nivel_{MES_3}"; col_n4 = f"Nivel_{MES_4}"
        if col_n3 in df_b1_mes34.columns and col_n4 in df_b1_mes34.columns:
            df_b1_mes34["Trayectoria"] = df_b1_mes34.apply(
                lambda row: etiqueta_cambio_simple(row[col_n3], row[col_n4]), axis=1)
            n_b1_34 = len(df_b1_mes34[df_b1_mes34["Materia"] == "Matemática"])
            for materia in ["Matemática", "Lengua"]:
                n_mat = len(df_b1_mes34[df_b1_mes34["Materia"] == materia])
                dibujar_sankey(
                    df_wide       = df_b1_mes34,
                    meses_req     = MESES_3_4,
                    materia       = materia,
                    sufijo        = "B1_Mes3a4",
                    dir_sankey    = dir_b1,
                    orden_etiq    = ORDEN_ETIQ_4,
                    desc_poblacion= (f"Grupo B1 — {n_mat:,} estudiantes con 4 meses completos "
                                     f"(vista Mes 3→4)"),
                    titulo_extra  = "Mes 3 → Mes 4",
                )
        else:
            print(f"  [!] Columnas Mes 3/4 no disponibles en df_b1_wide.")

    # ══════════════════════════════════════════════════════════════════════════
    # C) B2 — MES 3 Y MES 4
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[*] (C) B2 — estudiantes con Mes 3 y Mes 4...")
    df_b2_wide, _ = construir_tabla_wide_cambio_simple(df_b2_long, MES_3, MES_4)
    n_b2 = df_b2_wide["Documento"].nunique() if not df_b2_wide.empty else 0
    print(f"  [OK] {n_b2:,} estudiantes B2 con Mes 3 y Mes 4")

    dir_b2 = os.path.join(PATH_OUTPUT, "B2", "Sankey")
    if not df_b2_wide.empty:
        for materia in ["Matemática", "Lengua"]:
            n_mat = len(df_b2_wide[df_b2_wide["Materia"] == materia])
            dibujar_sankey(
                df_wide       = df_b2_wide,
                meses_req     = MESES_3_4,
                materia       = materia,
                sufijo        = "B2",
                dir_sankey    = dir_b2,
                orden_etiq    = ORDEN_ETIQ_4,
                desc_poblacion= (f"Grupo B2 — {n_mat:,} estudiantes con "
                                 f"Mes 3 y Mes 4 completos"),
            )
    else:
        print("  [!] Sin datos B2 — Sankeys B2 omitidos.")

    # ══════════════════════════════════════════════════════════════════════════
    # D) B1 — todos los estudiantes con Mes 3 y Mes 4, sin importar Mes 1/2
    #    Criterio: ser de grupo B1 Y tener puntaje en Mes 3 Y Mes 4.
    #    NO se exige que hayan tomado Mes 1 o Mes 2.
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[*] (D) B1 — todos los que tomaron Mes 3 y Mes 4 (sin importar Mes 1/2)...")
    # Solo registros B1 de Mes 3 y Mes 4
    df_b1_34_long = df_b1_long[df_b1_long["Mes"].isin(MESES_3_4)].copy()
    df_b1_34_wide, _ = construir_tabla_wide_cambio_simple(df_b1_34_long, MES_3, MES_4)
    n_b1_34 = df_b1_34_wide["Documento"].nunique() if not df_b1_34_wide.empty else 0
    print(f"  [OK] {n_b1_34:,} estudiantes B1 con Mes 3 y Mes 4 (sin filtro Mes 1/2)")

    dir_todos = os.path.join(PATH_OUTPUT, "B1_Mes3_4", "Sankey")
    if not df_b1_34_wide.empty:
        for materia in ["Matemática", "Lengua"]:
            n_mat = len(df_b1_34_wide[df_b1_34_wide["Materia"] == materia])
            dibujar_sankey(
                df_wide       = df_b1_34_wide,
                meses_req     = MESES_3_4,
                materia       = materia,
                sufijo        = "B1_TodosMes3a4",
                dir_sankey    = dir_todos,
                orden_etiq    = ORDEN_ETIQ_4,
                desc_poblacion= (f"Grupo B1 — {n_mat:,} estudiantes con Mes 3 y Mes 4 "
                                 f"(independiente de Mes 1/2)"),
                titulo_extra  = "Mes 3 → Mes 4",
            )
    else:
        print("  [!] Sin datos B1 Mes 3+4 — omitido.")

    # ══════════════════════════════════════════════════════════════════════════
    # E) Excel B1 (4 meses) + B2 (Mes 3 y 4)
    # ══════════════════════════════════════════════════════════════════════════
    exportar_excel(
        df_b1_wide    = df_b1_wide,
        df_b2_wide    = df_b2_wide,
        meses_b1      = MESES_B1,
        meses_b2      = MESES_3_4,
        nombre_archivo= "Trayectorias_Estudiantes_B1_B2.xlsx"
    )

    # ══════════════════════════════════════════════════════════════════════════
    # F) Excel B1 Mes 3→4 (todos los B1 con Mes 3 y 4, sin filtro de Mes 1/2)
    # ══════════════════════════════════════════════════════════════════════════
    exportar_excel(
        df_b1_wide    = df_b1_34_wide,
        df_b2_wide    = pd.DataFrame(),   # B2 no aplica en este grupo
        meses_b1      = MESES_3_4,
        meses_b2      = MESES_3_4,
        nombre_archivo= "Trayectorias_Mes3a4_B1_SinFiltroMes1y2.xlsx"
    )

    print(f"\n{'='*65}")
    print(f"  [OK] ¡Proceso completado!")
    print(f"       Salida: {PATH_OUTPUT}")
    print(f"{'='*65}\n")