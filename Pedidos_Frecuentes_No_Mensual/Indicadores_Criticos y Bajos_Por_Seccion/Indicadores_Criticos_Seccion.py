"""
Script: actualizar_indicadores.py
Propósito: Agrega/reemplaza TRES hojas nuevas en el archivo de salida:
  1. "Indicadores_Críticos_SecLXP"   – indicadores Crítico/Bajo por sección y mes
                                        (Mes 1 Marzo, Mes 2 Abril, Mes 3 Mayo),
                                        separados por Matemática y Lengua.
  2. "Estudiantes_SIGES_LXP"         – datos del estudiante + Código Sección SIGES
                                        + columna Sección (letra) después de Grado.
  3. "Diccionario_Sección_SIGES_LXP" – tabla-diccionario única por sección con
                                        Código, CE, Tipo, Grado, Sección, Cód LXP, Cód SIGES.

Requiere: pip install openpyxl beautifulsoup4 lxml
"""

import re
import pathlib
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup, Tag, NavigableString
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
#  RUTAS
# ══════════════════════════════════════════════════════════════════════════════
EXCEL_PATH = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Criterio_Inclusión_Exclusión_Escuelas"
    r"\Centros_Escolares_B1.xlsx"
)

OUTPUT_PATH = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Indicadores_Criticos y Bajos_Por_Seccion"
    r"\Indicadores_Críticos_Bajos_Mes 3_GB1.xlsx"
)

# Un directorio de HTMLs por mes  (clave = etiqueta que aparece en el Excel)
HTML_ROOTS = {
    "Mes 1 (Marzo)": (
        r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis"
        r"\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Final_Reports\Reportes_Por_Secciones"
    ),
    "Mes 2 (Abril)": (
        r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis"
        r"\PAARS_Warehouse\2026\02_PROGRESO_Abril\Final_Reports\Reportes_Por_Secciones"
    ),
    "Mes 3 (Mayo)": (
        r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis"
        r"\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Final_Reports\Reportes_Por_Secciones"
    ),
}

# Etiquetas cortas para encabezados de columna
MES_LABELS = {
    "Mes 1 (Marzo)": "M1",
    "Mes 2 (Abril)": "M2",
    "Mes 3 (Mayo)":  "M3",
}

# ══════════════════════════════════════════════════════════════════════════════
#  COLORES DE ESTATUS
# ══════════════════════════════════════════════════════════════════════════════
STATUS_BG = {
    "Excelente": "1E7145",
    "Bueno":     "70AD47",
    "Regular":   "FFD966",
    "Alerta":    "FF4B4B",
}
STATUS_FG = {
    "Excelente": "FFFFFF",
    "Bueno":     "FFFFFF",
    "Regular":   "000000",
    "Alerta":    "FFFFFF",
}

# ══════════════════════════════════════════════════════════════════════════════
#  ESTILOS GENERALES
# ══════════════════════════════════════════════════════════════════════════════
FILL_DARK   = PatternFill("solid", start_color="1F4E79")
FILL_MAT    = PatternFill("solid", start_color="2E75B6")
FILL_LEN    = PatternFill("solid", start_color="375623")
FILL_GROUP  = PatternFill("solid", start_color="404040")
FILL_DICT   = PatternFill("solid", start_color="2E4057")
FILL_M1     = PatternFill("solid", start_color="1F618D")   # azul M1
FILL_M2     = PatternFill("solid", start_color="2874A6")   # azul M2
FILL_M3     = PatternFill("solid", start_color="2E86C1")   # azul M3
FILL_M1_LEN = PatternFill("solid", start_color="1E8449")   # verde M1
FILL_M2_LEN = PatternFill("solid", start_color="27AE60")   # verde M2
FILL_M3_LEN = PatternFill("solid", start_color="2ECC71")   # verde M3
WHITE_BOLD  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
NORMAL_F    = Font(name="Arial", size=10)
THIN        = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def hdr(ws, row, col, value, fill):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = fill; c.font = WHITE_BOLD; c.border = THIN; c.alignment = CENTER
    return c


def dat(ws, row, col, value, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = NORMAL_F; c.border = THIN
    c.alignment = CENTER if align == "center" else LEFT
    return c


def status_dat(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value or "")
    if value and value in STATUS_BG:
        c.fill = PatternFill("solid", start_color=STATUS_BG[value])
        c.font = Font(name="Arial", size=10, bold=True, color=STATUS_FG[value])
    else:
        c.font = NORMAL_F
    c.border = THIN; c.alignment = CENTER
    return c


# ══════════════════════════════════════════════════════════════════════════════
#  PARSEAR UN HTML
# ══════════════════════════════════════════════════════════════════════════════
def parse_html(path: Path) -> dict:
    result = {
        "grado": None, "grupo": None, "seccion_letra": None, "codigo_siges": None,
        "nies_seccion": set(),
        "indicadores_seccion": {
            "mat_criticos": [], "mat_bajos": [],
            "len_criticos": [], "len_bajos": []
        },
        "nie_nivel": {}
    }
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            soup = BeautifulSoup(f, "lxml")

        # Encabezado
        for item in soup.select("div.info-item"):
            label = item.find("strong")
            value = item.find("span")
            if not label or not value:
                continue
            lbl = label.get_text(strip=True).lower()
            val = value.get_text(strip=True)
            if "grado" in lbl:
                result["grado"] = val
            elif "grupo" in lbl:
                result["grupo"] = val
                m = re.match(r"([A-Za-z])\s*\((\d+)\)", val)
                if m:
                    result["seccion_letra"] = m.group(1).upper()
                    result["codigo_siges"]  = m.group(2)

        # Indicadores Crítico / Bajo por asignatura
        for subject_div in soup.select("div.subject-section"):
            title_el = subject_div.find(class_="subject-title")
            if not title_el:
                continue
            title_text = title_el.get_text(strip=True).lower()
            is_mat = "matem" in title_text
            is_len = ("lectura" in title_text or "lengua" in title_text
                      or "español" in title_text)

            for ind_item in subject_div.select("div.indicator-item"):
                nivel_el = ind_item.select_one("div.pct-label")
                if not nivel_el:
                    continue
                # Normalize: some HTMLs use "Crítico" (with accent), others "Critico" (without)
                nivel = nivel_el.get_text(strip=True).replace("Crítico", "Critico")

                # Texto del indicador SIN mutar el DOM
                parts = []
                for child in ind_item.children:
                    if isinstance(child, Tag):
                        if "indicator-pct" in child.get("class", []):
                            continue
                        parts.append(child.get_text(" ", strip=True))
                    elif isinstance(child, NavigableString):
                        t = str(child).strip()
                        if t:
                            parts.append(t)
                texto = " ".join(p for p in parts if p).strip()

                if nivel == "Critico":
                    if is_mat: result["indicadores_seccion"]["mat_criticos"].append(texto)
                    if is_len: result["indicadores_seccion"]["len_criticos"].append(texto)
                elif nivel == "Bajo":
                    if is_mat: result["indicadores_seccion"]["mat_bajos"].append(texto)
                    if is_len: result["indicadores_seccion"]["len_bajos"].append(texto)

            # NIEs y niveles individuales
            for tbody in subject_div.find_all("tbody"):
                tbody_id  = tbody.get("id", "").lower()
                mat_table = "matem" in tbody_id
                len_table = "lectura" in tbody_id or "lengua" in tbody_id
                for tr in tbody.find_all("tr"):
                    tds = tr.find_all("td")
                    if len(tds) < 2:
                        continue
                    nie_str = tds[0].get_text(strip=True)
                    if not nie_str.isdigit():
                        continue
                    result["nies_seccion"].add(nie_str)
                    nivel_text  = tds[1].get_text(" ", strip=True)
                    nivel_match = re.search(r"NIVEL:\s*(\w+)", nivel_text, re.IGNORECASE)
                    nivel_ind   = nivel_match.group(1) if nivel_match else ""
                    entry = result["nie_nivel"].setdefault(nie_str, {"mat": "", "len": ""})
                    if mat_table and not entry["mat"]:
                        entry["mat"] = nivel_ind
                    if len_table and not entry["len"]:
                        entry["len"] = nivel_ind

    except Exception as e:
        print(f"  ⚠  Error parseando {path.name}: {e}")
    return result


# ══════════════════════════════════════════════════════════════════════════════
#  ESCANEAR UN DIRECTORIO DE HTML  →  {nie: reporte}
# ══════════════════════════════════════════════════════════════════════════════
def scan_htmls(html_root: str) -> dict:
    root = Path(html_root)
    if not root.exists():
        print(f"    ⚠  Directorio no encontrado: {html_root}")
        return {}
    html_files = sorted(root.rglob("*.html")) + sorted(root.rglob("*.htm"))
    print(f"    HTMLs encontrados: {len(html_files)}")
    nie_map = {}
    for hf in html_files:
        data = parse_html(hf)
        data["html_file"] = str(hf)
        for nie in data["nies_seccion"]:
            nie_map[nie] = data
    return nie_map


# ══════════════════════════════════════════════════════════════════════════════
#  LEER EL EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def read_excel(excel_path: str):
    def flatten_cols(df):
        new_cols = []
        for a, b in df.columns:
            a = str(a).strip() if not str(a).startswith("Unnamed") else ""
            b = str(b).strip() if not str(b).startswith("Unnamed") else ""
            new_cols.append((a + " " + b).strip())
        df.columns = new_cols
        return df

    df_est = flatten_cols(pd.read_excel(excel_path, sheet_name="Estudiantes",
                                         header=[0, 1], dtype=str))
    df_ce  = flatten_cols(pd.read_excel(excel_path, sheet_name="Centros Escolares",
                                         header=[0, 1], dtype=str))
    return df_est, df_ce


def find_col(df, *kws):
    for c in df.columns:
        if all(k.lower() in c.lower() for k in kws):
            return c
    return None


# ══════════════════════════════════════════════════════════════════════════════
#  CONSTRUIR DATOS PARA LAS TRES HOJAS
# ══════════════════════════════════════════════════════════════════════════════
def build_data(df_est, df_ce, nie_maps: dict):
    """
    nie_maps: {"Mes 1 (Marzo)": {nie: reporte}, "Mes 2 (Abril)": ..., "Mes 3 (Mayo)": ...}
    """
    meses = list(HTML_ROOTS.keys())   # orden fijo: M1, M2, M3

    # Estatus de centro escolar
    col_ce_cod = find_col(df_ce, "Código")
    col_ce_est = find_col(df_ce, "Estatus")
    ce_estatus = {}
    if col_ce_cod and col_ce_est:
        for _, r in df_ce.iterrows():
            ce_estatus[str(r[col_ce_cod]).strip()] = str(r[col_ce_est]).strip()

    # Columnas clave
    c_cod   = find_col(df_est, "Código")              or df_est.columns[0]
    c_ce    = find_col(df_est, "Centro", "Escolar")
    c_tipo  = find_col(df_est, "Tipo")
    c_nie   = find_col(df_est, "NIE")
    c_pnom  = find_col(df_est, "Primer",  "Nombre")
    c_snom  = find_col(df_est, "Segundo", "Nombre")
    c_pape  = find_col(df_est, "Primer",  "Apellido")
    c_sape  = find_col(df_est, "Segundo", "Apellido")
    c_grado = find_col(df_est, "Grado")
    c_lxp   = find_col(df_est, "Sección", "LXP") or find_col(df_est, "Seccion", "LXP")

    rows_ind  = []
    rows_est  = []
    rows_dict = []

    g_cols = [c for c in [c_cod, c_ce, c_lxp] if c]
    if not g_cols:
        print("  ⚠  No se encontraron columnas de agrupación.")
        return [], [], []

    for keys, grp in df_est.fillna("").groupby(g_cols, sort=False):
        keys      = keys if isinstance(keys, tuple) else (keys,)
        codigo_ce = str(keys[0]).strip() if len(keys) > 0 else ""
        nombre_ce = str(keys[1]).strip() if len(keys) > 1 else ""
        cod_lxp   = str(keys[2]).strip() if len(keys) > 2 else ""

        estatus_ce = ce_estatus.get(codigo_ce, "")

        # Indicadores por mes: {mes: {mat_criticos, mat_bajos, len_criticos, len_bajos}}
        ind_por_mes = {
            mes: {"mat_criticos": set(), "mat_bajos": set(),
                  "len_criticos": set(), "len_bajos": set()}
            for mes in meses
        }

        cod_siges = grado_html = sec_letra = tipo_ce = None

        for _, stu in grp.iterrows():
            nie_raw = str(stu.get(c_nie, "")).strip() if c_nie else ""
            nie_str = nie_raw.split(".")[0]

            if not tipo_ce and c_tipo:
                tipo_ce = str(stu.get(c_tipo, "")).strip()

            for mes in meses:
                html_data = nie_maps[mes].get(nie_str, {})
                if not html_data:
                    continue

                # Metadatos de sección (primer NIE/mes que tenga HTML)
                if not cod_siges:
                    cod_siges  = html_data.get("codigo_siges")
                    grado_html = html_data.get("grado")
                    sec_letra  = html_data.get("seccion_letra")

                ind = html_data.get("indicadores_seccion", {})
                ind_por_mes[mes]["mat_criticos"].update(ind.get("mat_criticos", []))
                ind_por_mes[mes]["mat_bajos"].update(ind.get("mat_bajos",    []))
                ind_por_mes[mes]["len_criticos"].update(ind.get("len_criticos", []))
                ind_por_mes[mes]["len_bajos"].update(ind.get("len_bajos",    []))

            # Fila Estudiantes_SIGES_LXP
            # Tomar código SIGES del primer mes disponible para este NIE
            nie_siges = cod_siges
            if not nie_siges:
                for mes in meses:
                    hd = nie_maps[mes].get(nie_str, {})
                    if hd.get("codigo_siges"):
                        nie_siges = hd["codigo_siges"]
                        break

            rows_est.append({
                "Código":               str(stu.get(c_cod,  "")).strip() if c_cod  else "",
                "Centro Escolar":       nombre_ce,
                "Tipo de Centro":       str(stu.get(c_tipo, "")).strip() if c_tipo else "",
                "NIE":                  nie_str,
                "Primer Nombre":        str(stu.get(c_pnom, "")).strip() if c_pnom else "",
                "Segundo Nombre":       str(stu.get(c_snom, "")).strip() if c_snom else "",
                "Primer Apellido":      str(stu.get(c_pape, "")).strip() if c_pape else "",
                "Segundo Apellido":     str(stu.get(c_sape, "")).strip() if c_sape else "",
                "Grado":                str(stu.get(c_grado,"")).strip() if c_grado else "",
                "Sección":              sec_letra or "",
                "Código Sección LXP":   cod_lxp,
                "Código Sección SIGES": nie_siges or "",
            })

        # Fila Indicadores_Críticos_SecLXP  (una fila por sección, 6 cols por materia)
        row_ind = {
            "Código":                 codigo_ce,
            "Centro Escolar":         nombre_ce,
            "Estatus Centro Escolar": estatus_ce,
            "Código Sección LXP":     cod_lxp,
            "Código Sección SIGES":   cod_siges or "",
            "Grado":                  grado_html or "",
            "Sección":                sec_letra or "",
        }
        for mes in meses:
            lbl = MES_LABELS[mes]
            row_ind[f"Mat – Críticos {lbl}"] = "\n".join(sorted(ind_por_mes[mes]["mat_criticos"])) if ind_por_mes[mes]["mat_criticos"] else ""
            row_ind[f"Mat – Bajos {lbl}"]    = "\n".join(sorted(ind_por_mes[mes]["mat_bajos"]))    if ind_por_mes[mes]["mat_bajos"]    else ""
            row_ind[f"Len – Críticos {lbl}"] = "\n".join(sorted(ind_por_mes[mes]["len_criticos"])) if ind_por_mes[mes]["len_criticos"] else ""
            row_ind[f"Len – Bajos {lbl}"]    = "\n".join(sorted(ind_por_mes[mes]["len_bajos"]))    if ind_por_mes[mes]["len_bajos"]    else ""
        rows_ind.append(row_ind)

        # Fila Diccionario
        rows_dict.append({
            "Código":               codigo_ce,
            "Centro Escolar":       nombre_ce,
            "Tipo de Centro":       tipo_ce or "",
            "Grado":                grado_html or "",
            "Sección":              sec_letra or "",
            "Código Sección LXP":   cod_lxp,
            "Código Sección SIGES": cod_siges or "",
        })

    # Deduplicar diccionario
    seen = set()
    rows_dict_dedup = []
    for r in rows_dict:
        key = (r["Código"], r["Código Sección LXP"])
        if key not in seen:
            seen.add(key)
            rows_dict_dedup.append(r)

    return rows_ind, rows_est, rows_dict_dedup


# ══════════════════════════════════════════════════════════════════════════════
#  HOJA 1: Indicadores_Críticos_SecLXP
#  Estructura de encabezado (3 filas):
#    Fila 1: Info | Matemática (M1+M2+M3) | Lengua (M1+M2+M3)
#    Fila 2:       M1          M2          M3  |  M1   M2   M3
#    Fila 3: columnas individuales
# ══════════════════════════════════════════════════════════════════════════════
def write_indicadores(wb, rows):
    sname = "Indicadores_Críticos_SecLXP"
    if sname in wb.sheetnames:
        del wb[sname]
    ws = wb.create_sheet(sname)

    meses  = list(HTML_ROOTS.keys())
    labels = [MES_LABELS[m] for m in meses]   # M1, M2, M3

    INFO_COLS = ["Código", "Centro Escolar", "Estatus Centro Escolar",
                 "Código Sección LXP", "Código Sección SIGES", "Grado", "Sección"]

    # Sub-columnas por mes: Críticos + Bajos
    MAT_SUB = []
    LEN_SUB = []
    for lbl in labels:
        MAT_SUB += [f"Mat – Críticos {lbl}", f"Mat – Bajos {lbl}"]
        LEN_SUB += [f"Len – Críticos {lbl}", f"Len – Bajos {lbl}"]

    ALL_COLS = INFO_COLS + MAT_SUB + LEN_SUB
    n_info   = len(INFO_COLS)
    n_sub    = len(labels) * 2   # cols por bloque (Mat o Len)

    mat_start = n_info + 1
    mat_end   = n_info + n_sub
    len_start = mat_end + 1
    len_end   = mat_end + n_sub

    # ── Fila 1: Info (merge rows 1-3) | Matemática | Lengua ─────────────
    for i, c in enumerate(INFO_COLS, 1):
        ws.merge_cells(start_row=1, start_column=i, end_row=3, end_column=i)
        hdr(ws, 1, i, c, FILL_GROUP)

    ws.merge_cells(start_row=1, start_column=mat_start,
                   end_row=1,   end_column=mat_end)
    hdr(ws, 1, mat_start, "Matemática", FILL_MAT)

    ws.merge_cells(start_row=1, start_column=len_start,
                   end_row=1,   end_column=len_end)
    hdr(ws, 1, len_start, "Lengua", FILL_LEN)

    # ── Fila 2: M1 | M2 | M3  (merge 2 cols cada uno) ───────────────────
    MES_FILLS_MAT = [FILL_M1, FILL_M2, FILL_M3]
    MES_FILLS_LEN = [FILL_M1_LEN, FILL_M2_LEN, FILL_M3_LEN]
    MES_NAMES = ["Mes 1 (Marzo)", "Mes 2 (Abril)", "Mes 3 (Mayo)"]

    for mi, (mes_name, f_mat, f_len) in enumerate(zip(MES_NAMES, MES_FILLS_MAT, MES_FILLS_LEN)):
        col_m = mat_start + mi * 2
        col_l = len_start + mi * 2
        ws.merge_cells(start_row=2, start_column=col_m, end_row=2, end_column=col_m + 1)
        hdr(ws, 2, col_m, mes_name, f_mat)
        ws.merge_cells(start_row=2, start_column=col_l, end_row=2, end_column=col_l + 1)
        hdr(ws, 2, col_l, mes_name, f_len)

    # ── Fila 3: Críticos / Bajos por cada mes ───────────────────────────
    for mi, (f_mat, f_len) in enumerate(zip(MES_FILLS_MAT, MES_FILLS_LEN)):
        col_m = mat_start + mi * 2
        col_l = len_start + mi * 2
        hdr(ws, 3, col_m,     "Críticos", f_mat)
        hdr(ws, 3, col_m + 1, "Bajos",    f_mat)
        hdr(ws, 3, col_l,     "Críticos", f_len)
        hdr(ws, 3, col_l + 1, "Bajos",    f_len)

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 22

    # ── Datos ────────────────────────────────────────────────────────────
    estatus_col_idx = ALL_COLS.index("Estatus Centro Escolar") + 1

    for ri, row in enumerate(rows, 4):
        for ci, c in enumerate(ALL_COLS, 1):
            val = row.get(c, "")
            if ci == estatus_col_idx:
                status_dat(ws, ri, ci, val)
            else:
                dat(ws, ri, ci, val)
        ws.row_dimensions[ri].height = 75

    # ── Anchos ───────────────────────────────────────────────────────────
    info_widths = [12, 52, 22, 20, 20, 8, 8]
    for i, w in enumerate(info_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(mat_start, len_end + 1):
        ws.column_dimensions[get_column_letter(i)].width = 55

    ws.freeze_panes = "A4"
    print(f"  ✔  '{sname}' → {len(rows)} filas  ({len(meses)} meses)")


# ══════════════════════════════════════════════════════════════════════════════
#  HOJA 2: Estudiantes_SIGES_LXP
# ══════════════════════════════════════════════════════════════════════════════
def write_estudiantes(wb, rows):
    sname = "Estudiantes_SIGES_LXP"
    for old in ("Estudiantes_SIGES", "Estudiantes_SIGES_LXP"):
        if old in wb.sheetnames:
            del wb[old]
    ws = wb.create_sheet(sname)

    COLS = [
        "Código", "Centro Escolar", "Tipo de Centro", "NIE",
        "Primer Nombre", "Segundo Nombre", "Primer Apellido", "Segundo Apellido",
        "Grado", "Sección",
        "Código Sección LXP", "Código Sección SIGES",
    ]

    for ci, c in enumerate(COLS, 1):
        hdr(ws, 1, ci, c, FILL_DARK)
    ws.row_dimensions[1].height = 32

    for ri, row in enumerate(rows, 2):
        for ci, c in enumerate(COLS, 1):
            align = "left" if c in ("Centro Escolar", "Primer Nombre",
                                    "Segundo Nombre", "Primer Apellido",
                                    "Segundo Apellido") else "center"
            dat(ws, ri, ci, row.get(c, ""), align)

    widths = [12, 55, 20, 14, 20, 20, 22, 22, 8, 8, 20, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    print(f"  ✔  '{sname}' → {len(rows)} filas")


# ══════════════════════════════════════════════════════════════════════════════
#  HOJA 3: Diccionario_Sección_SIGES_LXP
# ══════════════════════════════════════════════════════════════════════════════
def write_diccionario(wb, rows):
    sname = "Diccionario_Sección_SIGES_LXP"
    if sname in wb.sheetnames:
        del wb[sname]
    ws = wb.create_sheet(sname)

    COLS = ["Código", "Centro Escolar", "Tipo de Centro",
            "Grado", "Sección", "Código Sección LXP", "Código Sección SIGES"]

    for ci, c in enumerate(COLS, 1):
        hdr(ws, 1, ci, c, FILL_DICT)
    ws.row_dimensions[1].height = 32

    rows_sorted = sorted(rows, key=lambda r: (
        r.get("Centro Escolar", ""), r.get("Grado", ""), r.get("Sección", "")
    ))

    ALT_FILL = PatternFill("solid", start_color="EEF2F7")
    for ri, row in enumerate(rows_sorted, 2):
        for ci, c in enumerate(COLS, 1):
            align = "left" if c == "Centro Escolar" else "center"
            cell = dat(ws, ri, ci, row.get(c, ""), align)
            if ri % 2 == 0:
                cell.fill = ALT_FILL

    widths = [12, 55, 20, 8, 10, 20, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    print(f"  ✔  '{sname}' → {len(rows_sorted)} filas")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    sep = "═" * 65
    print(sep)
    print("  PAARS – Indicadores Críticos/Bajos por Sección – 3 Meses")
    print(sep)

    print(f"\n[1/4] Leyendo Excel …\n  {EXCEL_PATH}")
    df_est, df_ce = read_excel(EXCEL_PATH)
    print(f"  Estudiantes: {len(df_est)} filas  |  Centros: {len(df_ce)} filas")

    print("\n[2/4] Escaneando HTMLs por mes …")
    nie_maps = {}
    for mes, root in HTML_ROOTS.items():
        print(f"  {mes}  →  {root}")
        nie_maps[mes] = scan_htmls(root)
        print(f"    NIEs mapeados: {len(nie_maps[mes])}")

    print("\n[3/4] Construyendo datos para las nuevas hojas …")
    rows_ind, rows_est, rows_dict = build_data(df_est, df_ce, nie_maps)
    print(f"  Indicadores_Críticos_SecLXP      : {len(rows_ind)} filas")
    print(f"  Estudiantes_SIGES_LXP            : {len(rows_est)} filas")
    print(f"  Diccionario_Sección_SIGES_LXP    : {len(rows_dict)} filas")

    print(f"\n[4/4] Guardando …\n  {OUTPUT_PATH}")
    pathlib.Path(OUTPUT_PATH).parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(EXCEL_PATH)
    write_indicadores(wb, rows_ind)
    write_estudiantes(wb, rows_est)
    write_diccionario(wb, rows_dict)
    wb.save(OUTPUT_PATH)

    print(f"\n✅  Listo.\n{sep}")


if __name__ == "__main__":
    main()