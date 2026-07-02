"""
agregar_mes4.py
---------------
Adds Prueba de Progreso Mes 4 (Junio) results to Centros_Escolares B1 and B2.

OUTPUT STRUCTURE (grounded on exact inspection of reference files):

  Estudiantes sheet
  -----------------
  Single flat header row. ALL header cells: blue fill FF2E75B6, white bold. Row height=36.
  Fixed column names:
    1-12  : Código, Centro Escolar, Tipo de Centro, NIE, Primer Nombre,
             Segundo Nombre, Primer Apellido, Segundo Apellido, Grado,
             Estatus Centro Escolar, Estatus Grado, Código Sección LXP
    13-14 : MAT_Puntaje_CML / MAT_Nivel_CML
    15-16 : LEN_Puntaje_CML / LEN_Nivel_CML          ← LEN (not LEC)
    17-18 : " MAT Fund." / "LEN Fund."
    19-...: existing Mat/Len month score cols (kept verbatim from source)
    last4 : MAT_Puntaje_M4, MAT_Nivel_M4, Len  Puntaje M4 (Mayo), Len  Nivel M4 (Mayo)
             appended at end — MAT pair then LEN pair
  Nivel data: write_nivel styles all 5 values on new M4 cells only.
              Pre-existing cells keep their original styling (no restyle pass).

  Centros Escolares sheet
  -----------------------
  Single flat header row. ALL header cells: blue fill FF2E75B6, white bold. Row height=36.
  Fixed column names:
    Metadata : Código, Centro_Escolar, Tipo_de_Centro, Estatus_General,
               Mat_veces_univ_N, Len_veces_univ_N, [N_Aplicaciones_Participadas B1 only],
               Estatus_2Grado … Estatus_11Grado, Estatus_Escuela_Fundamentos
    CML      : MAT_Prom_CML, MAT_Nivel_CML, LEN_Prom_CML, LEN_Nivel_CML
               (NO Fundamentos score columns in CE)
    MAT block: MAT_Promedio_M1…M4, MAT_Nivel_M1…M4
    LEN block: LEN_Promedio_M1…M4, LEN_Nivel_M1…M4
  Nivel data: ALL cells with Crítico/Bajo/Medio/Bueno/Excelente get colour fill.

Requirements:  pip install openpyxl pandas
"""

import re
import time
import tempfile
import os
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =============================================================================
# PATHS — edit these
# =============================================================================
B_FILES_DIR = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Todos_los_puntajes_a_la_Fecha"
    r"\Incluiding_PM4\B_FilesToModify"
)
RESULTS_DIR = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Todos_los_puntajes_a_la_Fecha"
    r"\Incluiding_PM4\ExcelFiles_M4"
)
OUTPUT_DIR = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Todos_los_puntajes_a_la_Fecha"
    r"\Incluiding_PM4"
)

B1_FILE  = rf"{B_FILES_DIR}\Centros_Escolares_B1_Actualizado.xlsx"
B2_FILE  = rf"{B_FILES_DIR}\Centros_Escolares_B2_Actualizado.xlsx"
MAT_FILE = rf"{RESULTS_DIR}\MAT-resultados.xlsx"
LEC_FILE = rf"{RESULTS_DIR}\LEC-resultados.xlsx"

B1_OUT = rf"{OUTPUT_DIR}\Centros_Escolares_B1_Actualizado_M4.xlsx"
B2_OUT = rf"{OUTPUT_DIR}\Centros_Escolares_B2_Actualizado_M4.xlsx"

GRADE_SHEETS = ['2do', '3er', '4to', '5to', '6to',
                '7mo', '8vo', '9no', 'Bach-1er', 'Bach-2do']

# =============================================================================
# SCHOOLS TO REMOVE
# =============================================================================
SCHOOLS_TO_REMOVE = {
    "10065", "10184", "10185", "10188", "10262", "10280", "10291", "10308",
    "10465", "10494", "10513", "10646", "11005", "11058", "11114", "11227",
    "11257", "11280", "11298", "11329", "11350", "11358", "11477", "11478",
    "11503", "11505", "11506", "11546", "11583", "11633", "11652", "11654",
    "11673", "11694", "11719", "11727", "11734", "11742", "11808", "11848",
    "11861", "11869", "11971", "11998", "12009", "12055", "12072", "12076",
    "12096", "12097", "12109", "12116", "12122", "12123", "60122", "68143",
    "70037", "70039", "70047", "72019", "72058", "74030", "74060", "74071",
    "74084", "74091", "78089", "86284", "86351",
}

# =============================================================================
# STYLES
# =============================================================================
_thin        = Side(style="thin", color="BFBFBF")
BORDER       = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_FILL  = PatternFill("solid", fgColor="FF2E75B6")   # blue — used for ALL headers
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=9)

NIVEL_FILL = {
    "Crítico":   PatternFill("solid", fgColor="FFC00000"),
    "Bajo":      PatternFill("solid", fgColor="FFFF8C2E"),
    "Medio":     PatternFill("solid", fgColor="FFFFD966"),
    "Bueno":     PatternFill("solid", fgColor="FF92D050"),
    "Excelente": PatternFill("solid", fgColor="FF00B050"),
}
NIVEL_FONT = {
    "Crítico":   Font(color="FFFFFF", bold=True, size=9),
    "Bajo":      Font(color="FFFFFF", bold=True, size=9),
    "Medio":     Font(color="000000", bold=True, size=9),
    "Bueno":     Font(color="000000", bold=True, size=9),
    "Excelente": Font(color="FFFFFF", bold=True, size=9),
}
NIVEL_VALUES = set(NIVEL_FILL.keys())


def apply_header(cell, text):
    """All header cells: blue fill, white bold font, centered, border."""
    cell.value     = text
    cell.fill      = HEADER_FILL
    cell.font      = HEADER_FONT
    cell.alignment = CENTER
    cell.border    = BORDER


def write_score(ws, row, col, value):
    c           = ws.cell(row, col, value)
    c.alignment = CENTER
    c.border    = BORDER


def write_nivel(ws, row, col, nivel):
    """Write nivel cell with colour fill (used for new M4 cells)."""
    c           = ws.cell(row, col, nivel)
    c.alignment = CENTER
    c.border    = BORDER
    if nivel in NIVEL_FILL:
        c.fill = NIVEL_FILL[nivel]
        c.font = NIVEL_FONT[nivel]


def blank(ws, row, col):
    ws.cell(row, col).border = BORDER


# =============================================================================
# HELPERS
# =============================================================================
def normalize_codigo(raw):
    if raw is None:
        return None
    try:
        return str(int(float(str(raw).strip())))
    except (ValueError, TypeError):
        return str(raw).strip()


def clasificar(val):
    if pd.isna(val):
        return ""
    v = float(val)
    if   v <= 35: return "Crítico"
    elif v <= 45: return "Bajo"
    elif v <= 55: return "Medio"
    elif v <= 65: return "Bueno"
    else:         return "Excelente"


def build_lookup(path):
    lookup = {}
    for sheet in GRADE_SHEETS:
        df = pd.read_excel(
            path, sheet_name=sheet, dtype=str,
            usecols=['Documento', 'theta.global (escala 0-100)']
        )
        df['Documento'] = (df['Documento'].str.strip()
                                          .str.replace(r'\.0$', '', regex=True))
        df['score'] = pd.to_numeric(df['theta.global (escala 0-100)'], errors='coerce')
        for _, row in df.iterrows():
            if pd.notna(row['score']):
                lookup[row['Documento']] = round(float(row['score']), 2)
    return lookup


def month_tag(text):
    if not text:
        return ""
    m = re.search(r'\bM(\d)\b', str(text), re.IGNORECASE)
    if m:
        return f"M{m.group(1)}"
    tl = str(text).lower()
    if "febrero" in tl or "feb" in tl:
        return "CML"
    if "fundamento" in tl or "fund." in tl:
        return "Fundamentos"
    return ""


def subject_prefix(text):
    if not text:
        return None
    tl = str(text).lower()
    if "matem" in tl or tl.strip().endswith(" mat") or re.search(r'\bmat\b', tl):
        return "MAT"
    if ("lenguaje" in tl or "lectura" in tl or "lengua" in tl
            or tl.strip().endswith(" len") or re.search(r'\b(len|lec)\b', tl)):
        return "LEN"
    return None


# =============================================================================
# SCHOOL FILTER  (read_only → filter → write_only temp → reload)
# =============================================================================
def filter_and_rebuild(src_path, sheet_data_start):
    """
    sheet_data_start: dict {sheet_name: data_start_row}
    For Estudiantes, if data_start_row=2 is passed but the source has a
    2-row merged header (group row + sub-label row), data_start is
    auto-detected as 3.
    """
    wb_ro = openpyxl.load_workbook(src_path, read_only=True)
    sheets_rows    = {}
    removed_counts = {}

    for sname in wb_ro.sheetnames:
        ws_ro    = wb_ro[sname]
        all_rows = [list(row) for row in ws_ro.iter_rows(values_only=True)]
        dstart   = sheet_data_start.get(sname)
        if dstart is not None:
            # Auto-detect 2-row header for Estudiantes: if row 1 contains
            # 'Matemática' or 'Lenguaje' as a standalone group label,
            # the actual data starts at row 3.
            if sname == "Estudiantes" and dstart == 2 and all_rows:
                row1_vals = {str(v).strip() for v in all_rows[0] if v is not None}
                if "Matemática" in row1_vals or "Lenguaje" in row1_vals:
                    dstart = 3
                    print(f"    filter: auto-detected 2-row header in Estudiantes → data_start=3")
            header = all_rows[:dstart - 1]
            data   = all_rows[dstart - 1:]
            kept, removed = [], 0
            for row in data:
                if normalize_codigo(row[0]) in SCHOOLS_TO_REMOVE:
                    removed += 1
                else:
                    kept.append(row)
            sheets_rows[sname]    = header + kept
            removed_counts[sname] = removed
        else:
            sheets_rows[sname]    = all_rows
            removed_counts[sname] = 0

    wb_ro.close()

    tmp = tempfile.mktemp(suffix=".xlsx")
    wb_wo = Workbook(write_only=True)
    for sname in wb_ro.sheetnames:
        ws_wo = wb_wo.create_sheet(sname)
        for row in sheets_rows[sname]:
            ws_wo.append(list(row))
    wb_wo.save(tmp)

    return tmp, removed_counts


# =============================================================================
# RESTYLE NIVEL — CE only (all 5 nivel values get colour fill)
# =============================================================================
def restyle_nivel_ce(ws, data_start_row=2):
    """Full scan — fast because CE is only ~190 rows."""
    count = 0
    for row in ws.iter_rows(min_row=data_start_row):
        for cell in row:
            if cell.value in NIVEL_VALUES:
                cell.fill      = NIVEL_FILL[cell.value]
                cell.font      = NIVEL_FONT[cell.value]
                cell.alignment = CENTER
                cell.border    = BORDER
                count += 1
    return count


def restyle_nivel_estudiantes(ws, data_start_row=2):
    """
    Targeted scan for Estudiantes (up to 75k rows).
    Only scans columns whose header contains 'Nivel', 'Fund', or is a known
    nivel column label — avoids iterating all 34 cols × 75k rows.
    """
    # Identify which columns are nivel columns from their header label
    nivel_cols = []
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h is None:
            continue
        h_lower = str(h).lower()
        if ("nivel" in h_lower or "fund" in h_lower):
            nivel_cols.append(c)

    count = 0
    for col in nivel_cols:
        for (cell,) in ws.iter_rows(min_row=data_start_row,
                                     min_col=col, max_col=col):
            if cell.value in NIVEL_VALUES:
                cell.fill      = NIVEL_FILL[cell.value]
                cell.font      = NIVEL_FONT[cell.value]
                cell.alignment = CENTER
                cell.border    = BORDER
                count += 1
    return count, len(nivel_cols)


# =============================================================================
# ESTUDIANTES — rename header + append M4 cols
#
# Exact label mapping (from reference files):
#   Original 'Puntaje M1 (Marzo) Mat' → 'Mat Puntaje M1 (Marzo)'
#   Original 'Nivel M1 (Marzo) Mat'   → 'Mat  Nivel M1 (Marzo)'
#   Original 'Puntaje M1 (Marzo) Len' → 'Len Puntaje M1 (Marzo)'
#   Original 'Nivel M1 (Marzo) Len'   → 'Len  Nivel M1 (Marzo)'
#   CML MAT pair                       → 'MAT_Puntaje_CML', 'MAT_Nivel_CML'
#   CML LEN pair                       → 'LEN_Puntaje_CML', 'LEN_Nivel_CML'
#   Fundamentos MAT single col         → ' MAT Fund.'   (leading space)
#   Fundamentos LEN single col         → 'LEN Fund.'
#   Metadata cols (1-12)               → kept verbatim from source
#   M4 new headers:
#     MAT puntaje  → 'MAT_Puntaje_M4'
#     MAT nivel    → 'MAT_Nivel_M4'
#     LEN puntaje  → 'Len  Puntaje M4 (Mayo)'
#     LEN nivel    → 'Len  Nivel M4 (Mayo)'
# =============================================================================
def _is_two_row_source(ws):
    """
    Returns True if the Estudiantes sheet has a 2-row merged header
    (row 1 has group labels like 'Matemática', 'Lenguaje').
    Returns False if row 1 is already a flat single-row header.
    """
    row1 = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    for v in row1:
        if v is not None and str(v).strip() in ("Matemática", "Lenguaje"):
            return True
    return False


def _categorise_col(label_r1, label_r2, first_data_val_fn, c_idx,
                    mat_cml_cols, len_cml_cols, mat_fund_cols, len_fund_cols,
                    mat_month_cols, len_month_cols, metadata_cols, cml_seen):
    """
    Given a (row1, row2) pair for one column, decide which bucket it belongs to
    and append (output_label, c_idx) to the right list.
    Works for both source formats:
      - Single-row: label_r1 = full flat label, label_r2 = None
      - Two-row:    label_r1 = group name ('Matemática'), label_r2 = sub-label ('Puntaje M3')
    """
    # For single-row source, the meaningful text is entirely in label_r1.
    # For two-row source, label_r1 is the group, label_r2 is the sub-label.
    # We combine them: text = label_r2 if it exists, else label_r1.
    grp = str(label_r1).strip() if label_r1 is not None else ""
    sub = str(label_r2).strip() if label_r2 is not None else ""
    tl_grp = grp.lower()
    tl_sub = sub.lower()
    tl_full = (sub or grp).lower()

    # Determine subject from the group header (or full text for single-row)
    raw_subj = subject_prefix(grp) or subject_prefix(sub) or subject_prefix(grp + " " + sub)

    # ── CML / Febrero ──────────────────────────────────────────────────────
    cml_in_full = ("conociendo" in tl_full or "feb" in tl_full)
    if cml_in_full:
        subj = raw_subj if raw_subj else "MAT"
        first = first_data_val_fn(c_idx)
        if isinstance(first, (int, float)):
            kind = "Puntaje"
        elif isinstance(first, str):
            kind = "Nivel"
        else:
            cnt  = cml_seen.get(subj, 0) + 1
            cml_seen[subj] = cnt
            kind = "Puntaje" if cnt % 2 == 1 else "Nivel"
        label = f"{subj}_{kind}_CML"
        if subj == "MAT":
            mat_cml_cols.append((label, c_idx))
        else:
            len_cml_cols.append((label, c_idx))
        return

    # ── Fundamentos ────────────────────────────────────────────────────────
    fund_in_full = ("fundamento" in tl_full or "fund" in tl_full)
    if fund_in_full and raw_subj:
        if raw_subj == "MAT":
            # Keep only the first MAT Fund col (Puntaje/score); skip Nivel duplicate
            if not mat_fund_cols:
                mat_fund_cols.append((" MAT Fund.", c_idx))
        else:
            # Keep only the first LEN Fund col
            if not len_fund_cols:
                len_fund_cols.append(("LEN Fund.", c_idx))
        return

    # ── Prueba de Progreso month score cols ────────────────────────────────
    # Single-row: text = 'Puntaje M1 (Marzo) Mat'
    # Two-row:    grp='Matemática', sub='Puntaje M3 (Mayo)' or 'Puntaje M4'
    score_text = sub if sub else grp   # prefer row2 sub-label for two-row format

    m_flat = re.match(r'^(Puntaje|Nivel)\s+(M\d)\s*\(([^)]+)\)\s*(Mat|Len)',
                      grp, re.IGNORECASE)   # single-row full label in row1
    m_sub  = re.match(r'^(Puntaje|Nivel)\s+(M\d)(?:\s*\(([^)]+)\))?',
                      score_text, re.IGNORECASE)  # row2 sub-label

    if m_flat:
        kind    = m_flat.group(1).capitalize()
        mtag    = m_flat.group(2).upper()
        month_n = m_flat.group(3).strip()
        subj    = m_flat.group(4).capitalize()
        sp = " " if (kind == "Puntaje" and mtag == "M1") else "  "
        label = f"{subj}{sp}{kind} {mtag} ({month_n})"
        d = mat_month_cols if subj == "Mat" else len_month_cols
        d.setdefault(mtag, []).append((label, c_idx))
        return

    if m_sub and raw_subj:
        kind    = m_sub.group(1).capitalize()
        mtag    = m_sub.group(2).upper()
        month_n = m_sub.group(3).strip() if m_sub.group(3) else ""
        subj    = "Mat" if raw_subj == "MAT" else "Len"

        if mtag == "M4":
            # M4 cols → sentinel labels (written by update_estudiantes)
            if raw_subj == "MAT":
                mat_month_cols.setdefault("M4", []).append(("MAT_Puntaje_M4" if kind=="Puntaje" else "MAT_Nivel_M4", None))
            else:
                mn = f"{month_n}" if month_n else "Mayo"
                lbl = f"Len  {kind} M4 ({mn})" if kind=="Puntaje" else f"Len  {kind} M4 ({mn})"
                len_month_cols.setdefault("M4", []).append((lbl, None))
            return

        # M1/M2/M3
        if not month_n:
            # Infer month name from tag
            month_map = {"M1":"Marzo","M2":"Abril","M3":"Mayo"}
            month_n = month_map.get(mtag, mtag)
        sp = " " if (kind == "Puntaje" and mtag == "M1") else "  "
        label = f"{subj}{sp}{kind} {mtag} ({month_n})"
        d = mat_month_cols if raw_subj == "MAT" else len_month_cols
        d.setdefault(mtag, []).append((label, c_idx))
        return

    # ── Metadata ───────────────────────────────────────────────────────────
    meta_label = grp if grp else sub
    metadata_cols.append((meta_label if meta_label else f"Col_{c_idx}", c_idx))


def rename_estudiantes_header(ws):
    """
    Handles BOTH source formats:

    Format A — single-row flat (e.g. uploaded test files):
      Row 1: 'Puntaje M1 (Marzo) Mat', 'Nivel M1 (Marzo) Mat', ...
      Data starts row 2.

    Format B — 2-row merged (production files):
      Row 1: ..., 'Matemática', None, 'Lenguaje', None, ...
      Row 2: ..., 'Puntaje M3 (Mayo)', 'Nivel M3 (Mayo)', ...
      Data starts row 3.

    Both formats produce identical output matching the reference files.
    """
    # Unmerge everything first
    merge_tl = {}
    for mr in ws.merged_cells.ranges:
        v = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merge_tl[(r, c)] = v
    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))

    def eff(row, col):
        return merge_tl.get((row, col), ws.cell(row, col).value)

    two_row = _is_two_row_source(ws)
    max_col = ws.max_column

    if two_row:
        row1 = [eff(1, c) for c in range(1, max_col + 1)]
        row2 = [eff(2, c) for c in range(1, max_col + 1)]
        # Forward-fill row1 group names
        cur = None
        r1f = []
        for v in row1:
            if v is not None and str(v).strip():
                cur = str(v).strip()
            r1f.append(cur)
        label_r1_list = r1f
        label_r2_list = [str(v).strip() if v is not None else "" for v in row2]
        data_start = 3
        print(f"    Estudiantes: detected 2-row merged header (data starts row 3)")
    else:
        label_r1_list = [str(ws.cell(1, c).value).strip() if ws.cell(1, c).value is not None else ""
                         for c in range(1, max_col + 1)]
        label_r2_list = [""] * max_col
        data_start = 2
        print(f"    Estudiantes: detected single-row flat header (data starts row 2)")

    def first_data_val(col_idx):
        for r in range(data_start, min(ws.max_row + 1, data_start + 1000)):
            v = ws.cell(r, col_idx).value
            if v is not None:
                return v
        return None

    # Categorise columns
    metadata_cols  = []
    mat_cml_cols   = []
    len_cml_cols   = []
    mat_fund_cols  = []
    len_fund_cols  = []
    mat_month_cols = {}
    len_month_cols = {}
    cml_seen       = {}

    for c_idx in range(1, max_col + 1):
        _categorise_col(
            label_r1_list[c_idx - 1],
            label_r2_list[c_idx - 1],
            first_data_val, c_idx,
            mat_cml_cols, len_cml_cols,
            mat_fund_cols, len_fund_cols,
            mat_month_cols, len_month_cols,
            metadata_cols, cml_seen
        )

    # Build ordered column list:
    # metadata → CML MAT → CML LEN → Fund MAT → Fund LEN →
    # MAT months M1/M2/M3 → MAT_M4 sentinel →
    # LEN months M1/M2/M3 → LEN_M4 sentinel
    month_order = ["M1", "M2", "M3"]

    # M4 sentinels: always use fixed reference labels
    mat_m4_sent = [("MAT_Puntaje_M4", None), ("MAT_Nivel_M4", None)]
    len_m4_sent = [("Len  Puntaje M4 (Mayo)", None), ("Len  Nivel M4 (Mayo)", None)]

    # If M4 was detected in the source (two-row), use source col indices
    if mat_month_cols.get("M4"):
        mat_m4_sent = mat_month_cols["M4"]
    if len_month_cols.get("M4"):
        len_m4_sent = len_month_cols["M4"]

    ordered = (
        metadata_cols
        + mat_cml_cols
        + len_cml_cols
        + mat_fund_cols
        + len_fund_cols
        + [c for tag in month_order for c in mat_month_cols.get(tag, [])]
        + mat_m4_sent
        + [c for tag in month_order for c in len_month_cols.get(tag, [])]
        + len_m4_sent
    )

    # Read all data rows
    data_rows = []
    for r in range(data_start, ws.max_row + 1):
        data_rows.append([ws.cell(r, c).value for c in range(1, max_col + 1)])

    # Clear sheet
    ws._cells.clear()
    ws.row_dimensions.clear()

    # Write reordered flat header — ALL cols get blue style
    for new_c, (label, _) in enumerate(ordered, start=1):
        apply_header(ws.cell(1, new_c), label)
        ws.column_dimensions[get_column_letter(new_c)].width = 16

    ws.row_dimensions[1].height = 36

    # Write reordered data rows (skip sentinels with orig_c=None)
    for r_offset, data_row in enumerate(data_rows, start=2):
        for new_c, (_, orig_c) in enumerate(ordered, start=1):
            if orig_c is None:
                continue
            val = data_row[orig_c - 1]
            if val is not None:
                ws.cell(r_offset, new_c).value = val

    label_to_col = {lbl: i for i, (lbl, _) in enumerate(ordered, start=1)}
    print(f"    Estudiantes header reordered+renamed ({len(ordered)} columns)")
    return label_to_col, data_start

    max_col = ws.max_column
    raw     = [ws.cell(1, c).value for c in range(1, max_col + 1)]

    def first_data_val(col_idx):
        for r in range(2, min(ws.max_row + 1, 1002)):
            v = ws.cell(r, col_idx).value
            if v is not None:
                return v
        return None

    # Categorise each original column
    metadata_cols = []   # (new_label, orig_col)
    mat_cml_cols  = []
    len_cml_cols  = []
    mat_fund_cols = []
    len_fund_cols = []
    mat_month_cols = {}  # tag -> [(new_label, orig_col)]
    len_month_cols = {}
    cml_seen = {}

    for c_idx, val in enumerate(raw, start=1):
        text = str(val).strip() if val is not None else ""
        tl   = text.lower()

        # Prueba de Progreso score cols: 'Puntaje M1 (Marzo) Mat' / 'Nivel M1 (Marzo) Len'
        m = re.match(r'^(Puntaje|Nivel)\s+(M\d)\s*\(([^)]+)\)\s*(Mat|Len)',
                     text, re.IGNORECASE)
        if m:
            kind    = m.group(1).capitalize()
            mtag    = m.group(2).upper()
            month_n = m.group(3).strip()
            subj    = m.group(4).capitalize()   # Mat | Len
            # Reference spacing: M1 uses single space, M2+ uses double space before kind
            # 'Mat Puntaje M1 (Marzo)'  vs  'Mat  Puntaje M2 (Abril)'
            # 'Mat  Nivel M1 (Marzo)'   vs  'Mat  Nivel M2 (Abril)'   ← Nivel always double
            sp = " " if (kind == "Puntaje" and mtag == "M1") else "  "
            label = f"{subj}{sp}{kind} {mtag} ({month_n})"
            d = mat_month_cols if subj == "Mat" else len_month_cols
            d.setdefault(mtag, []).append((label, c_idx))
            continue

        # CML cols
        if "conociendo" in tl or (
                "feb" in tl and ("matem" in tl or "lenguaje" in tl or "lengua" in tl)):
            raw_subj = subject_prefix(text)
            subj     = raw_subj if raw_subj else "MAT"
            first    = first_data_val(c_idx)
            if isinstance(first, (int, float)):
                kind = "Puntaje"
            elif isinstance(first, str):
                kind = "Nivel"
            else:
                cnt  = cml_seen.get(subj, 0) + 1
                cml_seen[subj] = cnt
                kind = "Puntaje" if cnt % 2 == 1 else "Nivel"
            label = f"{subj}_{kind}_CML"
            if subj == "MAT":
                mat_cml_cols.append((label, c_idx))
            else:
                len_cml_cols.append((label, c_idx))
            continue

        # Fundamentos cols
        if "fundamento" in tl:
            raw_subj = subject_prefix(text)
            subj     = raw_subj if raw_subj else "MAT"
            if subj == "MAT":
                mat_fund_cols.append((" MAT Fund.", c_idx))
            else:
                len_fund_cols.append(("LEN Fund.", c_idx))
            continue

        # Metadata — kept verbatim
        metadata_cols.append((text if text else f"Col_{c_idx}", c_idx))

    # M4 cols appended after MAT block then after LEN block (not at absolute end)
    # Reference order: MAT_months → MAT_M4 → LEN_months → LEN_M4
    # We use sentinel tuples for M4; update_estudiantes fills them in
    MAT_M4_SENTINEL = "MAT_Puntaje_M4"
    MAT_N4_SENTINEL = "MAT_Nivel_M4"
    LEN_M4_SENTINEL = "Len  Puntaje M4 (Mayo)"
    LEN_N4_SENTINEL = "Len  Nivel M4 (Mayo)"

    month_order = ["M1", "M2", "M3"]
    ordered = (
        metadata_cols
        + mat_cml_cols
        + len_cml_cols
        + mat_fund_cols
        + len_fund_cols
        + [c for tag in month_order for c in mat_month_cols.get(tag, [])]
        + [(MAT_M4_SENTINEL, None), (MAT_N4_SENTINEL, None)]   # M4 MAT placeholder
        + [c for tag in month_order for c in len_month_cols.get(tag, [])]
        + [(LEN_M4_SENTINEL, None), (LEN_N4_SENTINEL, None)]   # M4 LEN placeholder
    )

    # Read all data rows
    data_rows = []
    for r in range(2, ws.max_row + 1):
        data_rows.append([ws.cell(r, c).value for c in range(1, max_col + 1)])

    # Clear sheet
    ws._cells.clear()
    ws.row_dimensions.clear()

    # Write reordered flat header — ALL cols get blue style
    for new_c, (label, _) in enumerate(ordered, start=1):
        apply_header(ws.cell(1, new_c), label)
        ws.column_dimensions[get_column_letter(new_c)].width = 16

    ws.row_dimensions[1].height = 36

    # Write reordered data rows (skip M4 sentinel placeholders — orig_c is None)
    for r_offset, data_row in enumerate(data_rows, start=2):
        for new_c, (_, orig_c) in enumerate(ordered, start=1):
            if orig_c is None:
                continue    # M4 sentinel — filled later by update_estudiantes
            val = data_row[orig_c - 1]
            if val is not None:
                ws.cell(r_offset, new_c).value = val

    label_to_col = {lbl: i for i, (lbl, _) in enumerate(ordered, start=1)}
    print(f"    Estudiantes header reordered+renamed ({len(ordered)} columns)")
    return label_to_col


# =============================================================================
# CENTROS ESCOLARES — flatten 2-row header, drop Mediana + Fundamentos score cols,
#                     reorder, rename with exact reference labels
#
# Exact CE label mapping:
#   CML: MAT_Prom_CML, MAT_Nivel_CML, LEN_Prom_CML, LEN_Nivel_CML
#   Month scores: MAT_Promedio_M1…M4, MAT_Nivel_M1…M4,
#                 LEN_Promedio_M1…M4, LEN_Nivel_M1…M4
#   Fundamentos score cols: DROPPED entirely
# =============================================================================
def flatten_centros_header(ws):
    # Snapshot merge top-left values
    merge_tl = {}
    for mr in ws.merged_cells.ranges:
        v = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merge_tl[(r, c)] = v

    def eff(row, col):
        return merge_tl.get((row, col), ws.cell(row, col).value)

    max_col = ws.max_column
    row1    = [eff(1, c) for c in range(1, max_col + 1)]
    row2    = [eff(2, c) for c in range(1, max_col + 1)]

    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))

    # Forward-fill row-1 group names
    cur_grp = None
    r1f = []
    for v in row1:
        if v is not None and str(v).strip():
            cur_grp = str(v).strip()
        r1f.append(cur_grp)

    # Categorise each original column
    metadata_cols  = []   # (label, orig_col)
    mat_cml_cols   = []
    len_cml_cols   = []
    mat_month_cols = {}   # tag -> [(label, orig_col)]
    len_month_cols = {}
    # Fundamentos and Mediana cols are DROPPED

    for c_idx in range(1, max_col + 1):
        grp   = r1f[c_idx - 1] or ""
        sub   = str(row2[c_idx - 1]).strip() if row2[c_idx - 1] is not None else ""
        sub_l = sub.lower()
        grp_l = grp.lower()

        # Drop Mediana / Nivel Med.
        if "mediana" in sub_l or "nivel med" in sub_l:
            continue

        # Drop Fundamentos score cols (keep only Estatus_Escuela_Fundamentos metadata)
        if "fundamento" in grp_l and ("promedio" in sub_l or "nivel" in sub_l):
            continue

        raw_subj = subject_prefix(grp)

        if raw_subj is not None:
            tag = month_tag(grp) or month_tag(sub)

            # Drop CML Mediana/Nivel Med already handled above
            if "promedio" in sub_l:
                kind = "Promedio"
            elif "nivel" in sub_l:
                kind = "Nivel"
            else:
                kind = "Promedio"

            subj = raw_subj   # MAT or LEN (subject_prefix now returns LEN)

            if tag == "CML":
                # Reference: MAT_Prom_CML / MAT_Nivel_CML / LEN_Prom_CML / LEN_Nivel_CML
                prom_label = f"{subj}_Prom_CML"
                nivel_label= f"{subj}_Nivel_CML"
                label = prom_label if kind == "Promedio" else nivel_label
                if subj == "MAT":
                    mat_cml_cols.append((label, c_idx))
                else:
                    len_cml_cols.append((label, c_idx))

            elif tag and tag != "Fundamentos":
                # Monthly scores: MAT_Promedio_M1 / MAT_Nivel_M1 / LEN_Promedio_M1 ...
                label = f"{subj}_{kind}_{tag}"
                d = mat_month_cols if subj == "MAT" else len_month_cols
                d.setdefault(tag, []).append((label, c_idx))

        elif "estatus por grado" in grp_l:
            grade_m = re.search(r'(\d+)', sub)
            grade   = grade_m.group(1) if grade_m else sub
            label   = f"Estatus_{grade}Grado" if grade else f"Estatus_{sub}"
            metadata_cols.append((label, c_idx))

        else:
            # Plain metadata
            text  = grp if grp else sub
            label = re.sub(r'[\n\r]+', ' ', text).strip()
            label = re.sub(r'\s+', '_',
                    re.sub(r'[^\w\sáéíóúÁÉÍÓÚüÜñÑ]', '', label)).strip('_')
            label = label if label else f"Col_{c_idx}"
            metadata_cols.append((label, c_idx))

    # Build ordered column list matching reference:
    # metadata → CML MAT → CML LEN → MAT M1..M4 → LEN M1..M4
    month_order = ["M1", "M2", "M3"]   # M4 appended separately
    ordered = (
        metadata_cols
        + mat_cml_cols
        + len_cml_cols
        + [c for tag in month_order for c in mat_month_cols.get(tag, [])]
        + [c for tag in month_order for c in len_month_cols.get(tag, [])]
    )

    # Read all data rows
    data_rows = []
    for r in range(3, ws.max_row + 1):
        data_rows.append([ws.cell(r, c).value for c in range(1, max_col + 1)])

    # Clear sheet
    ws._cells.clear()
    ws.row_dimensions.clear()

    # Write reordered flat header
    for new_c, (label, _) in enumerate(ordered, start=1):
        apply_header(ws.cell(1, new_c), label)
        ws.column_dimensions[get_column_letter(new_c)].width = 16

    ws.row_dimensions[1].height = 36

    # Write reordered data rows
    for r_offset, data_row in enumerate(data_rows, start=2):
        for new_c, (_, orig_c) in enumerate(ordered, start=1):
            val = data_row[orig_c - 1]
            if val is not None:
                ws.cell(r_offset, new_c).value = val

    label_to_col = {lbl: i for i, (lbl, _) in enumerate(ordered, start=1)}
    n_cols = len(ordered)
    print(f"    CE flattened+reordered ({n_cols} columns)")
    return label_to_col, n_cols


# =============================================================================
# SHEET 1 — Estudiantes
# =============================================================================
def update_estudiantes(ws, mat_lookup, lec_lookup):
    label_to_col, data_start = rename_estudiantes_header(ws)

    # Restyle ALL nivel cells in Estudiantes (Crítico/Bajo/Medio/Bueno/Excelente).
    # Only scans columns whose header contains 'Nivel' or 'Fund' — fast even on 75k rows.
    t0 = time.time()
    n_cells, n_cols = restyle_nivel_estudiantes(ws, data_start_row=2)
    print(f"    Estudiantes restyle: {time.time()-t0:.1f}s ({n_cells} cells across {n_cols} nivel cols)")

    nie_col = label_to_col.get("NIE", 4)

    # M4 columns land in the sentinel positions created during reorder.
    # Sentinels: 'MAT_Puntaje_M4', 'MAT_Nivel_M4', 'Len  Puntaje M4 (Mayo)', 'Len  Nivel M4 (Mayo)'
    c_mat_p = label_to_col.get("MAT_Puntaje_M4")
    c_mat_n = label_to_col.get("MAT_Nivel_M4")
    c_len_p = label_to_col.get("Len  Puntaje M4 (Mayo)")
    c_len_n = label_to_col.get("Len  Nivel M4 (Mayo)")

    # Fallback: append at end if sentinels not found
    if c_mat_p is None:
        last = ws.max_column
        c_mat_p, c_mat_n, c_len_p, c_len_n = last+1, last+2, last+3, last+4

    # Apply header style to M4 sentinel cols (they were already written but need style)
    for col, text in [
        (c_mat_p, "MAT_Puntaje_M4"),
        (c_mat_n, "MAT_Nivel_M4"),
        (c_len_p, "Len  Puntaje M4 (Mayo)"),
        (c_len_n, "Len  Nivel M4 (Mayo)"),
    ]:
        apply_header(ws.cell(1, col), text)
        ws.column_dimensions[get_column_letter(col)].width = 16

    matched_mat = matched_lec = total = 0
    for row_idx in range(2, ws.max_row + 1):
        nie_raw = ws.cell(row_idx, nie_col).value
        if nie_raw is None:
            continue
        nie   = str(nie_raw).strip().replace(".0", "")
        total += 1

        s_mat = mat_lookup.get(nie)
        if s_mat is not None:
            matched_mat += 1
            write_score(ws, row_idx, c_mat_p, s_mat)
            write_nivel(ws, row_idx, c_mat_n, clasificar(s_mat))
        else:
            blank(ws, row_idx, c_mat_p)
            blank(ws, row_idx, c_mat_n)

        s_lec = lec_lookup.get(nie)
        if s_lec is not None:
            matched_lec += 1
            write_score(ws, row_idx, c_len_p, s_lec)
            write_nivel(ws, row_idx, c_len_n, clasificar(s_lec))
        else:
            blank(ws, row_idx, c_len_p)
            blank(ws, row_idx, c_len_n)

    pct = lambda n: f"{n/total*100:.1f}%" if total else "n/a"
    print(f"    Estudiantes — {total} rows | "
          f"MAT {matched_mat} ({pct(matched_mat)}) | "
          f"LEN {matched_lec} ({pct(matched_lec)})")

    # Build school metadata for new-school detection
    school_meta = {}
    for row_idx in range(2, ws.max_row + 1):
        codigo = normalize_codigo(ws.cell(row_idx, 1).value)
        if codigo is None or codigo in school_meta:
            continue
        school_meta[codigo] = (ws.cell(row_idx, 2).value,
                               ws.cell(row_idx, 3).value)

    return c_mat_p, c_len_p, school_meta


# =============================================================================
# SHEET 2 — Centros Escolares
# =============================================================================
def update_centros_escolares(ws, ws_est, c_mat_p_est, c_len_p_est, school_meta):
    label_to_col, _ = flatten_centros_header(ws)

    # Restyle ALL nivel cells in CE (Crítico/Bajo/Medio/Bueno/Excelente)
    t0 = time.time()
    n  = restyle_nivel_ce(ws, data_start_row=2)
    print(f"    CE restyle: {time.time()-t0:.3f}s ({n} cells)")

    # Build per-school M4 score lists from Estudiantes
    school_mat, school_len = {}, {}
    for row_idx in range(2, ws_est.max_row + 1):
        codigo = normalize_codigo(ws_est.cell(row_idx, 1).value)
        if codigo is None:
            continue
        for d, col in [(school_mat, c_mat_p_est), (school_len, c_len_p_est)]:
            val = ws_est.cell(row_idx, col).value
            if val is not None:
                try:
                    d.setdefault(codigo, []).append(float(val))
                except (ValueError, TypeError):
                    pass

    def avg(scores):
        return round(pd.Series(scores).mean(), 2) if scores else None

    # Insert M4 cols inside MAT block (after last MAT col) and LEN block (after last LEN col)
    last_mat = None
    last_len = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v and str(v).startswith("MAT_"):
            last_mat = c
        elif v and str(v).startswith("LEN_"):
            last_len = c

    ins_mat = (last_mat or ws.max_column) + 1
    ws.insert_cols(ins_mat, 2)
    c_mat_base = ins_mat

    if last_len is not None and last_len >= ins_mat:
        last_len += 2
    ins_len = (last_len or ws.max_column) + 1
    ws.insert_cols(ins_len, 2)
    c_len_base = ins_len

    # Exact reference label names for M4 CE cols
    apply_header(ws.cell(1, c_mat_base),     "MAT_Promedio_M4")
    apply_header(ws.cell(1, c_mat_base + 1), "MAT_Nivel_M4")
    apply_header(ws.cell(1, c_len_base),     "LEN_Promedio_M4")
    apply_header(ws.cell(1, c_len_base + 1), "LEN_Nivel_M4")
    for col in [c_mat_base, c_mat_base+1, c_len_base, c_len_base+1]:
        ws.column_dimensions[get_column_letter(col)].width = 16

    ws.row_dimensions[1].height = 36

    # Write M4 scores for existing schools
    for row_idx in range(2, ws.max_row + 1):
        codigo = normalize_codigo(ws.cell(row_idx, 1).value)
        if codigo is None:
            continue
        avg_m = avg(school_mat.get(codigo, []))
        if avg_m is not None:
            write_score(ws, row_idx, c_mat_base,     avg_m)
            write_nivel(ws, row_idx, c_mat_base + 1, clasificar(avg_m))
        else:
            blank(ws, row_idx, c_mat_base)
            blank(ws, row_idx, c_mat_base + 1)

        avg_l = avg(school_len.get(codigo, []))
        if avg_l is not None:
            write_score(ws, row_idx, c_len_base,     avg_l)
            write_nivel(ws, row_idx, c_len_base + 1, clasificar(avg_l))
        else:
            blank(ws, row_idx, c_len_base)
            blank(ws, row_idx, c_len_base + 1)

    # Detect and append new schools
    existing_ce = set()
    for r in range(2, ws.max_row + 1):
        c = normalize_codigo(ws.cell(r, 1).value)
        if c:
            existing_ce.add(c)

    all_m4      = set(school_mat.keys()) | set(school_len.keys())
    new_schools = sorted(all_m4 - existing_ce)

    if new_schools:
        print(f"    NEW schools: {len(new_schools)} — appending")
        for codigo in new_schools:
            name, tipo = school_meta.get(codigo, (None, None))
            r = ws.max_row + 1
            ws.cell(r, 1).value = int(codigo)
            ws.cell(r, 1).alignment = CENTER
            ws.cell(r, 1).border    = BORDER
            if name:
                ws.cell(r, 2).value     = name
                ws.cell(r, 2).alignment = CENTER
                ws.cell(r, 2).border    = BORDER
            if tipo:
                ws.cell(r, 3).value     = tipo
                ws.cell(r, 3).alignment = CENTER
                ws.cell(r, 3).border    = BORDER
            for c in range(4, ws.max_column + 1):
                ws.cell(r, c).border = BORDER
            avg_m = avg(school_mat.get(codigo, []))
            if avg_m is not None:
                write_score(ws, r, c_mat_base,     avg_m)
                write_nivel(ws, r, c_mat_base + 1, clasificar(avg_m))
            avg_l = avg(school_len.get(codigo, []))
            if avg_l is not None:
                write_score(ws, r, c_len_base,     avg_l)
                write_nivel(ws, r, c_len_base + 1, clasificar(avg_l))
            print(f"      Added: {codigo} — {name}")
    else:
        print("    No new schools to add.")

    print(f"    CE — MAT {len(school_mat)} | LEN {len(school_len)} schools")


# =============================================================================
# MAIN PIPELINE
# =============================================================================
def process_file(src_path, dst_path, mat_lookup, lec_lookup, label):
    print(f"\n  Processing {label}...")
    t_total = time.time()

    t0 = time.time()
    tmp, removed = filter_and_rebuild(
        src_path, {"Estudiantes": 2, "Centros Escolares": 3}
    )
    print(f"    Filter+write_only: {time.time()-t0:.1f}s | "
          f"Est -{removed.get('Estudiantes',0)} | "
          f"CE -{removed.get('Centros Escolares',0)}")

    t0 = time.time()
    wb     = openpyxl.load_workbook(tmp)
    ws_est = wb["Estudiantes"]
    ws_ce  = wb["Centros Escolares"]
    print(f"    Reload: {time.time()-t0:.1f}s | "
          f"Est {ws_est.max_row-1} rows | CE {ws_ce.max_row-2} rows")

    c_mat_p, c_len_p, school_meta = update_estudiantes(ws_est, mat_lookup, lec_lookup)
    update_centros_escolares(ws_ce, ws_est, c_mat_p, c_len_p, school_meta)

    wb.save(dst_path)
    os.remove(tmp)
    print(f"    Saved → {dst_path}  (total {time.time()-t_total:.0f}s)")


if __name__ == "__main__":
    print("Building MAT lookup...")
    mat_lookup = build_lookup(MAT_FILE)
    print(f"  {len(mat_lookup)} students found")

    print("Building LEC lookup...")
    lec_lookup = build_lookup(LEC_FILE)
    print(f"  {len(lec_lookup)} students found")

    process_file(B1_FILE, B1_OUT, mat_lookup, lec_lookup, "B1")
    process_file(B2_FILE, B2_OUT, mat_lookup, lec_lookup, "B2")

    print("\nDone. Both files written successfully.")