"""
fill_fundamentos.py
-------------------
Updates B1 and B2 Centros_Escolares files with Fundamentos data:

TASK 1 — Estudiantes sheet:
  For each student row, look up their NIE in Reporte_fundamentos_Final.xlsx
  (Estudiantes sheet, column NIE=col 4) and fill:
    ' MAT Fund.' ← Fundamentos.Mat  (col 12 of reporte)
    'LEN Fund.'  ← Fundamentos.Lec  (col 28 of reporte)
  Colour fills applied:
    'Aprobado'                           → green  FF00B050
    'Aprobado con expectativa de mejora' → yellow FFFFD966
    'No Aprobado'                        → red    FFC00000

TASK 2 — Centros Escolares sheet:
  Remove columns: Estatus_2Grado … Estatus_11Grado, Estatus_Escuela_Fundamentos
  Insert 6 new columns BEFORE MAT_Prom_CML:
    % Aprobado (Mat) Fundamentos
    % Aprobado Condicional (Mat) Fundamentos
    % No Aprobado (Mat) Fundamentos
    % Aprobado (Lec) Fundamentos
    % Aprobado Condicional (Lec) Fundamentos
    % No Aprobado (Lec) Fundamentos
  Values = weighted average (by Matrícula Evaluada) across grades from
  Reporte_fundamentos_Final.xlsx → Escuelas sheet, keyed by Código de
  Infraestructura == Código (col 1 of B files CE sheet).

INPUT / OUTPUT files are in the same directory (files updated in place via
separate output paths to avoid overwriting originals).

Requirements: pip install openpyxl pandas
"""

import re
import os
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =============================================================================
# PATHS — edit these
# =============================================================================
BASE_DIR = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Todos_los_puntajes_a_la_Fecha"
    r"\Incluiding_PM4"
)

B1_FILE  = rf"{BASE_DIR}\Centros_Escolares_B1_Actualizado_M4.xlsx"
B2_FILE  = rf"{BASE_DIR}\Centros_Escolares_B2_Actualizado_M4.xlsx"
RPT_FILE = rf"{BASE_DIR}\Reporte_fundamentos_Final.xlsx"

# Output goes into a new subfolder created automatically
OUT_DIR  = rf"{BASE_DIR}\Actualizados_Con_Fundamentos"
B1_OUT   = rf"{OUT_DIR}\Centros_Escolares_B1_Actualizado_M4_Fund.xlsx"
B2_OUT   = rf"{OUT_DIR}\Centros_Escolares_B2_Actualizado_M4_Fund.xlsx"

# =============================================================================
# STYLES
# =============================================================================
_thin       = Side(style="thin", color="BFBFBF")
BORDER      = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_FILL = PatternFill("solid", fgColor="FF2E75B6")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=9)

FUND_FILL = {
    "Aprobado":                           PatternFill("solid", fgColor="FF00B050"),
    "Aprobado con expectativa de mejora": PatternFill("solid", fgColor="FFFFD966"),
    "No Aprobado":                        PatternFill("solid", fgColor="FFC00000"),
}
FUND_FONT = {
    "Aprobado":                           Font(color="FFFFFFFF", bold=True, size=9),
    "Aprobado con expectativa de mejora": Font(color="FF000000", bold=True, size=9),
    "No Aprobado":                        Font(color="FFFFFFFF", bold=True, size=9),
}


def apply_header(cell, text):
    cell.value     = text
    cell.fill      = HEADER_FILL
    cell.font      = HEADER_FONT
    cell.alignment = CENTER
    cell.border    = BORDER


def write_fund_cell(ws, row, col, value):
    """Write a Fundamentos status cell with colour fill."""
    cell           = ws.cell(row, col, value)
    cell.alignment = CENTER
    cell.border    = BORDER
    if value in FUND_FILL:
        cell.fill = FUND_FILL[value]
        cell.font = FUND_FONT[value]


def write_pct_cell(ws, row, col, value):
    """Write a percentage cell (0.0–1.0) formatted as percentage."""
    cell            = ws.cell(row, col, value)
    cell.alignment  = CENTER
    cell.border     = BORDER
    cell.number_format = "0.0%"


def blank_cell(ws, row, col):
    ws.cell(row, col).border = BORDER


def normalize_code(raw):
    if raw is None:
        return None
    try:
        return str(int(float(str(raw).strip())))
    except (ValueError, TypeError):
        return str(raw).strip()


# =============================================================================
# BUILD LOOKUPS FROM Reporte_fundamentos_Final.xlsx
# =============================================================================
def build_student_lookup(rpt_path):
    """
    Returns dict { NIE_int -> ('Aprobado'|..., 'Aprobado'|...) }
                               Fundamentos.Mat  Fundamentos.Lec
    """
    wb  = openpyxl.load_workbook(rpt_path, read_only=True)
    ws  = wb["Estudiantes"]
    lkp = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        nie      = row[3]    # col 4 = NIE
        fund_mat = row[11]   # col 12 = Fundamentos.Mat
        fund_lec = row[27]   # col 28 = Fundamentos.Lec
        if nie is not None:
            try:
                nie_key = int(nie)
            except (ValueError, TypeError):
                nie_key = str(nie).strip()
            lkp[nie_key] = (fund_mat, fund_lec)
    wb.close()
    print(f"  Student lookup: {len(lkp)} entries")
    return lkp


def build_school_lookup(rpt_path):
    """
    Returns dict { codigo_str -> (pct_apr_mat, pct_cond_mat, pct_no_mat,
                                  pct_apr_lec, pct_cond_lec, pct_no_lec) }
    Percentages are weighted averages across grades by Matrícula Evaluada.
    All values are floats 0.0–1.0.
    """
    wb  = openpyxl.load_workbook(rpt_path, read_only=True)
    ws  = wb["Escuelas"]

    # Accumulate: {codigo -> [(n, p5, p6, p7, p8, p9, p10)]}
    accum = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        codigo = normalize_code(row[1])
        n      = row[4]   # Matrícula Evaluada
        if codigo is None or n is None or n == 0:
            continue
        pcts = row[5:11]  # 6 pct columns
        if all(v is None for v in pcts):
            continue
        pcts = tuple(float(v) if v is not None else 0.0 for v in pcts)
        accum.setdefault(codigo, []).append((float(n), *pcts))

    wb.close()

    lkp = {}
    for codigo, rows in accum.items():
        n_total = sum(r[0] for r in rows)
        if n_total == 0:
            continue
        weighted = tuple(
            sum(r[0] * r[i+1] for r in rows) / n_total
            for i in range(6)
        )
        lkp[codigo] = weighted

    print(f"  School lookup: {len(lkp)} entries")
    return lkp


# =============================================================================
# TASK 1 — Fill Fundamentos in Estudiantes sheet
# =============================================================================
CE_COLS_TO_REMOVE = {
    "Estatus_2Grado", "Estatus_3Grado", "Estatus_4Grado", "Estatus_5Grado",
    "Estatus_6Grado", "Estatus_7Grado", "Estatus_8Grado", "Estatus_9Grado",
    "Estatus_10Grado", "Estatus_11Grado", "Estatus_Escuela_Fundamentos",
}

FUND_PCT_LABELS = [
    "% Aprobado (Mat) Fundamentos",
    "% Aprobado Condicional (Mat) Fundamentos",
    "% No Aprobado (Mat) Fundamentos",
    "% Aprobado (Lec) Fundamentos",
    "% Aprobado Condicional (Lec) Fundamentos",
    "% No Aprobado (Lec) Fundamentos",
]


def fill_estudiantes_fundamentos(ws, student_lkp):
    """
    Finds the ' MAT Fund.' and 'LEN Fund.' columns, then for each student
    row looks up their NIE in student_lkp and fills both cells.
    """
    max_col = ws.max_column

    # Find column indices for NIE, MAT Fund., LEN Fund.
    nie_col = mat_fund_col = len_fund_col = None
    for c in range(1, max_col + 1):
        h = ws.cell(1, c).value
        if h is None:
            continue
        h_str = str(h).strip()
        if h_str == "NIE":
            nie_col = c
        elif h_str == "MAT Fund.":     # strip() removes leading space from ' MAT Fund.'
            mat_fund_col = c
        elif h_str == "LEN Fund.":
            len_fund_col = c

    if nie_col is None:
        print("    WARNING: NIE column not found in Estudiantes")
        return 0, 0
    if mat_fund_col is None or len_fund_col is None:
        print(f"    WARNING: Fund columns not found (MAT={mat_fund_col}, LEN={len_fund_col})")
        return 0, 0

    print(f"    Estudiantes: NIE=col{nie_col}, MAT Fund.=col{mat_fund_col}, LEN Fund.=col{len_fund_col}")

    matched = 0
    unmatched = 0
    for row_idx in range(2, ws.max_row + 1):
        nie_raw = ws.cell(row_idx, nie_col).value
        if nie_raw is None:
            continue
        try:
            nie_key = int(nie_raw)
        except (ValueError, TypeError):
            nie_key = str(nie_raw).strip()

        result = student_lkp.get(nie_key)
        if result:
            fund_mat, fund_lec = result
            write_fund_cell(ws, row_idx, mat_fund_col, fund_mat)
            write_fund_cell(ws, row_idx, len_fund_col, fund_lec)
            matched += 1
        else:
            blank_cell(ws, row_idx, mat_fund_col)
            blank_cell(ws, row_idx, len_fund_col)
            unmatched += 1

    pct = f"{matched/(matched+unmatched)*100:.1f}%" if (matched+unmatched) > 0 else "n/a"
    print(f"    Matched {matched} ({pct}) | Unmatched {unmatched}")
    return matched, unmatched


# =============================================================================
# TASK 2 — Update Centros Escolares sheet
# =============================================================================
def update_centros_escolares(ws, school_lkp):
    """
    1. Remove Estatus_*Grado and Estatus_Escuela_Fundamentos columns.
    2. Insert 6 Fundamentos % columns just before MAT_Prom_CML.
    3. Fill values from school_lkp.
    """
    max_col = ws.max_column

    # ── Step 1: find columns to delete ───────────────────────────────────────
    del_cols = []
    for c in range(1, max_col + 1):
        h = ws.cell(1, c).value
        if h is not None and str(h).strip() in CE_COLS_TO_REMOVE:
            del_cols.append(c)

    print(f"    CE: removing {len(del_cols)} estatus columns: {del_cols}")

    # Delete in reverse so indices stay valid
    for c in reversed(del_cols):
        ws.delete_cols(c)

    # ── Step 2: find MAT_Prom_CML insertion point (after deletion) ───────────
    ins_col = None
    for c in range(1, ws.max_column + 1):
        if ws.cell(1, c).value == "MAT_Prom_CML":
            ins_col = c
            break

    if ins_col is None:
        # Fallback: append at end
        ins_col = ws.max_column + 1
        print("    WARNING: MAT_Prom_CML not found — appending at end")
    else:
        print(f"    CE: inserting 6 Fundamentos % cols before col {ins_col} (MAT_Prom_CML)")
        ws.insert_cols(ins_col, 6)

    # ── Step 3: write headers for new columns ────────────────────────────────
    for offset, label in enumerate(FUND_PCT_LABELS):
        col = ins_col + offset
        apply_header(ws.cell(1, col), label)
        ws.column_dimensions[get_column_letter(col)].width = 22

    # ── Step 4: find Código column (col 1 in both B files) ───────────────────
    codigo_col = 1  # always col 1

    # ── Step 5: fill data rows ────────────────────────────────────────────────
    matched = unmatched = 0
    for row_idx in range(2, ws.max_row + 1):
        codigo = normalize_code(ws.cell(row_idx, codigo_col).value)
        if codigo is None:
            continue
        result = school_lkp.get(codigo)
        if result:
            for offset, pct_val in enumerate(result):
                write_pct_cell(ws, row_idx, ins_col + offset, round(pct_val, 6))
            matched += 1
        else:
            for offset in range(6):
                blank_cell(ws, row_idx, ins_col + offset)
            unmatched += 1

    pct_str = f"{matched/(matched+unmatched)*100:.1f}%" if (matched+unmatched) > 0 else "n/a"
    print(f"    CE: matched {matched} schools ({pct_str}) | unmatched {unmatched}")


# =============================================================================
# MAIN
# =============================================================================
def process_file(src_path, dst_path, student_lkp, school_lkp, label):
    print(f"\n  Processing {label}...")
    wb     = openpyxl.load_workbook(src_path)
    ws_est = wb["Estudiantes"]
    ws_ce  = wb["Centros Escolares"]

    print("  -- Estudiantes --")
    fill_estudiantes_fundamentos(ws_est, student_lkp)

    print("  -- Centros Escolares --")
    update_centros_escolares(ws_ce, school_lkp)

    wb.save(dst_path)
    print(f"  Saved → {dst_path}")


if __name__ == "__main__":
    # Create output folder if it doesn't exist
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"Output folder: {OUT_DIR}")

    print("Building lookups from Reporte_fundamentos_Final.xlsx...")
    student_lkp = build_student_lookup(RPT_FILE)
    school_lkp  = build_school_lookup(RPT_FILE)

    process_file(B1_FILE, B1_OUT, student_lkp, school_lkp, "B1")
    process_file(B2_FILE, B2_OUT, student_lkp, school_lkp, "B2")

    print("\nDone.")