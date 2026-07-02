"""
cobertura_participacion.py
==========================
Cruza la matrícula esperada (resumen_escuelas.txt) con la participación
real (Participacion_Progreso4.xlsx) y genera:
  - cobertura_participacion.txt   → tabla en texto plano
  - cobertura_participacion.xlsx  → libro Excel con una hoja por grupo

Estructura de cada tabla (MAT y LEC lado a lado):
  Grado | Matrícula | MAT Partic. | MAT % | LEC Partic. | LEC % | Prom. %

Una hoja / sección por grupo:
  • Global (B1 + B2 combinados)
  • Solo B1
  • Solo B2
"""

import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── RUTAS ──────────────────────────────────────────────────────────────────────
TXT_MATRICULA = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Matricula Progreso 4\resumen_escuelas.txt"
XLS_PARTICIPA = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Participación en la prueba de Progreso 4\Participacion_Progreso4.xlsx"

TXT_OUT       = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Porcentaje de cobertura\cobertura_participacion.txt"
XLS_OUT       = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Porcentaje de cobertura\cobertura_participacion.xlsx"

# ── MAPEO DE GRADOS ────────────────────────────────────────────────────────────
GRADE_LABEL_TO_SHORT = {
    "Segundo Grado": "2°",
    "Tercer Grado":  "3°",
    "Cuarto Grado":  "4°",
    "Quinto Grado":  "5°",
    "Sexto Grado":   "6°",
    "Séptimo Grado": "7°",
    "Octavo Grado":  "8°",
    "Noveno Grado":  "9°",
    "Primer Año":    "10°",
    "Segundo Año":   "11°",
}
GRADE_ORDER_SHORT = list(GRADE_LABEL_TO_SHORT.values())

GRADE_DISPLAY = {
    "2°":  "2° (Segundo)",
    "3°":  "3° (Tercero)",
    "4°":  "4° (Cuarto)",
    "5°":  "5° (Quinto)",
    "6°":  "6° (Sexto)",
    "7°":  "7° (Séptimo)",
    "8°":  "8° (Octavo)",
    "9°":  "9° (Noveno)",
    "10°": "10° (1° Bach.)",
    "11°": "11° (2° Bach.)",
}

GROUPS = {
    "global": "Global (B1 + B2)",
    "B1":     "Grupo B1",
    "B2":     "Grupo B2",
}


# ══════════════════════════════════════════════════════════════════════════════
# 1. PARSE ENROLLMENT FROM TXT
# ══════════════════════════════════════════════════════════════════════════════

def parse_enrollment(txt_path):
    """
    Returns {"global": {"2°": N, ...}, "B1": {...}, "B2": {...}}
    Parses the section "ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)".
    """
    with open(txt_path, encoding="utf-8") as f:
        text = f.read()

    section_pat = re.compile(
        r"ESTUDIANTES POR GRADO Y GRUPO.*?\n(.*?)(?=\n\u2550{10,}|\Z)",
        re.DOTALL
    )
    m = section_pat.search(text)
    if not m:
        raise ValueError("Sección 'ESTUDIANTES POR GRADO Y GRUPO' no encontrada en el TXT.")

    block      = m.group(1)
    enrollment = {"global": {}, "B1": {}, "B2": {}}

    for label, short in GRADE_LABEL_TO_SHORT.items():
        pat = re.compile(
            rf"\s{re.escape(label)}\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)"
        )
        row = pat.search(block)
        if row:
            b1  = int(row.group(1).replace(",", ""))
            b2  = int(row.group(2).replace(",", ""))
            tot = int(row.group(3).replace(",", ""))
            enrollment["B1"][short]     = b1
            enrollment["B2"][short]     = b2
            enrollment["global"][short] = tot
        else:
            enrollment["B1"][short]     = 0
            enrollment["B2"][short]     = 0
            enrollment["global"][short] = 0

    return enrollment


# ══════════════════════════════════════════════════════════════════════════════
# 2. PARSE PARTICIPATION FROM EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def parse_participation(xls_path):
    """
    Returns {"MAT": {"global": {"2°": N, ...}}, "LEC": {...}}
    Excludes school code 99999 (Virtual) and the TOTAL GENERAL row.
    """
    xl     = pd.ExcelFile(xls_path)
    result = {}

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, dtype=str)
        df.columns = df.columns.str.strip()

        code_col = next(
            (c for c in df.columns
             if "nro" in c.lower()
             or "código" in c.lower()
             or ("centro" in c.lower() and len(c) < 20)),
            df.columns[0]
        )

        df = df[~df[code_col].astype(str).str.upper().str.contains(
            "99999|TOTAL", na=False)]
        df = df.dropna(subset=[code_col])

        totals = {
            g: int(pd.to_numeric(df[g], errors="coerce").fillna(0).sum())
            if g in df.columns else 0
            for g in GRADE_ORDER_SHORT
        }

        key = "LEC" if ("LEC" in sheet.upper() or "LENGUA" in sheet.upper()) else "MAT"
        result[key] = {"global": totals}

    return result


# ══════════════════════════════════════════════════════════════════════════════
# 3. BUILD COMBINED ROWS (one row per grade, both subjects)
# ══════════════════════════════════════════════════════════════════════════════

def build_combined_rows(enrollment, participation, group="global"):
    """
    Returns one dict per grade with fields:
      display, expected, mat_real, mat_pct, lec_real, lec_pct, avg_pct
    Plus a totals row at the end.
    """
    enroll  = enrollment.get(group, {})
    mat_par = participation.get("MAT", {}).get("global", {})
    lec_par = participation.get("LEC", {}).get("global", {})

    rows = []
    tot_exp = tot_mat = tot_lec = 0

    for short in GRADE_ORDER_SHORT:
        exp      = enroll.get(short, 0)
        mat_real = mat_par.get(short, 0)
        lec_real = lec_par.get(short, 0)
        mat_pct  = (mat_real / exp * 100) if exp > 0 else 0.0
        lec_pct  = (lec_real / exp * 100) if exp > 0 else 0.0
        avg_pct  = (mat_pct + lec_pct) / 2

        tot_exp += exp
        tot_mat += mat_real
        tot_lec += lec_real

        rows.append({
            "grade":    short,
            "display":  GRADE_DISPLAY[short],
            "expected": exp,
            "mat_real": mat_real,
            "mat_pct":  mat_pct,
            "lec_real": lec_real,
            "lec_pct":  lec_pct,
            "avg_pct":  avg_pct,
        })

    # Totals row
    tot_mat_pct = (tot_mat / tot_exp * 100) if tot_exp > 0 else 0.0
    tot_lec_pct = (tot_lec / tot_exp * 100) if tot_exp > 0 else 0.0
    tot_avg_pct = (tot_mat_pct + tot_lec_pct) / 2

    rows.append({
        "grade":    "TOTAL",
        "display":  "TOTAL GENERAL",
        "expected": tot_exp,
        "mat_real": tot_mat,
        "mat_pct":  tot_mat_pct,
        "lec_real": tot_lec,
        "lec_pct":  tot_lec_pct,
        "avg_pct":  tot_avg_pct,
    })
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# 4. TXT OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

SEP2    = "═" * 95
W_GRADE = 20
W_NUM   = 13
W_PCT   =  9


def _txt_header():
    return (
        f"  {'Grado':<{W_GRADE}}"
        f" {'Matrícula':>{W_NUM}}"
        f" {'MAT Partic.':>{W_NUM}}"
        f" {'MAT %':>{W_PCT}}"
        f" {'LEC Partic.':>{W_NUM}}"
        f" {'LEC %':>{W_PCT}}"
        f" {'Prom. %':>{W_PCT}}"
    )


def _txt_sep():
    return "  " + "─" * (W_GRADE + W_NUM * 3 + W_PCT * 3 + 14)


def _txt_row(row):
    def fp(p): return f"{p:.1f}%"
    return (
        f"  {row['display']:<{W_GRADE}}"
        f" {row['expected']:>{W_NUM},}"
        f" {row['mat_real']:>{W_NUM},}"
        f" {fp(row['mat_pct']):>{W_PCT}}"
        f" {row['lec_real']:>{W_NUM},}"
        f" {fp(row['lec_pct']):>{W_PCT}}"
        f" {fp(row['avg_pct']):>{W_PCT}}"
    )


def build_txt(enrollment, participation):
    lines = []

    for grp_key, grp_label in GROUPS.items():
        lines.append(f"\n{SEP2}")
        lines.append(f"  COBERTURA DE PARTICIPACIÓN — {grp_label}")
        lines.append(f"  Matemática (MAT) y Lengua (LEC) — Prueba de Progreso 4")
        lines.append(SEP2)
        lines.append(_txt_header())
        lines.append(_txt_sep())
        rows = build_combined_rows(enrollment, participation, grp_key)
        for row in rows[:-1]:
            lines.append(_txt_row(row))
        lines.append(_txt_sep())
        lines.append(_txt_row(rows[-1]))
        lines.append("")

    lines.append(SEP2)
    lines.append("  Nota: % Promedio = (MAT % + LEC %) / 2")
    lines.append("  Los valores >100% en Bachillerato reflejan participantes")
    lines.append("  de B2 que se suman a una matrícula de referencia parcial.")
    lines.append(f"{SEP2}\n")

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# 5. EXCEL OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

DARK_BLUE   = "1F3864"
MID_BLUE    = "2E5FAC"
MAT_BLUE    = "1F6B9E"
LEC_TEAL    = "1A6B5A"
AVG_PURPLE  = "4A3A7A"
LIGHT_BLUE  = "D6E4F0"
LIGHT_GRAY  = "F2F2F2"
WHITE       = "FFFFFF"
GREEN_FILL  = "E2EFDA"
ORANGE_FILL = "FCE4D6"
RED_FILL    = "FFE7E7"

THIN_BORDER = Border(
    left=Side(style="thin",   color="CCCCCC"),
    right=Side(style="thin",  color="CCCCCC"),
    top=Side(style="thin",    color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
MED_BORDER_LEFT = Border(
    left=Side(style="medium", color="8EA9C1"),
    right=Side(style="thin",  color="CCCCCC"),
    top=Side(style="thin",    color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def _pct_color(pct):
    if pct >= 90: return GREEN_FILL
    if pct >= 70: return LIGHT_BLUE
    if pct >= 50: return ORANGE_FILL
    return RED_FILL


def _hdr(cell, value, bg=DARK_BLUE, wrap=True):
    cell.value = value
    cell.font  = Font(bold=True, color=WHITE, size=9, name="Calibri")
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=wrap)
    cell.border = THIN_BORDER


def _cell(cell, value, align="right", bold=False, bg=None,
          num_fmt=None, border=None):
    cell.value = value
    cell.font  = Font(bold=bold, size=9, name="Calibri")
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = border if border else THIN_BORDER
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if num_fmt:
        cell.number_format = num_fmt


def write_group_sheet(wb, sheet_name, rows, group_label):
    """
    Write one sheet with combined MAT + LEC table.

    Column layout (7 columns):
      A  Grado
      B  Matrícula Esperada
      C  MAT Participantes    ← medium left border (visual group separator)
      D  MAT % Cobertura
      E  LEC Participantes    ← medium left border
      F  LEC % Cobertura
      G  % Promedio MAT+LEC   ← medium left border
    """
    ws     = wb.create_sheet(sheet_name[:31])
    N_COLS = 7

    def bdr(col):
        return MED_BORDER_LEFT if col in (3, 5, 7) else THIN_BORDER

    # ── Row 1: main title ─────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N_COLS)
    c = ws.cell(row=1, column=1,
                value=f"Cobertura de Participación — {group_label} — Prueba de Progreso 4")
    c.font      = Font(bold=True, color=WHITE, size=12, name="Calibri")
    c.fill      = PatternFill("solid", fgColor=DARK_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border    = THIN_BORDER
    ws.row_dimensions[1].height = 20

    # ── Row 2: subject-group headers ──────────────────────────────────────────
    # Merged spans: A(1), B(1), C:D(2) MAT, E:F(2) LEC, G(1) Promedio
    merge_spans = [(1, 1), (2, 2), (3, 4), (5, 6), (7, 7)]
    grp_labels  = ["", "", "Matemática (MAT)", "Lengua (LEC)", "Promedio"]
    grp_colors  = [DARK_BLUE, DARK_BLUE, MAT_BLUE, LEC_TEAL, AVG_PURPLE]

    for (start_c, end_c), text, bg in zip(merge_spans, grp_labels, grp_colors):
        if start_c != end_c:
            ws.merge_cells(start_row=2, start_column=start_c,
                           end_row=2, end_column=end_c)
        c = ws.cell(row=2, column=start_c, value=text)
        c.font      = Font(bold=True, color=WHITE, size=10, name="Calibri")
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = THIN_BORDER
        # Fill merged cells that are invisible but need styling
        for col in range(start_c, end_c + 1):
            ws.cell(row=2, column=col).fill = PatternFill("solid", fgColor=bg)
    ws.row_dimensions[2].height = 18

    # ── Row 3: column headers ─────────────────────────────────────────────────
    col_headers = [
        ("Grado",               DARK_BLUE),
        ("Matrícula\nEsperada", DARK_BLUE),
        ("Participantes",       MAT_BLUE),
        ("% Cobertura",         MAT_BLUE),
        ("Participantes",       LEC_TEAL),
        ("% Cobertura",         LEC_TEAL),
        ("% Promedio\nMAT+LEC", AVG_PURPLE),
    ]
    for ci, (text, bg) in enumerate(col_headers, start=1):
        _hdr(ws.cell(row=3, column=ci), text, bg=bg)
    ws.row_dimensions[3].height = 32

    # ── Row 4: color legend ────────────────────────────────────────────────────
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=N_COLS)
    note = ws.cell(row=4, column=1,
                   value="Color % cobertura:  \U0001f7e2 \u226590%   "
                         "\U0001f535 70\u201389%   \U0001f7e0 50\u201369%   "
                         "\U0001f534 <50%")
    note.font      = Font(italic=True, size=8, color="555555", name="Calibri")
    note.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[4].height = 13

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r_i, row in enumerate(rows[:-1], start=5):
        bg = WHITE if r_i % 2 == 0 else LIGHT_GRAY
        _cell(ws.cell(row=r_i, column=1), row["display"],
              align="left", bg=bg)
        _cell(ws.cell(row=r_i, column=2), row["expected"],
              num_fmt="#,##0", bg=bg)
        _cell(ws.cell(row=r_i, column=3), row["mat_real"],
              num_fmt="#,##0", bg=bg, border=bdr(3))
        _cell(ws.cell(row=r_i, column=4), round(row["mat_pct"], 1),
              num_fmt='0.0"%"', bg=_pct_color(row["mat_pct"]))
        _cell(ws.cell(row=r_i, column=5), row["lec_real"],
              num_fmt="#,##0", bg=bg, border=bdr(5))
        _cell(ws.cell(row=r_i, column=6), round(row["lec_pct"], 1),
              num_fmt='0.0"%"', bg=_pct_color(row["lec_pct"]))
        _cell(ws.cell(row=r_i, column=7), round(row["avg_pct"], 1),
              num_fmt='0.0"%"', bg=_pct_color(row["avg_pct"]),
              border=bdr(7))

    # ── Totals row ─────────────────────────────────────────────────────────────
    tot     = rows[-1]
    tot_row = len(rows) + 4
    _cell(ws.cell(row=tot_row, column=1), "TOTAL GENERAL",
          align="left", bold=True, bg=LIGHT_BLUE)
    _cell(ws.cell(row=tot_row, column=2), tot["expected"],
          num_fmt="#,##0", bold=True, bg=LIGHT_BLUE)
    _cell(ws.cell(row=tot_row, column=3), tot["mat_real"],
          num_fmt="#,##0", bold=True, bg=LIGHT_BLUE, border=bdr(3))
    _cell(ws.cell(row=tot_row, column=4), round(tot["mat_pct"], 1),
          num_fmt='0.0"%"', bold=True, bg=_pct_color(tot["mat_pct"]))
    _cell(ws.cell(row=tot_row, column=5), tot["lec_real"],
          num_fmt="#,##0", bold=True, bg=LIGHT_BLUE, border=bdr(5))
    _cell(ws.cell(row=tot_row, column=6), round(tot["lec_pct"], 1),
          num_fmt='0.0"%"', bold=True, bg=_pct_color(tot["lec_pct"]))
    _cell(ws.cell(row=tot_row, column=7), round(tot["avg_pct"], 1),
          num_fmt='0.0"%"', bold=True, bg=_pct_color(tot["avg_pct"]),
          border=bdr(7))

    # ── Column widths ──────────────────────────────────────────────────────────
    for ci, w in enumerate([22, 16, 16, 14, 16, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A5"


# ══════════════════════════════════════════════════════════════════════════════
# 6. MAIN
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("[*] Leyendo matrícula desde TXT...")
    enrollment = parse_enrollment(TXT_MATRICULA)
    print(f"    Matrícula global total: "
          f"{sum(enrollment['global'].values()):,} estudiantes")

    print("[*] Leyendo participación desde Excel...")
    participation = parse_participation(XLS_PARTICIPA)
    for subj, data in participation.items():
        print(f"    {subj}: {sum(data['global'].values()):,} participantes reales")

    # Crear carpeta de salida si no existe
    os.makedirs(os.path.dirname(TXT_OUT), exist_ok=True)

    # ── TXT ───────────────────────────────────────────────────────────────────
    print("\n[*] Generando TXT...")
    txt_content = build_txt(enrollment, participation)
    with open(TXT_OUT, "w", encoding="utf-8") as f:
        f.write(txt_content)
    print(f"    ✔ Guardado: {TXT_OUT}")

    # ── EXCEL ─────────────────────────────────────────────────────────────────
    print("\n[*] Generando Excel...")
    wb = Workbook()
    wb.remove(wb.active)

    for grp_key, grp_label in GROUPS.items():
        rows = build_combined_rows(enrollment, participation, grp_key)
        write_group_sheet(wb, grp_label, rows, grp_label)
        print(f"    Hoja '{grp_label}' creada")

    wb.save(XLS_OUT)
    print(f"    ✔ Guardado: {XLS_OUT}")

    print("\n" + txt_content)