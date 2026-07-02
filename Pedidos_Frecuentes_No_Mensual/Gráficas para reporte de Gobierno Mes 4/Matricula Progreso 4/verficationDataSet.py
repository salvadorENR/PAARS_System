import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

FILE    = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Matricula Progreso 4\Recolección de datos G1-G2_25_05_2026.xlsx"
TXT_OUT = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Matricula Progreso 4\resumen_escuelas.txt"
PDF_OUT = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Matricula Progreso 4\resumen_escuelas.pdf"
XLS_OUT = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Matricula Progreso 4\resumen_escuelas.xlsx"

# ── Load data ──────────────────────────────────────────────────────────────────
df = pd.read_excel(FILE)

SEP  = "─" * 72
SEP2 = "═" * 72

# ── Grade order and short column labels ───────────────────────────────────────
grade_order = [
    "Segundo Grado", "Tercer Grado", "Cuarto Grado", "Quinto Grado",
    "Sexto Grado",   "Séptimo Grado", "Octavo Grado", "Noveno Grado",
    "Primer Año",    "Segundo Año"
]
# Short labels used as column headers in the wide tables
grade_abbrev = {
    "Segundo Grado":  "2°",
    "Tercer Grado":   "3°",
    "Cuarto Grado":   "4°",
    "Quinto Grado":   "5°",
    "Sexto Grado":    "6°",
    "Séptimo Grado":  "7°",
    "Octavo Grado":   "8°",
    "Noveno Grado":   "9°",
    "Primer Año":     "1°B",
    "Segundo Año":    "2°B",
}
short_grades = [grade_abbrev[g] for g in grade_order]

# ── Derived data ───────────────────────────────────────────────────────────────
schools = (
    df.groupby(["GRUPO", "CODIGO", "NOMBRE"])
    .size()
    .reset_index(name="Total")
    .sort_values(["GRUPO", "NOMBRE"])
)
b1 = schools[schools["GRUPO"] == "B1"].reset_index(drop=True)
b2 = schools[schools["GRUPO"] == "B2"].reset_index(drop=True)

by_grade = df.groupby("GRADO").size().reindex(grade_order).fillna(0).astype(int)

pivot = df.groupby(["GRADO", "GRUPO"]).size().unstack(fill_value=0)
pivot["TOTAL"] = pivot.sum(axis=1)
pivot = pivot.reindex(grade_order).fillna(0).astype(int)

total_schools = df["CODIGO"].nunique()
total_b1      = int(pivot["B1"].sum()) if "B1" in pivot else 0
total_b2      = int(pivot["B2"].sum()) if "B2" in pivot else 0
grand_total   = total_b1 + total_b2

b1_names = set(df[df["GRUPO"] == "B1"]["NOMBRE"].unique())
b2_names = set(df[df["GRUPO"] == "B2"]["NOMBRE"].unique())
shared    = b1_names & b2_names

# ── Per-school, per-grade pivot ───────────────────────────────────────────────
# school_grade_pivot: index = (CODIGO, NOMBRE), columns = grade names, values = counts
school_grade_pivot = (
    df.groupby(["GRUPO", "CODIGO", "NOMBRE", "GRADO"])
    .size()
    .reset_index(name="n")
    .pivot_table(index=["GRUPO", "CODIGO", "NOMBRE"],
                 columns="GRADO", values="n", fill_value=0)
    .reindex(columns=grade_order, fill_value=0)
    .reset_index()
)
school_grade_pivot["TOTAL"] = school_grade_pivot[grade_order].sum(axis=1)
school_grade_pivot = school_grade_pivot.sort_values(["GRUPO", "NOMBRE"])

sgp_b1 = school_grade_pivot[school_grade_pivot["GRUPO"] == "B1"].reset_index(drop=True)
sgp_b2 = school_grade_pivot[school_grade_pivot["GRUPO"] == "B2"].reset_index(drop=True)

# ══════════════════════════════════════════════════════════════════════════════
# TXT OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

# Column widths for the school-by-grade table
W_NUM  = 4    # row number
W_COD  = 7    # school code
W_NAME = 35   # school name (truncated)
W_GRD  = 4    # each grade column
W_TOT  = 7    # total column

# Build separator for the wide table
N_GRADES   = len(grade_order)
WIDE_WIDTH = W_NUM + W_COD + W_NAME + N_GRADES * W_GRD + W_TOT + N_GRADES + 5
SEP_WIDE   = "─" * WIDE_WIDTH

def wide_header():
    """Returns two header lines for the school-by-grade table."""
    grade_cols = "".join(f"{g:>{W_GRD}}" for g in short_grades)
    h1 = (f"  {'#':<{W_NUM}} {'Cód':<{W_COD}} {'Nombre':<{W_NAME}}"
          f"  {grade_cols}  {'Total':>{W_TOT}}")
    h2 = (f"  {'─'*W_NUM} {'─'*W_COD} {'─'*W_NAME}"
          f"  {'─'*(N_GRADES*W_GRD)}  {'─'*W_TOT}")
    return h1, h2

def wide_row(i, codigo, nombre, grade_data, total):
    name_str   = str(nombre).strip()[:W_NAME]
    grade_cols = "".join(f"{int(grade_data.get(g, 0)):>{W_GRD}}" for g in grade_order)
    return (f"  {i:<{W_NUM}} {str(codigo):<{W_COD}} {name_str:<{W_NAME}}"
            f"  {grade_cols}  {int(total):>{W_TOT},}")

def wide_total_row(label, grade_totals, grand):
    """Footer totals row."""
    name_str   = label[:W_NAME + W_COD + W_NUM + 2]
    grade_cols = "".join(f"{int(grade_totals.get(g, 0)):>{W_GRD}}" for g in grade_order)
    pad        = W_NUM + W_COD + W_NAME + 2
    return (f"  {name_str:<{pad}}"
            f"  {grade_cols}  {int(grand):>{W_TOT},}")

lines = []

def sec(title):
    lines.append(f"\n{SEP2}")
    lines.append(f"  {title}")
    lines.append(SEP2)

# 1. Schools per group
sec("NÚMERO DE CENTROS ESCOLARES POR GRUPO")
lines.append(f"  {'Grupo':<10} {'Centros Escolares':>20}")
lines.append(f"  {SEP[:40]}")
lines.append(f"  {'B1':<10} {len(b1):>20,}")
lines.append(f"  {'B2':<10} {len(b2):>20,}")
lines.append(f"  {SEP[:40]}")
lines.append(f"  {'TOTAL':<10} {total_schools:>20,}")
if shared:
    lines.append(f"\n  ⚠  Nota: {len(shared)} centro(s) aparece(n) en ambos grupos:")
    for s in shared:
        lines.append(f"     • {s.strip()}")

# 2. Students by grade
sec("TOTAL DE ESTUDIANTES POR GRADO")
lines.append(f"  {'Grado':<22} {'Total de Estudiantes':>22}")
lines.append(f"  {SEP[:50]}")
for grado in grade_order:
    lines.append(f"  {grado:<22} {by_grade[grado]:>22,}")

# 3. Students by grade + group
sec("ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)")
lines.append(f"  {'Grado':<22} {'B1':>12} {'B2':>12} {'TOTAL':>12}")
lines.append(f"  {SEP[:62]}")
for grado in grade_order:
    b1c = pivot.loc[grado, "B1"] if "B1" in pivot.columns else 0
    b2c = pivot.loc[grado, "B2"] if "B2" in pivot.columns else 0
    tot = pivot.loc[grado, "TOTAL"]
    lines.append(f"  {grado:<22} {b1c:>12,} {b2c:>12,} {tot:>12,}")
lines.append(f"  {SEP[:62]}")
lines.append(f"  {'GRAN TOTAL':<22} {total_b1:>12,} {total_b2:>12,} {grand_total:>12,}")

# 4. All B1 schools — WITH GRADE BREAKDOWN
sec("LISTADO DE CENTROS ESCOLARES — GRUPO B1 (con detalle por grado)")
h1, h2 = wide_header()
lines.append(h1)
lines.append(h2)
for i, row in sgp_b1.iterrows():
    grade_data = {g: row[g] for g in grade_order}
    lines.append(wide_row(i + 1, row["CODIGO"], row["NOMBRE"],
                           grade_data, row["TOTAL"]))
lines.append(f"  {SEP_WIDE}")
# Totals row for B1
b1_grade_totals = {g: int(sgp_b1[g].sum()) for g in grade_order}
lines.append(wide_total_row("TOTAL B1", b1_grade_totals, total_b1))

# 5. All B2 schools — WITH GRADE BREAKDOWN
sec("LISTADO DE CENTROS ESCOLARES — GRUPO B2 (con detalle por grado)")
h1, h2 = wide_header()
lines.append(h1)
lines.append(h2)
for i, row in sgp_b2.iterrows():
    grade_data = {g: row[g] for g in grade_order}
    lines.append(wide_row(i + 1, row["CODIGO"], row["NOMBRE"],
                           grade_data, row["TOTAL"]))
lines.append(f"  {SEP_WIDE}")
b2_grade_totals = {g: int(sgp_b2[g].sum()) for g in grade_order}
lines.append(wide_total_row("TOTAL B2", b2_grade_totals, total_b2))

# 6. Summary
sec("RESUMEN GENERAL")
lines.append(f"  {'Métrica':<45} {'Valor':>10}")
lines.append(f"  {SEP[:60]}")
lines.append(f"  {'Total de Escuelas':<45} {total_schools:>10,}")
lines.append(f"  {'Total de Estudiantes':<45} {grand_total:>10,}")
lines.append(f"  {'Total de Estudiantes B1':<45} {total_b1:>10,}")
lines.append(f"  {'Total de Estudiantes B2':<45} {total_b2:>10,}")
lines.append(f"\n{SEP2}\n")

txt_content = "\n".join(lines)
with open(TXT_OUT, "w", encoding="utf-8") as f:
    f.write(txt_content)
print(f"✔ TXT guardado: {TXT_OUT}")

# ══════════════════════════════════════════════════════════════════════════════
# PDF OUTPUT
# ══════════════════════════════════════════════════════════════════════════════
DARK_BLUE  = colors.HexColor("#1F3864")
MID_BLUE   = colors.HexColor("#2E5FAC")
LIGHT_BLUE = colors.HexColor("#D6E4F0")
WHITE      = colors.white
LIGHT_GRAY = colors.HexColor("#F2F2F2")

styles = getSampleStyleSheet()
title_style = ParagraphStyle("title", parent=styles["Title"],
    textColor=WHITE, backColor=DARK_BLUE, fontSize=14,
    spaceAfter=4, spaceBefore=4, alignment=TA_CENTER,
    leftIndent=-6, rightIndent=-6, leading=20)
h2_style = ParagraphStyle("h2", parent=styles["Heading2"],
    textColor=WHITE, backColor=MID_BLUE, fontSize=11,
    spaceAfter=4, spaceBefore=10, alignment=TA_LEFT,
    leftIndent=0, leading=16)
note_style = ParagraphStyle("note", parent=styles["Normal"],
    fontSize=7, textColor=colors.HexColor("#555555"),
    spaceBefore=2, spaceAfter=2)
normal = styles["Normal"]


def make_table(data, col_widths, header_bg=DARK_BLUE, font_size=8):
    t = Table(data, colWidths=col_widths, repeatRows=1)
    row_count = len(data)
    style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 0),  header_bg),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  font_size),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  5),
        ("TOPPADDING",    (0, 0), (-1, 0),  5),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), font_size),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("ALIGN",         (1, 1), (-1, -1), "RIGHT"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 3),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
        ("TOPPADDING",    (0, 1), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2),
        # Bold + highlight last (totals) row
        ("BACKGROUND",    (0, row_count-1), (-1, row_count-1), LIGHT_BLUE),
        ("FONTNAME",      (0, row_count-1), (-1, row_count-1), "Helvetica-Bold"),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


def make_school_grade_table(sgp_df, total_students, W):
    """
    Build a PDF table with columns:
      # | Código | Nombre | 2° | 3° | 4° | 5° | 6° | 7° | 8° | 9° | 1°B | 2°B | Total
    Uses a small font to fit all grade columns on one page width.
    """
    header = ["#", "Código", "Nombre"] + short_grades + ["Total"]
    data   = [header]

    for i, row in sgp_df.iterrows():
        data_row = [
            str(i + 1),
            str(row["CODIGO"]),
            row["NOMBRE"].strip(),
        ] + [str(int(row[g])) if int(row[g]) > 0 else "—"
             for g in grade_order] + [f"{int(row['TOTAL']):,}"]
        data.append(data_row)

    # Totals row
    grade_totals = [int(sgp_df[g].sum()) for g in grade_order]
    data.append(
        ["", "", "TOTAL"]
        + [f"{v:,}" for v in grade_totals]
        + [f"{total_students:,}"]
    )

    # Column widths — tight to fit 13 columns in letter width
    n_grade_cols = len(grade_order)
    w_num    = W * 0.03
    w_cod    = W * 0.07
    w_name   = W * 0.25
    w_grade  = (W - w_num - w_cod - w_name - W * 0.08) / n_grade_cols
    w_total  = W * 0.08
    col_widths = [w_num, w_cod, w_name] + [w_grade] * n_grade_cols + [w_total]

    t = Table(data, colWidths=col_widths, repeatRows=1)
    row_count = len(data)
    style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 0),  DARK_BLUE),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  6),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  4),
        ("TOPPADDING",    (0, 0), (-1, 0),  4),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 6),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        # Numeric columns right-aligned (grade cols + total)
        ("ALIGN",         (3, 1), (-1, -1), "RIGHT"),
        # Name col left-aligned
        ("ALIGN",         (2, 1), (2, -1),  "LEFT"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ("TOPPADDING",    (0, 1), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 1),
        # Totals row
        ("BACKGROUND",    (0, row_count-1), (-1, row_count-1), LIGHT_BLUE),
        ("FONTNAME",      (0, row_count-1), (-1, row_count-1), "Helvetica-Bold"),
        ("FONTSIZE",      (0, row_count-1), (-1, row_count-1), 6),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


doc = SimpleDocTemplate(PDF_OUT, pagesize=letter,
    leftMargin=0.5*inch, rightMargin=0.5*inch,
    topMargin=0.6*inch,  bottomMargin=0.6*inch)
story = []
W = letter[0] - 1.0*inch   # usable width

# Title
story.append(Paragraph("RECOLECCIÓN DE DATOS G1-G2 — 25/05/2026", title_style))
story.append(Spacer(1, 10))

# ── Section 1: schools per group ─────────────────────────────────────────────
story.append(Paragraph("NÚMERO DE CENTROS ESCOLARES POR GRUPO", h2_style))
data = [["Grupo", "Centros Escolares"],
        ["B1",    f"{len(b1):,}"],
        ["B2",    f"{len(b2):,}"],
        ["TOTAL", f"{total_schools:,}"]]
story.append(make_table(data, [W*0.5, W*0.5]))
if shared:
    msg = (f"⚠ Nota: {len(shared)} centro(s) aparece(n) en ambos grupos: "
           + ", ".join(s.strip() for s in shared))
    story.append(Spacer(1, 4))
    story.append(Paragraph(msg, normal))

# ── Section 2: students by grade ─────────────────────────────────────────────
story.append(Paragraph("TOTAL DE ESTUDIANTES POR GRADO", h2_style))
data = [["Grado", "Total de Estudiantes"]]
for g in grade_order:
    data.append([g, f"{by_grade[g]:,}"])
data.append(["GRAN TOTAL", f"{grand_total:,}"])
story.append(make_table(data, [W*0.6, W*0.4]))

# ── Section 3: students by grade + group ─────────────────────────────────────
story.append(Paragraph("ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)", h2_style))
data = [["Grado", "B1", "B2", "TOTAL"]]
for g in grade_order:
    b1c = int(pivot.loc[g, "B1"]) if "B1" in pivot.columns else 0
    b2c = int(pivot.loc[g, "B2"]) if "B2" in pivot.columns else 0
    tot = int(pivot.loc[g, "TOTAL"])
    data.append([g, f"{b1c:,}", f"{b2c:,}", f"{tot:,}"])
data.append(["GRAN TOTAL", f"{total_b1:,}", f"{total_b2:,}", f"{grand_total:,}"])
story.append(make_table(data, [W*0.4, W*0.2, W*0.2, W*0.2]))

# ── Section 4: B1 schools with grade breakdown ────────────────────────────────
story.append(PageBreak())
story.append(Paragraph("LISTADO DE CENTROS ESCOLARES — GRUPO B1 (detalle por grado)",
                        h2_style))
story.append(Paragraph(
    "Columnas de grado: 2°=Segundo, 3°=Tercer, 4°=Cuarto, 5°=Quinto, "
    "6°=Sexto, 7°=Séptimo, 8°=Octavo, 9°=Noveno, 1°B=Primer Año Bach., "
    "2°B=Segundo Año Bach.  |  — = sin estudiantes en ese grado.",
    note_style))
story.append(make_school_grade_table(sgp_b1, total_b1, W))

# ── Section 5: B2 schools with grade breakdown ────────────────────────────────
story.append(PageBreak())
story.append(Paragraph("LISTADO DE CENTROS ESCOLARES — GRUPO B2 (detalle por grado)",
                        h2_style))
story.append(Paragraph(
    "Columnas de grado: 2°=Segundo, 3°=Tercer, 4°=Cuarto, 5°=Quinto, "
    "6°=Sexto, 7°=Séptimo, 8°=Octavo, 9°=Noveno, 1°B=Primer Año Bach., "
    "2°B=Segundo Año Bach.  |  — = sin estudiantes en ese grado.",
    note_style))
story.append(make_school_grade_table(sgp_b2, total_b2, W))

# ── Section 6: summary ────────────────────────────────────────────────────────
story.append(PageBreak())
story.append(Paragraph("RESUMEN GENERAL", h2_style))
data = [["Métrica", "Valor"],
        ["Total de Escuelas",       f"{total_schools:,}"],
        ["Total de Estudiantes",     f"{grand_total:,}"],
        ["Total de Estudiantes B1",  f"{total_b1:,}"],
        ["Total de Estudiantes B2",  f"{total_b2:,}"]]
story.append(make_table(data, [W*0.7, W*0.3]))

doc.build(story)
print(f"✔ PDF guardado: {PDF_OUT}")


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

# ── openpyxl style helpers ────────────────────────────────────────────────────
XL_DARK_BLUE  = "1F3864"
XL_MID_BLUE   = "2E5FAC"
XL_LIGHT_BLUE = "D6E4F0"
XL_LIGHT_GRAY = "F2F2F2"
XL_WHITE      = "FFFFFF"
XL_TOTAL_BG   = "DDEBF7"

THIN_BORDER = Border(
    left=Side(style="thin",   color="CCCCCC"),
    right=Side(style="thin",  color="CCCCCC"),
    top=Side(style="thin",    color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
THICK_BOTTOM = Border(
    left=Side(style="thin",    color="CCCCCC"),
    right=Side(style="thin",   color="CCCCCC"),
    top=Side(style="thin",     color="CCCCCC"),
    bottom=Side(style="medium", color="1F3864"),
)


def xl_header_cell(cell, value, bg=XL_DARK_BLUE, font_size=10, wrap=False):
    cell.value = value
    cell.font  = Font(bold=True, color=XL_WHITE, size=font_size,
                      name="Calibri")
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=wrap)
    cell.border = THIN_BORDER


def xl_data_cell(cell, value, align="right", bold=False,
                 bg=None, num_fmt=None):
    cell.value = value
    cell.font  = Font(bold=bold, size=9, name="Calibri")
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = THIN_BORDER
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if num_fmt:
        cell.number_format = num_fmt


def xl_total_row(ws, row, values, alignments=None):
    """Write a totals row with highlighted background."""
    if alignments is None:
        alignments = ["left"] + ["right"] * (len(values) - 1)
    for col_i, (val, aln) in enumerate(zip(values, alignments), start=1):
        c = ws.cell(row=row, column=col_i)
        xl_data_cell(c, val, align=aln, bold=True,
                     bg=XL_LIGHT_BLUE,
                     num_fmt="#,##0" if isinstance(val, (int, float)) else None)


def xl_section_title(ws, row, title, n_cols):
    """Write a full-width section title row (merged)."""
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=n_cols)
    c = ws.cell(row=row, column=1, value=title)
    c.font      = Font(bold=True, color=XL_WHITE, size=11, name="Calibri")
    c.fill      = PatternFill("solid", fgColor=XL_MID_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center",
                             indent=1)
    c.border    = THIN_BORDER
    ws.row_dimensions[row].height = 18


def xl_auto_col_width(ws, min_width=8, max_width=55):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(
            max(max_len + 2, min_width), max_width)


wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# ────────────────────────────────────────────────────────────────────────────
# Sheet 1 — Centros por Grupo
# ────────────────────────────────────────────────────────────────────────────
ws1 = wb.create_sheet("Centros por Grupo")
xl_section_title(ws1, 1, "NÚMERO DE CENTROS ESCOLARES POR GRUPO", 2)
for ci, hdr in enumerate(["Grupo", "Centros Escolares"], start=1):
    xl_header_cell(ws1.cell(row=2, column=ci), hdr, bg=XL_DARK_BLUE)
ws1.row_dimensions[2].height = 16
for r, (grp, cnt) in enumerate([("B1", len(b1)), ("B2", len(b2))], start=3):
    xl_data_cell(ws1.cell(row=r, column=1), grp, align="left",
                 bg=XL_WHITE if r % 2 == 1 else XL_LIGHT_GRAY)
    xl_data_cell(ws1.cell(row=r, column=2), cnt, num_fmt="#,##0",
                 bg=XL_WHITE if r % 2 == 1 else XL_LIGHT_GRAY)
xl_total_row(ws1, 5, ["TOTAL", total_schools])
if shared:
    ws1.cell(row=7, column=1,
             value=f"⚠ Centros en ambos grupos: {', '.join(s.strip() for s in shared)}")
    ws1.cell(row=7, column=1).font = Font(color="C00000", size=9, name="Calibri")
xl_auto_col_width(ws1)
ws1.freeze_panes = "A3"

# ────────────────────────────────────────────────────────────────────────────
# Sheet 2 — Estudiantes por Grado
# ────────────────────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("Estudiantes por Grado")
xl_section_title(ws2, 1, "TOTAL DE ESTUDIANTES POR GRADO", 2)
for ci, hdr in enumerate(["Grado", "Total de Estudiantes"], start=1):
    xl_header_cell(ws2.cell(row=2, column=ci), hdr)
ws2.row_dimensions[2].height = 16
for r, grado in enumerate(grade_order, start=3):
    bg = XL_WHITE if r % 2 == 1 else XL_LIGHT_GRAY
    xl_data_cell(ws2.cell(row=r, column=1), grado,  align="left", bg=bg)
    xl_data_cell(ws2.cell(row=r, column=2), int(by_grade[grado]),
                 num_fmt="#,##0", bg=bg)
xl_total_row(ws2, r + 1, ["GRAN TOTAL", grand_total])
xl_auto_col_width(ws2)
ws2.freeze_panes = "A3"

# ────────────────────────────────────────────────────────────────────────────
# Sheet 3 — Estudiantes por Grado y Grupo
# ────────────────────────────────────────────────────────────────────────────
ws3 = wb.create_sheet("Grado x Grupo")
xl_section_title(ws3, 1, "ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)", 4)
for ci, hdr in enumerate(["Grado", "B1", "B2", "Total"], start=1):
    xl_header_cell(ws3.cell(row=2, column=ci), hdr)
ws3.row_dimensions[2].height = 16
for r, grado in enumerate(grade_order, start=3):
    bg  = XL_WHITE if r % 2 == 1 else XL_LIGHT_GRAY
    b1c = int(pivot.loc[grado, "B1"])   if "B1"    in pivot.columns else 0
    b2c = int(pivot.loc[grado, "B2"])   if "B2"    in pivot.columns else 0
    tot = int(pivot.loc[grado, "TOTAL"])
    xl_data_cell(ws3.cell(row=r, column=1), grado, align="left", bg=bg)
    for ci, val in enumerate([b1c, b2c, tot], start=2):
        xl_data_cell(ws3.cell(row=r, column=ci), val, num_fmt="#,##0", bg=bg)
xl_total_row(ws3, r + 1, ["GRAN TOTAL", total_b1, total_b2, grand_total])
xl_auto_col_width(ws3)
ws3.freeze_panes = "A3"

# ────────────────────────────────────────────────────────────────────────────
# Sheet 4 — Centros B1 con detalle por grado
# Sheet 5 — Centros B2 con detalle por grado
# ────────────────────────────────────────────────────────────────────────────
def write_school_grade_sheet(wb, sheet_name, sgp_df, grupo_label, total_est):
    ws = wb.create_sheet(sheet_name)
    n_cols = 3 + len(grade_order) + 1    # # + Código + Nombre + grades + Total

    xl_section_title(ws, 1,
                     f"LISTADO DE CENTROS ESCOLARES — GRUPO {grupo_label} "
                     f"(detalle por grado)", n_cols)

    # Header row
    headers = ["#", "Código", "Nombre"] + short_grades + ["Total"]
    for ci, hdr in enumerate(headers, start=1):
        xl_header_cell(ws.cell(row=2, column=ci), hdr,
                       bg=XL_DARK_BLUE, font_size=9, wrap=True)
    ws.row_dimensions[2].height = 22

    # Data rows
    for r, row in sgp_df.iterrows():
        bg = XL_WHITE if (r % 2 == 0) else XL_LIGHT_GRAY
        xl_data_cell(ws.cell(row=r + 3, column=1), r + 1,
                     align="center", bg=bg)
        xl_data_cell(ws.cell(row=r + 3, column=2), str(row["CODIGO"]),
                     align="left", bg=bg)
        xl_data_cell(ws.cell(row=r + 3, column=3), row["NOMBRE"].strip(),
                     align="left", bg=bg)
        for gi, grado in enumerate(grade_order, start=4):
            val = int(row[grado])
            xl_data_cell(ws.cell(row=r + 3, column=gi),
                         val if val > 0 else None,
                         num_fmt="#,##0", bg=bg)
        xl_data_cell(ws.cell(row=r + 3, column=n_cols),
                     int(row["TOTAL"]), num_fmt="#,##0",
                     bold=True, bg=bg)

    # Totals row
    total_row = len(sgp_df) + 3
    grade_totals = [int(sgp_df[g].sum()) for g in grade_order]
    vals  = ["", "", f"TOTAL {grupo_label}"] + grade_totals + [total_est]
    alns  = ["center", "left", "left"] + ["right"] * (len(grade_order) + 1)
    for ci, (val, aln) in enumerate(zip(vals, alns), start=1):
        c = ws.cell(row=total_row, column=ci)
        xl_data_cell(c, val if val != 0 else None,
                     align=aln, bold=True, bg=XL_LIGHT_BLUE,
                     num_fmt="#,##0" if isinstance(val, int) and val != 0 else None)

    # Column widths: narrow for grade cols, wider for name
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 9
    ws.column_dimensions["C"].width = 38
    for gi in range(4, 4 + len(grade_order)):
        ws.column_dimensions[get_column_letter(gi)].width = 5
    ws.column_dimensions[get_column_letter(n_cols)].width = 8

    # Freeze header
    ws.freeze_panes = "A3"

    # Add grade legend as a note below the table
    note_row = total_row + 2
    ws.cell(row=note_row, column=1,
            value=("Leyenda de grados: 2°=Segundo Grado, 3°=Tercer Grado, "
                   "4°=Cuarto, 5°=Quinto, 6°=Sexto, 7°=Séptimo, 8°=Octavo, "
                   "9°=Noveno, 1°B=Primer Año Bachillerato, "
                   "2°B=Segundo Año Bachillerato"))
    ws.cell(row=note_row, column=1).font = Font(italic=True, size=8,
                                                color="555555", name="Calibri")
    ws.merge_cells(start_row=note_row, start_column=1,
                   end_row=note_row, end_column=n_cols)

write_school_grade_sheet(wb, "Centros B1 por Grado", sgp_b1, "B1", total_b1)
write_school_grade_sheet(wb, "Centros B2 por Grado", sgp_b2, "B2", total_b2)

# ────────────────────────────────────────────────────────────────────────────
# Sheet 6 — Resumen General
# ────────────────────────────────────────────────────────────────────────────
ws6 = wb.create_sheet("Resumen General")
xl_section_title(ws6, 1, "RESUMEN GENERAL", 2)
for ci, hdr in enumerate(["Métrica", "Valor"], start=1):
    xl_header_cell(ws6.cell(row=2, column=ci), hdr)
ws6.row_dimensions[2].height = 16
summary_rows = [
    ("Total de Escuelas",           total_schools),
    ("Total de Estudiantes",        grand_total),
    ("Total de Estudiantes B1",     total_b1),
    ("Total de Estudiantes B2",     total_b2),
]
for r, (label, val) in enumerate(summary_rows, start=3):
    bg = XL_WHITE if r % 2 == 1 else XL_LIGHT_GRAY
    xl_data_cell(ws6.cell(row=r, column=1), label, align="left", bg=bg)
    xl_data_cell(ws6.cell(row=r, column=2), val,   num_fmt="#,##0", bg=bg)
xl_auto_col_width(ws6)

wb.save(XLS_OUT)
print(f"✔ Excel guardado: {XLS_OUT}")

# ── Also print to console ─────────────────────────────────────────────────────
print(txt_content)