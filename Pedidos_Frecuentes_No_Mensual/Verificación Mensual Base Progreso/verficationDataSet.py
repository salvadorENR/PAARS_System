import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

FILE = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Verificación Mensual Base Progreso\Recolección de datos G1-G2_25_05_2026.xlsx"
TXT_OUT = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Verificación Mensual Base Progreso\resumen_escuelas.txt"
PDF_OUT = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Verificación Mensual Base Progreso\resumen_escuelas.pdf"

# ── Load data ──────────────────────────────────────────────────────────────────
df = pd.read_excel(FILE)

SEP  = "─" * 72
SEP2 = "═" * 72

# ── Derived data ───────────────────────────────────────────────────────────────
schools = (
    df.groupby(["GRUPO", "CODIGO", "NOMBRE"])
    .size()
    .reset_index(name="Total")
    .sort_values(["GRUPO", "NOMBRE"])
)
b1 = schools[schools["GRUPO"] == "B1"].reset_index(drop=True)
b2 = schools[schools["GRUPO"] == "B2"].reset_index(drop=True)

grade_order = [
    "Segundo Grado", "Tercer Grado", "Cuarto Grado", "Quinto Grado",
    "Sexto Grado", "Séptimo Grado", "Octavo Grado", "Noveno Grado",
    "Primer Año", "Segundo Año"
]
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

# ══════════════════════════════════════════════════════════════════════════════
# TXT OUTPUT
# ══════════════════════════════════════════════════════════════════════════════
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

# 4. All schools — B1
sec("LISTADO DE CENTROS ESCOLARES — GRUPO B1")
lines.append(f"  {'#':<5} {'Código':<8} {'Nombre':<55} {'Estudiantes':>11}")
lines.append(f"  {SEP}")
for i, row in b1.iterrows():
    name = row["NOMBRE"].strip()[:54]
    lines.append(f"  {i+1:<5} {str(row['CODIGO']):<8} {name:<55} {row['Total']:>11,}")
lines.append(f"  {SEP}")
lines.append(f"  {'TOTAL B1':<68} {total_b1:>11,}")

# 5. All schools — B2
sec("LISTADO DE CENTROS ESCOLARES — GRUPO B2")
lines.append(f"  {'#':<5} {'Código':<8} {'Nombre':<55} {'Estudiantes':>11}")
lines.append(f"  {SEP}")
for i, row in b2.iterrows():
    name = row["NOMBRE"].strip()[:54]
    lines.append(f"  {i+1:<5} {str(row['CODIGO']):<8} {name:<55} {row['Total']:>11,}")
lines.append(f"  {SEP}")
lines.append(f"  {'TOTAL B2':<68} {total_b2:>11,}")

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
normal = styles["Normal"]

def make_table(data, col_widths, header_bg=DARK_BLUE):
    t = Table(data, colWidths=col_widths, repeatRows=1)
    row_count = len(data)
    style_cmds = [
        ("BACKGROUND",  (0, 0), (-1, 0),  header_bg),
        ("TEXTCOLOR",   (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0),  8),
        ("ALIGN",       (0, 0), (-1, 0),  "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING",    (0, 0), (-1, 0), 5),
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_GRAY]),
        ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("ALIGN",       (-1, 1), (-1, -1), "RIGHT"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    # Bold last row (totals)
    style_cmds += [
        ("BACKGROUND", (0, row_count-1), (-1, row_count-1), LIGHT_BLUE),
        ("FONTNAME",   (0, row_count-1), (-1, row_count-1), "Helvetica-Bold"),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t

doc = SimpleDocTemplate(PDF_OUT, pagesize=letter,
    leftMargin=0.6*inch, rightMargin=0.6*inch,
    topMargin=0.6*inch, bottomMargin=0.6*inch)
story = []
W = letter[0] - 1.2*inch  # usable width

# Title
story.append(Paragraph("RECOLECCIÓN DE DATOS G1-G2 — 25/05/2026", title_style))
story.append(Spacer(1, 10))

# ── Section 1: schools per group ──
story.append(Paragraph("NÚMERO DE CENTROS ESCOLARES POR GRUPO", h2_style))
data = [["Grupo", "Centros Escolares"],
        ["B1", f"{len(b1):,}"],
        ["B2", f"{len(b2):,}"],
        ["TOTAL", f"{total_schools:,}"]]
story.append(make_table(data, [W*0.5, W*0.5]))
if shared:
    msg = f"⚠ Nota: {len(shared)} centro(s) aparece(n) en ambos grupos: " + \
          ", ".join(s.strip() for s in shared)
    story.append(Spacer(1, 4))
    story.append(Paragraph(msg, normal))

# ── Section 2: students by grade ──
story.append(Paragraph("TOTAL DE ESTUDIANTES POR GRADO", h2_style))
data = [["Grado", "Total de Estudiantes"]]
for g in grade_order:
    data.append([g, f"{by_grade[g]:,}"])
data.append(["GRAN TOTAL", f"{grand_total:,}"])
story.append(make_table(data, [W*0.6, W*0.4]))

# ── Section 3: students by grade + group ──
story.append(Paragraph("ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)", h2_style))
data = [["Grado", "B1", "B2", "TOTAL"]]
for g in grade_order:
    b1c = int(pivot.loc[g, "B1"]) if "B1" in pivot.columns else 0
    b2c = int(pivot.loc[g, "B2"]) if "B2" in pivot.columns else 0
    tot = int(pivot.loc[g, "TOTAL"])
    data.append([g, f"{b1c:,}", f"{b2c:,}", f"{tot:,}"])
data.append(["GRAN TOTAL", f"{total_b1:,}", f"{total_b2:,}", f"{grand_total:,}"])
story.append(make_table(data, [W*0.4, W*0.2, W*0.2, W*0.2]))

# ── Section 4: all B1 schools ──
story.append(PageBreak())
story.append(Paragraph("LISTADO DE CENTROS ESCOLARES — GRUPO B1", h2_style))
data = [["#", "Código", "Nombre", "Estudiantes"]]
for i, row in b1.iterrows():
    data.append([str(i+1), str(row["CODIGO"]), row["NOMBRE"].strip(), f"{row['Total']:,}"])
data.append(["", "", "TOTAL B1", f"{total_b1:,}"])
story.append(make_table(data, [W*0.06, W*0.10, W*0.68, W*0.16]))

# ── Section 5: all B2 schools ──
story.append(PageBreak())
story.append(Paragraph("LISTADO DE CENTROS ESCOLARES — GRUPO B2", h2_style))
data = [["#", "Código", "Nombre", "Estudiantes"]]
for i, row in b2.iterrows():
    data.append([str(i+1), str(row["CODIGO"]), row["NOMBRE"].strip(), f"{row['Total']:,}"])
data.append(["", "", "TOTAL B2", f"{total_b2:,}"])
story.append(make_table(data, [W*0.06, W*0.10, W*0.68, W*0.16]))

# ── Section 6: summary ──
story.append(PageBreak())
story.append(Paragraph("RESUMEN GENERAL", h2_style))
data = [["Métrica", "Valor"],
        ["Total de Escuelas",          f"{total_schools:,}"],
        ["Total de Estudiantes",        f"{grand_total:,}"],
        ["Total de Estudiantes B1",     f"{total_b1:,}"],
        ["Total de Estudiantes B2",     f"{total_b2:,}"]]
story.append(make_table(data, [W*0.7, W*0.3]))

doc.build(story)
print(f"✔ PDF guardado: {PDF_OUT}")

# ── Also print to console ──────────────────────────────────────────────────────
print(txt_content)