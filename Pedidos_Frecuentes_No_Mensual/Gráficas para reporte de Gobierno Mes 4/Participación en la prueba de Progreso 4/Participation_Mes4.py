import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import glob

# ─── PATHS ────────────────────────────────────────────────────────────────────
INPUT_DIR  = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\04_PROGRESO_Junio\Interim_CSVs\Resultados"
OUTPUT_DIR = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\Gráficas para reporte de Gobierno Mes 4\Participación en la prueba de Progreso 4"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Participacion_Progreso4.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── GRADE ORDER (2 to 10) ────────────────────────────────────────────────────
# Both º (U+00BA masculine ordinal) and ° (U+00B0 degree sign) are normalized
# to ° so CSVs using either symbol are treated identically.
GRADE_ORDER = ["2°", "3°", "4°", "5°", "6°", "7°", "8°", "9°", "10°", "11°"]

def normalize_grade(val):
    return str(val).replace("\u00ba", "\u00b0").strip()


# ─── STYLES ───────────────────────────────────────────────────────────────────
SUBHDR_FILL = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
TOTAL_FILL  = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
ALT_FILL    = PatternFill("solid", start_color="EBF5FB", end_color="EBF5FB")

WHITE_BOLD  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
NORMAL_FONT = Font(name="Arial", size=10)
TOTAL_FONT  = Font(name="Arial", bold=True, size=10)

CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT        = Alignment(horizontal="left",   vertical="center")

THIN        = Side(style="thin",   color="B0C4DE")
MED         = Side(style="medium", color="1F4E79")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_MED  = Border(left=THIN, right=THIN, top=MED,  bottom=MED)


# ─── LOAD & COMBINE CSVs BY SUBJECT PREFIX ────────────────────────────────────
def load_subject_csvs(prefix):
    pattern = os.path.join(INPUT_DIR, f"{prefix}-*.csv")
    files = glob.glob(pattern)
    if not files:
        print(f"  ⚠  No CSV files found for prefix '{prefix}' in {INPUT_DIR}")
        return pd.DataFrame()
    frames = []
    for f in files:
        df = pd.read_csv(f, encoding="utf-8")
        df["Grado"] = df["Grado"].apply(normalize_grade)
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)
    print(f"  ✔  {prefix}: {len(files)} file(s), {len(combined):,} rows — "
          f"grades found: {sorted(combined['Grado'].unique())}")
    return combined


# ─── BUILD PIVOT: rows=school, columns=grade ──────────────────────────────────
def build_pivot(df):
    pivot = (
        df.groupby(["Nro de centro", "Centro", "Grado"])
        .size()
        .reset_index(name="Cantidad")
        .pivot_table(index=["Nro de centro", "Centro"], columns="Grado",
                     values="Cantidad", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    pivot.columns.name = None
    known = [g for g in GRADE_ORDER if g in pivot.columns]
    unknown = sorted([g for g in pivot.columns if g not in GRADE_ORDER and g not in ("Nro de centro", "Centro")])
    if unknown:
        print(f"  WARNING  Grades not in GRADE_ORDER (appended at end): {unknown}")
    present_grades = known + unknown
    pivot = pivot[["Nro de centro", "Centro"] + present_grades]
    pivot["Total General"] = pivot[present_grades].sum(axis=1)
    return pivot, present_grades


# ─── WRITE SHEET ──────────────────────────────────────────────────────────────
def write_sheet(wb, sheet_name, pivot, grade_cols):
    ws = wb.create_sheet(title=sheet_name)
    all_cols = ["Nro de centro", "Centro"] + grade_cols + ["Total General"]
    n_cols   = len(all_cols)

    # Row 1 – column headers (no merged cells)
    headers = {"Nro de centro": "Nro de Centro",
               "Centro":        "Centro Educativo",
               "Total General": "Total General"}
    for ci, col in enumerate(all_cols, start=1):
        c = ws.cell(row=1, column=ci, value=headers.get(col, col))
        c.font = WHITE_BOLD; c.fill = SUBHDR_FILL
        c.alignment = CENTER; c.border = BORDER_MED

    # Data rows
    for ri, (_, row) in enumerate(pivot.iterrows(), start=2):
        row_fill = ALT_FILL if ri % 2 == 0 else PatternFill()
        for ci, col in enumerate(all_cols, start=1):
            c = ws.cell(row=ri, column=ci, value=row[col])
            c.border = BORDER_THIN
            if ci == 2:
                c.alignment = LEFT;   c.font = NORMAL_FONT
            else:
                c.alignment = CENTER; c.font = NORMAL_FONT
            if ci == n_cols:
                c.font = TOTAL_FONT; c.fill = TOTAL_FILL
            elif row_fill:
                c.fill = row_fill

    # Grand total row (no merged cells)
    total_row = len(pivot) + 2
    for col in range(1, n_cols + 1):
        ws.cell(row=total_row, column=col).fill = TOTAL_FILL
    ws.cell(row=total_row, column=1, value="TOTAL GENERAL").font = TOTAL_FONT
    ws.cell(row=total_row, column=1).alignment = CENTER
    ws.cell(row=total_row, column=2).fill = TOTAL_FILL

    for ci in range(3, n_cols + 1):
        col_letter = get_column_letter(ci)
        c = ws.cell(row=total_row, column=ci,
                    value=f"=SUM({col_letter}2:{col_letter}{total_row - 1})")
        c.font = TOTAL_FONT; c.fill = TOTAL_FILL
        c.alignment = CENTER; c.border = BORDER_MED

    # Column widths & row heights
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 46
    for ci in range(3, n_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 11
    ws.row_dimensions[1].height = 26

    ws.freeze_panes = "C2"


# ─── MAIN ─────────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)

for subject_prefix, sheet_label in [("LEC", "Lengua (LEC)"),
                                     ("MAT", "Matemática (MAT)")]:
    print(f"\nProcessing {sheet_label}...")
    df = load_subject_csvs(subject_prefix)
    if df.empty:
        continue
    pivot, grade_cols = build_pivot(df)
    write_sheet(wb, sheet_label, pivot, grade_cols)
    print(f"  → {len(pivot)} schools, grades: {grade_cols}")

wb.save(OUTPUT_FILE)
print(f"\n✔  Archivo guardado en:\n   {OUTPUT_FILE}")