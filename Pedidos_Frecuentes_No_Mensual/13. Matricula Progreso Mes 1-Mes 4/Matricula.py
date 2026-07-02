"""
Cobertura_Participacion_Por_Mes.py
===================================
Genera para cada mes (Prueba 1–4):

  1. Tabla de Cobertura  → PDF + Excel por mes
  2. Reporte de Escuelas → TXT + PDF + Excel por mes
  3. Participación por Escuela → Excel por mes

Archivos maestros nuevos:
  4. Participacion_Maestra_Por_Escuela.xlsx
       Un sheet por mes.  Columnas por escuela:
         Matrícula Esperada | MAT Válidos | MAT Válidos % | LEC Válidos | LEC Válidos %
         | MAT Total | MAT Total % | LEC Total | LEC Total %
         | 4 grados con menor participación (todos los casos)

  5. Sheet extra "Centro Virtual":  filas=Mes 1–4, col= conteo 99999/99998 + % del total

  6. Sheet extra "Grado x Mes (todos)":
         filas=grados 2°–11°, columnas=Mes1..4 MAT/LEC (todos los casos,
         sin exclusión de Centro Virtual ni filtro de tiempo) + Porcentajes

Todas las salidas van a:
  PATH_OUTPUT_BASE (reportes mensuales)
  PATH_OUTPUT_COVERAGE (maestros de cobertura)
"""

import os, sys, glob, re
import unicodedata
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ══════════════════════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
MESES = [
    {
        "label":     "Mes 1 (Marzo)",
        "short":     "Mes1_Marzo",
        "matricula": r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4\Matricula_P1.xlsx",
        "csvs":      r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados",
    },
    {
        "label":     "Mes 2 (Abril)",
        "short":     "Mes2_Abril",
        "matricula": r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4\Matricula_P2.xlsx",
        "csvs":      r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados",
    },
    {
        "label":     "Mes 3 (Mayo)",
        "short":     "Mes3_Mayo",
        "matricula": r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4\Matricula_P3.xlsx",
        "csvs":      r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados",
    },
    {
        "label":     "Mes 4 (Junio)",
        "short":     "Mes4_Junio",
        "matricula": r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4\Matricula_P4.xlsx",
        "csvs":      r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\04_PROGRESO_Junio\Interim_CSVs\Resultados",
    },
]

# Rutas unificadas
PATH_OUTPUT_BASE     = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4"
PATH_OUTPUT_COVERAGE = r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4"

VIRTUAL_CODES = {"99999", "99998"}

GRADE_ORDER_LONG = [
    "Segundo Grado", "Tercer Grado", "Cuarto Grado", "Quinto Grado",
    "Sexto Grado", "Séptimo Grado", "Octavo Grado", "Noveno Grado",
    "Primer Año", "Segundo Año",
]
GRADE_ORDER_SHORT = ["2°","3°","4°","5°","6°","7°","8°","9°","10°","11°"]
LONG_TO_SHORT = dict(zip(GRADE_ORDER_LONG, GRADE_ORDER_SHORT))
SHORT_TO_LONG = dict(zip(GRADE_ORDER_SHORT, GRADE_ORDER_LONG))

GRADE_DISPLAY = {
    "2°":  "2.° (Segundo)",   "3°":  "3.° (Tercero)",
    "4°":  "4.° (Cuarto)",    "5°":  "5.° (Quinto)",
    "6°":  "6.° (Sexto)",     "7°":  "7.° (Séptimo)",
    "8°":  "8.° (Octavo)",    "9°":  "9.° (Noveno)",
    "10°": "10.° (Primer Año Bach.)", "11°": "11.° (Segundo Año Bach.)",
}
GRADE_ABBREV = {
    "Segundo Grado":"2°","Tercer Grado":"3°","Cuarto Grado":"4°",
    "Quinto Grado":"5°","Sexto Grado":"6°","Séptimo Grado":"7°",
    "Octavo Grado":"8°","Noveno Grado":"9°","Primer Año":"10°","Segundo Año":"11°",
}
SHORT_GRADES_ABBREV = [GRADE_ABBREV[g] for g in GRADE_ORDER_LONG]

XL_DARK_BLUE="1F3864"; XL_MID_BLUE="2E5FAC"; XL_LIGHT_BLUE="D6E4F0"
XL_GRAY="F2F2F2"; XL_WHITE="FFFFFF"
THIN = Side(style="thin",   color="CCCCCC")
MED  = Side(style="medium", color="1F3864")
B_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_MED  = Border(left=THIN, right=THIN, top=MED,  bottom=MED)

RL_DARK  = colors.HexColor("#1F3864")
RL_MID   = colors.HexColor("#2E5FAC")
RL_LIGHT = colors.HexColor("#D6E4F0")
RL_GRAY  = colors.HexColor("#F2F2F2")

# ══════════════════════════════════════════════════════════════════════════════
# 2. HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def normalize_grade_csv(val):
    return str(val).replace("\u00ba", "\u00b0").strip()

def quitar_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s))
                   if unicodedata.category(c) != "Mn")

def excluir_por_tiempo(val):
    if pd.isna(val): return False
    v = quitar_tildes(str(val).lower())
    return "5 min" in v or "demora" in v or "duracion" in v

def xl_hdr(cell, value, bg=XL_DARK_BLUE, sz=10, wrap=True, color=XL_WHITE):
    cell.value = value
    cell.font  = Font(bold=True, color=color, size=sz, name="Calibri")
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = B_THIN

def xl_data(cell, value, align="right", bold=False, bg=None, fmt=None):
    cell.value = value
    cell.font  = Font(bold=bold, size=9, name="Calibri")
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = B_THIN
    if bg:  cell.fill = PatternFill("solid", fgColor=bg)
    if fmt: cell.number_format = fmt

def xl_total_row(ws, row, values, aligns=None):
    if aligns is None:
        aligns = ["left"] + ["right"] * (len(values)-1)
    for ci,(val,aln) in enumerate(zip(values,aligns),start=1):
        c = ws.cell(row=row, column=ci)
        xl_data(c, val, align=aln, bold=True, bg=XL_LIGHT_BLUE,
                fmt="#,##0" if isinstance(val,(int,float)) else None)

def xl_auto_width(ws, mn=8, mx=55):
    for col in ws.columns:
        ml = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(ml+2,mn),mx)

def rl_make_table(data, col_widths, hdr_bg=RL_DARK, fs=8):
    t = Table(data, colWidths=col_widths, repeatRows=1)
    n = len(data)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0),  hdr_bg),
        ("TEXTCOLOR",     (0,0),(-1,0),  colors.white),
        ("FONTNAME",      (0,0),(-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,0),  fs),
        ("ALIGN",         (0,0),(-1,0),  "CENTER"),
        ("BOTTOMPADDING", (0,0),(-1,0),  5),
        ("TOPPADDING",    (0,0),(-1,0),  5),
        ("FONTNAME",      (0,1),(-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,1),(-1,-1), fs),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, RL_GRAY]),
        ("GRID",          (0,0),(-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("ALIGN",         (1,1),(-1,-1), "RIGHT"),
        ("LEFTPADDING",   (0,0),(-1,-1), 3),
        ("RIGHTPADDING",  (0,0),(-1,-1), 3),
        ("TOPPADDING",    (0,1),(-1,-1), 2),
        ("BOTTOMPADDING", (0,1),(-1,-1), 2),
        ("BACKGROUND",    (0,n-1),(-1,n-1), RL_LIGHT),
        ("FONTNAME",      (0,n-1),(-1,n-1), "Helvetica-Bold"),
    ]))
    return t

# ══════════════════════════════════════════════════════════════════════════════
# 3. LOAD MATRICULA
# ══════════════════════════════════════════════════════════════════════════════
def cargar_matricula(path, mes_num):
    print(f"  [*] Cargando matrícula: {os.path.basename(path)}")
    df = pd.read_excel(path, dtype=str)
    df.columns = df.columns.str.strip().str.replace(r"\s+", " ", regex=True)
    df["GRADO"] = df["GRADO"].astype(str).str.strip()
    df = df[df["GRADO"].isin(GRADE_ORDER_LONG)].copy()
    df["GRADO_LONG"]  = df["GRADO"]
    df["GRADO_SHORT"] = df["GRADO_LONG"].map(LONG_TO_SHORT)
    if "GRUPO" in df.columns:
        df["GRUPO"] = df["GRUPO"].astype(str).str.strip().str.upper()
    else:
        df["GRUPO"] = None
    if "NIE" in df.columns:
        df["NIE_LIMPIO"] = (df["NIE"].astype(str)
                            .str.replace(r"\.0$","",regex=True).str.strip())
    else:
        df["NIE_LIMPIO"] = ""
    dep_col = next((c for c in df.columns
                    if "DEPART" in c.upper()
                    and "CÓDIGO" not in c.upper()
                    and "CODIGO" not in c.upper()), None)
    df["DEPARTAMENTO"] = df[dep_col].astype(str).str.strip() if dep_col else ""
    print(f"    → {len(df):,} estudiantes en grados válidos  |  "
          f"GRUPO disponible: {'GRUPO' in df.columns and df['GRUPO'].notna().any()}")
    return df

# ══════════════════════════════════════════════════════════════════════════════
# 4. LOAD CSVs
# ══════════════════════════════════════════════════════════════════════════════
def inferir_materia(nombre):
    n = nombre.upper()
    if n.startswith("MAT"): return "Matemática"
    if n.startswith("LEC") or n.startswith("LEN"): return "Lengua"
    return None

def cargar_csvs_mes(ruta):
    archivos = glob.glob(os.path.join(ruta, "*.csv"))
    if not archivos:
        print(f"  [!] Sin CSVs en: {ruta}")
        return pd.DataFrame(), pd.DataFrame()

    valid_list = []
    all_list   = []
    COL_SCORE  = "theta.global (escala 0-100)"

    for f in archivos:
        nombre  = os.path.basename(f)
        if "legend" in nombre.lower(): continue
        materia = inferir_materia(nombre)
        if materia is None: continue
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors="ignore")
            df.columns = df.columns.str.strip()
        except Exception as e:
            print(f"  [!] Error leyendo {nombre}: {e}"); continue

        if COL_SCORE not in df.columns: continue

        if "Grado" in df.columns:
            df["Grado"] = df["Grado"].apply(normalize_grade_csv)
        if "Nro de centro" in df.columns:
            df["Nro de centro"] = (df["Nro de centro"].astype(str)
                                   .str.replace(r"\.0$","",regex=True).str.strip())
        if "Documento" in df.columns:
            df["Documento"] = (df["Documento"].astype(str)
                               .str.replace(r"\.0$","",regex=True).str.strip())
        df["Materia"] = materia

        cols = [c for c in ["Documento","Grado","Nro de centro","Centro",
                            "Materia",COL_SCORE,"anular_prueba"] if c in df.columns]
        df = df[cols].copy()

        df_a = df[df["Nro de centro"].notna()].copy()
        all_list.append(df_a)

        df_v = df.copy()
        if "anular_prueba" in df_v.columns:
            df_v = df_v[~df_v["anular_prueba"].apply(excluir_por_tiempo)].copy()
        df_v[COL_SCORE] = pd.to_numeric(
            df_v[COL_SCORE].astype(str).str.replace(",","."), errors="coerce")
        df_v = df_v.dropna(subset=[COL_SCORE])
        if "Nro de centro" in df_v.columns:
            df_v = df_v[~df_v["Nro de centro"].isin(VIRTUAL_CODES)]
        if "Centro" in df_v.columns:
            df_v = df_v[~df_v["Centro"].astype(str).str.contains(
                "Centro Virtual", case=False, na=False)]
        if not df_v.empty:
            valid_list.append(df_v)

    df_valid = pd.concat(valid_list, ignore_index=True) if valid_list else pd.DataFrame()
    df_all   = pd.concat(all_list,   ignore_index=True) if all_list   else pd.DataFrame()

    print(f"    → Válidos (no virtual): {len(df_valid):,}  |  "
          f"Todos (incl. virtual): {len(df_all):,}")
    return df_valid, df_all

# ══════════════════════════════════════════════════════════════════════════════
# 5. COVERAGE TABLE
# ══════════════════════════════════════════════════════════════════════════════
def build_coverage(df_mat, df_csvs_valid):
    mat_counts = (df_mat.groupby("GRADO_SHORT").size()
                  .reindex(GRADE_ORDER_SHORT, fill_value=0)
                  .rename("MATRICULA"))
    rows = []
    for gs in GRADE_ORDER_SHORT:
        mat_n = int(mat_counts.get(gs, 0))
        mat_p = lec_p = 0
        if not df_csvs_valid.empty and "Grado" in df_csvs_valid.columns:
            sub   = df_csvs_valid[df_csvs_valid["Grado"] == gs]
            mat_p = int(sub[sub["Materia"]=="Matemática"]["Documento"].nunique()) \
                    if "Documento" in df_csvs_valid.columns else int(len(sub[sub["Materia"]=="Matemática"]))
            lec_p = int(sub[sub["Materia"]=="Lengua"]["Documento"].nunique()) \
                    if "Documento" in df_csvs_valid.columns else int(len(sub[sub["Materia"]=="Lengua"]))
        mat_pct = mat_p / mat_n * 100 if mat_n > 0 else 0.0
        lec_pct = lec_p / mat_n * 100 if mat_n > 0 else 0.0
        rows.append({"GRADO_SHORT":gs, "GRADO_DISPLAY":GRADE_DISPLAY[gs],
                     "MATRICULA":mat_n, "MAT_PART":mat_p, "MAT_PCT":mat_pct,
                     "LEC_PART":lec_p, "LEC_PCT":lec_pct})
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════════════════════
# 6. COVERAGE → PDF
# ══════════════════════════════════════════════════════════════════════════════
def coverage_to_pdf(df_cov, mes_label, path_out):
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("T", parent=styles["Title"],
        textColor=colors.white, backColor=RL_DARK, fontSize=13,
        spaceAfter=4, spaceBefore=4, alignment=TA_CENTER, leading=20)
    note_s = ParagraphStyle("N", parent=styles["Normal"],
        fontSize=7, textColor=colors.HexColor("#555555"),
        spaceBefore=4, spaceAfter=2, alignment=TA_LEFT)
    doc   = SimpleDocTemplate(path_out, pagesize=letter,
                              leftMargin=0.5*inch, rightMargin=0.5*inch,
                              topMargin=0.6*inch,  bottomMargin=0.6*inch)
    story = []
    W     = letter[0] - 1.0*inch
    story.append(Paragraph(f"Cobertura de Participación: Matemática y Lengua ({mes_label})", title_s))
    story.append(Spacer(1, 8))
    
    header_row = ["Grado", "Matrícula\nEsperada", "MAT\nParticipantes", "MAT\n% Cobertura", "LEC\nParticipantes", "LEC\n% Cobertura"]
    data = [header_row]
    for _, r in df_cov.iterrows():
        data.append([r["GRADO_DISPLAY"], f"{int(r['MATRICULA']):,}",
                     f"{int(r['MAT_PART']):,}", f"{r['MAT_PCT']:.1f}%",
                     f"{int(r['LEC_PART']):,}", f"{r['LEC_PCT']:.1f}%"])
    tot_mat  = int(df_cov['MATRICULA'].sum())
    tot_matp = int(df_cov['MAT_PART'].sum())
    tot_lecp = int(df_cov['LEC_PART'].sum())
    tot_matp_pct = tot_matp/tot_mat*100 if tot_mat>0 else 0
    tot_lecp_pct = tot_lecp/tot_mat*100 if tot_mat>0 else 0
    data.append(["Total General", f"{tot_mat:,}",
                 f"{tot_matp:,}", f"{tot_matp_pct:.1f}%",
                 f"{tot_lecp:,}", f"{tot_lecp_pct:.1f}%"])
    n_rows = len(data)
    cw = [W*0.22,W*0.13,W*0.16,W*0.13,W*0.16,W*0.13]
    t = Table(data, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),RL_DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,0),8),
        ("ALIGN",(0,0),(-1,0),"CENTER"),("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("BOTTOMPADDING",(0,0),(-1,0),5),("TOPPADDING",(0,0),(-1,0),5),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),("FONTSIZE",(0,1),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white,RL_GRAY]),
        ("ALIGN",(1,1),(-1,-1),"RIGHT"),("ALIGN",(0,1),(0,-1),"LEFT"),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,1),(-1,-1),2),("BOTTOMPADDING",(0,1),(-1,-1),2),
        ("BACKGROUND",(0,n_rows-1),(-1,n_rows-1),RL_LIGHT),
        ("FONTNAME",(0,n_rows-1),(-1,n_rows-1),"Helvetica-Bold"),
        ("FONTSIZE",(0,n_rows-1),(-1,n_rows-1),8),
    ]))
    story.append(t)
    story.append(Spacer(1,6))
    story.append(Paragraph(
        "Nota: Matrícula esperada según registro oficial. Solo población con pruebas válidas.", note_s))
    doc.build(story)
    print(f"    [OK] PDF Cobertura: {os.path.basename(path_out)}")

# ══════════════════════════════════════════════════════════════════════════════
# 7. COVERAGE → EXCEL 
# ══════════════════════════════════════════════════════════════════════════════
def coverage_to_excel(df_cov, mes_label, path_out):
    wb = Workbook(); ws = wb.active; ws.title = "Cobertura"
    DARK=XL_DARK_BLUE; LIGHT=XL_LIGHT_BLUE; GRAY=XL_GRAY

    headers = [
        "Grado", "Matrícula Esperada", 
        "MAT Participantes", "MAT % Cobertura", 
        "LEC Participantes", "LEC % Cobertura"
    ]
    for ci, h in enumerate(headers, start=1):
        xl_hdr(ws.cell(1, ci), h, bg=DARK, sz=10)
    ws.row_dimensions[1].height = 35

    DATA_START = 2
    for r,row in enumerate(df_cov.itertuples(),start=DATA_START):
        bg=XL_WHITE if r%2==0 else GRAY
        xl_data(ws.cell(r,1),row.GRADO_DISPLAY,"left",bg=bg)
        xl_data(ws.cell(r,2),int(row.MATRICULA),"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(r,3),int(row.MAT_PART),"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(r,4),f"{row.MAT_PCT:.1f}%","right",bg=bg)
        xl_data(ws.cell(r,5),int(row.LEC_PART),"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(r,6),f"{row.LEC_PCT:.1f}%","right",bg=bg)
        
    tot_r = DATA_START + len(df_cov)
    tot_mat=int(df_cov["MATRICULA"].sum()); tot_matp=int(df_cov["MAT_PART"].sum())
    tot_lecp=int(df_cov["LEC_PART"].sum())
    tot_matp_pct=tot_matp/tot_mat*100 if tot_mat>0 else 0.0
    tot_lecp_pct=tot_lecp/tot_mat*100 if tot_mat>0 else 0.0
    
    for ci,(val,aln,fmt) in enumerate([
        ("Total General","left",None),(tot_mat,"right","#,##0"),
        (tot_matp,"right","#,##0"),(f"{tot_matp_pct:.1f}%","right",None),
        (tot_lecp,"right","#,##0"),(f"{tot_lecp_pct:.1f}%","right",None),
    ],start=1):
        c=ws.cell(tot_r,ci); c.value=val
        c.font=Font(bold=True,size=10,name="Calibri")
        c.alignment=Alignment(horizontal=aln,vertical="center")
        c.fill=PatternFill("solid",fgColor=LIGHT); c.border=B_THIN
        if fmt: c.number_format=fmt

    note_r = tot_r + 1
    c = ws.cell(note_r, 1, value="Nota: Matricula esperada segun registro oficial. Solo poblacion con pruebas validas.")
    c.font = Font(italic=True, size=8, color="555555", name="Calibri")
    
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    for col in ["C","D","E","F"]: ws.column_dimensions[col].width=18
    ws.freeze_panes = "A2"
    wb.save(path_out)
    print(f"    [OK] Excel Cobertura: {os.path.basename(path_out)}")

# ══════════════════════════════════════════════════════════════════════════════
# 8–9. SCHOOL REPORTS
# ══════════════════════════════════════════════════════════════════════════════
def build_school_data(df_mat):
    df = df_mat.copy()
    if df["GRUPO"].isna().all(): df["GRUPO"]="SIN_GRUPO"
    df["GRUPO"] = df["GRUPO"].fillna("SIN_GRUPO").str.upper().str.strip()
    df = df[df["GRUPO"].isin(["B1","B2"])].copy()
    schools=(df.groupby(["GRUPO","CODIGO","NOMBRE"]).size().reset_index(name="Total")
             .sort_values(["GRUPO","NOMBRE"]))
    b1=schools[schools["GRUPO"]=="B1"].reset_index(drop=True)
    b2=schools[schools["GRUPO"]=="B2"].reset_index(drop=True)
    by_grade=(df.groupby("GRADO_LONG").size().reindex(GRADE_ORDER_LONG).fillna(0).astype(int))
    pivot=df.groupby(["GRADO_LONG","GRUPO"]).size().unstack(fill_value=0)
    pivot["TOTAL"]=pivot.sum(axis=1)
    pivot=pivot.reindex(GRADE_ORDER_LONG).fillna(0).astype(int)
    sgp=(df.groupby(["GRUPO","CODIGO","NOMBRE","GRADO_LONG"]).size().reset_index(name="n")
         .pivot_table(index=["GRUPO","CODIGO","NOMBRE"],columns="GRADO_LONG",
                      values="n",fill_value=0)
         .reindex(columns=GRADE_ORDER_LONG,fill_value=0).reset_index())
    sgp["TOTAL"]=sgp[GRADE_ORDER_LONG].sum(axis=1)
    sgp=sgp.sort_values(["GRUPO","NOMBRE"])
    total_b1=int(pivot["B1"].sum()) if "B1" in pivot else 0
    total_b2=int(pivot["B2"].sum()) if "B2" in pivot else 0
    return {"b1":b1,"b2":b2,"by_grade":by_grade,"pivot":pivot,"sgp":sgp,
            "sgp_b1":sgp[sgp["GRUPO"]=="B1"].reset_index(drop=True),
            "sgp_b2":sgp[sgp["GRUPO"]=="B2"].reset_index(drop=True),
            "total_b1":total_b1,"total_b2":total_b2,
            "total_schools":df["CODIGO"].nunique(),"grand_total":total_b1+total_b2}

def school_reports_to_excel(sd, mes_label, path_out):
    wb=Workbook(); wb.remove(wb.active)
    
    ws1=wb.create_sheet("Centros por Grupo")
    xl_hdr(ws1.cell(1,1),"Grupo"); xl_hdr(ws1.cell(1,2),"Centros Escolares")
    for r,(g,n) in enumerate([("B1",len(sd["b1"])),("B2",len(sd["b2"]))],start=2):
        bg=XL_WHITE if r%2==0 else XL_GRAY
        xl_data(ws1.cell(r,1),g,"left",bg=bg); xl_data(ws1.cell(r,2),n,fmt="#,##0",bg=bg)
    xl_total_row(ws1,4,["TOTAL",sd["total_schools"]])
    xl_auto_width(ws1); ws1.freeze_panes="A2"
    
    ws2=wb.create_sheet("Estudiantes por Grado")
    xl_hdr(ws2.cell(1,1),"Grado"); xl_hdr(ws2.cell(1,2),"Total de Estudiantes")
    for r,g in enumerate(GRADE_ORDER_LONG,start=2):
        bg=XL_WHITE if r%2==0 else XL_GRAY
        xl_data(ws2.cell(r,1),g,"left",bg=bg)
        xl_data(ws2.cell(r,2),int(sd["by_grade"][g]),fmt="#,##0",bg=bg)
    xl_total_row(ws2,r+1,["GRAN TOTAL",sd["grand_total"]])
    xl_auto_width(ws2); ws2.freeze_panes="A2"
    
    ws3=wb.create_sheet("Grado x Grupo")
    for ci,h in enumerate(["Grado","Grupo B1","Grupo B2","Total"],start=1): xl_hdr(ws3.cell(1,ci),h)
    for r,g in enumerate(GRADE_ORDER_LONG,start=2):
        bg=XL_WHITE if r%2==0 else XL_GRAY
        b1c=int(sd["pivot"].loc[g,"B1"]) if "B1" in sd["pivot"].columns else 0
        b2c=int(sd["pivot"].loc[g,"B2"]) if "B2" in sd["pivot"].columns else 0
        xl_data(ws3.cell(r,1),g,"left",bg=bg)
        for ci,v in enumerate([b1c,b2c,int(sd["pivot"].loc[g,"TOTAL"])],start=2):
            xl_data(ws3.cell(r,ci),v,fmt="#,##0",bg=bg)
    xl_total_row(ws3,r+1,["GRAN TOTAL",sd["total_b1"],sd["total_b2"],sd["grand_total"]])
    xl_auto_width(ws3); ws3.freeze_panes="A2"
    
    def write_sgp_sheet(wb,name,sgp_df,grupo_label,total_est):
        ws=wb.create_sheet(name); n_cols=3+len(GRADE_ORDER_LONG)+1
        headers=["#","Código","Nombre"]+[f"Grado {g}" for g in SHORT_GRADES_ABBREV]+["Total"]
        for ci,h in enumerate(headers,start=1): xl_hdr(ws.cell(1,ci),h,sz=9,wrap=True)
        ws.row_dimensions[1].height=26
        for r,row in sgp_df.iterrows():
            bg=XL_WHITE if r%2==1 else XL_GRAY
            xl_data(ws.cell(r+2,1),r+1,"center",bg=bg)
            xl_data(ws.cell(r+2,2),str(row["CODIGO"]),"left",bg=bg)
            xl_data(ws.cell(r+2,3),str(row["NOMBRE"]).strip(),"left",bg=bg)
            for gi,g in enumerate(GRADE_ORDER_LONG,start=4):
                v=int(row[g]); xl_data(ws.cell(r+2,gi),v if v>0 else None,fmt="#,##0",bg=bg)
            xl_data(ws.cell(r+2,n_cols),int(row["TOTAL"]),bold=True,fmt="#,##0",bg=bg)
        tot_r=len(sgp_df)+2
        gt=[int(sgp_df[g].sum()) for g in GRADE_ORDER_LONG]
        vals=["","",f"TOTAL {grupo_label}"]+gt+[total_est]
        alns=["center","left","left"]+["right"]*(len(GRADE_ORDER_LONG)+1)
        for ci,(v,a) in enumerate(zip(vals,alns),start=1):
            xl_data(ws.cell(tot_r,ci),v if v!=0 else None,a,bold=True,bg=XL_LIGHT_BLUE,
                    fmt="#,##0" if isinstance(v,int) and v!=0 else None)
        ws.column_dimensions["A"].width=5; ws.column_dimensions["B"].width=9
        ws.column_dimensions["C"].width=38
        for gi in range(4,4+len(GRADE_ORDER_LONG)):
            ws.column_dimensions[get_column_letter(gi)].width=10
        ws.column_dimensions[get_column_letter(n_cols)].width=10
        ws.freeze_panes="A2"
        
    write_sgp_sheet(wb,"Centros B1 por Grado",sd["sgp_b1"],"B1",sd["total_b1"])
    write_sgp_sheet(wb,"Centros B2 por Grado",sd["sgp_b2"],"B2",sd["total_b2"])
    
    ws6=wb.create_sheet("Resumen General")
    xl_hdr(ws6.cell(1,1),"Métrica"); xl_hdr(ws6.cell(1,2),"Valor")
    for r,(lbl,v) in enumerate([("Total de Escuelas",sd["total_schools"]),
        ("Total de Estudiantes",sd["grand_total"]),
        ("Total de Estudiantes B1",sd["total_b1"]),
        ("Total de Estudiantes B2",sd["total_b2"])],start=2):
        bg=XL_WHITE if r%2==0 else XL_GRAY
        xl_data(ws6.cell(r,1),lbl,"left",bg=bg); xl_data(ws6.cell(r,2),v,fmt="#,##0",bg=bg)
    xl_auto_width(ws6)
    
    wb.save(path_out)
    print(f"    [OK] Excel Escuelas: {os.path.basename(path_out)}")

def school_reports_to_pdf(sd, mes_label, path_out):
    styles=getSampleStyleSheet()
    title_s=ParagraphStyle("T",parent=styles["Title"],textColor=colors.white,
        backColor=RL_DARK,fontSize=13,spaceAfter=4,spaceBefore=4,alignment=TA_CENTER,leading=20)
    h2_s=ParagraphStyle("H2",parent=styles["Heading2"],textColor=colors.white,
        backColor=RL_MID,fontSize=10,spaceAfter=4,spaceBefore=8,alignment=TA_LEFT,leading=15)
    note_s=ParagraphStyle("N",parent=styles["Normal"],fontSize=7,
        textColor=colors.HexColor("#555555"),spaceBefore=2,spaceAfter=2)
    doc=SimpleDocTemplate(path_out,pagesize=letter,leftMargin=0.5*inch,
        rightMargin=0.5*inch,topMargin=0.6*inch,bottomMargin=0.6*inch)
    W=letter[0]-1.0*inch; story=[]
    story.append(Paragraph(f"MATRÍCULA — {mes_label}",title_s)); story.append(Spacer(1,8))
    story.append(Paragraph("NÚMERO DE CENTROS ESCOLARES POR GRUPO",h2_s))
    d=[["Grupo","Centros Escolares"],["B1",f"{len(sd['b1']):,}"],
       ["B2",f"{len(sd['b2']):,}"],["TOTAL",f"{sd['total_schools']:,}"]]
    story.append(rl_make_table(d,[W*0.5,W*0.5]))
    story.append(Paragraph("TOTAL DE ESTUDIANTES POR GRADO",h2_s))
    d=[["Grado","Total de Estudiantes"]]
    for g in GRADE_ORDER_LONG: d.append([g,f"{int(sd['by_grade'][g]):,}"])
    d.append(["GRAN TOTAL",f"{sd['grand_total']:,}"])
    story.append(rl_make_table(d,[W*0.6,W*0.4]))
    story.append(Paragraph("ESTUDIANTES POR GRADO Y GRUPO (B1 / B2)",h2_s))
    d=[["Grado","B1","B2","TOTAL"]]
    for g in GRADE_ORDER_LONG:
        b1c=int(sd["pivot"].loc[g,"B1"]) if "B1" in sd["pivot"].columns else 0
        b2c=int(sd["pivot"].loc[g,"B2"]) if "B2" in sd["pivot"].columns else 0
        d.append([g,f"{b1c:,}",f"{b2c:,}",f"{b1c+b2c:,}"])
    d.append(["GRAN TOTAL",f"{sd['total_b1']:,}",f"{sd['total_b2']:,}",f"{sd['grand_total']:,}"])
    story.append(rl_make_table(d,[W*0.4,W*0.2,W*0.2,W*0.2]))
    for grupo_label,sgp_df,total_est in [("B1",sd["sgp_b1"],sd["total_b1"]),("B2",sd["sgp_b2"],sd["total_b2"])]:
        story.append(PageBreak())
        story.append(Paragraph(f"LISTADO CENTROS ESCOLARES — GRUPO {grupo_label} (detalle por grado)",h2_s))
        story.append(Paragraph(
            "2°=Segundo, 3°=Tercer, 4°=Cuarto, 5°=Quinto, 6°=Sexto, "
            "7°=Séptimo, 8°=Octavo, 9°=Noveno, 1°B=Primer Año Bach., 2°B=Segundo Año Bach.  |  — = sin estudiantes",note_s))
        header=["#","Código","Nombre"]+SHORT_GRADES_ABBREV+["Total"]
        d=[header]
        for i,row in sgp_df.iterrows():
            d.append([str(i+1),str(row["CODIGO"]),row["NOMBRE"].strip()]
                     +[str(int(row[g])) if int(row[g])>0 else "—" for g in GRADE_ORDER_LONG]
                     +[f"{int(row['TOTAL']):,}"])
        gt=[int(sgp_df[g].sum()) for g in GRADE_ORDER_LONG]
        d.append(["","","TOTAL"]+[f"{v:,}" for v in gt]+[f"{total_est:,}"])
        n_gc=len(GRADE_ORDER_LONG); w_num=W*0.03; w_cod=W*0.07; w_nam=W*0.25; w_tot=W*0.08
        w_grd=(W-w_num-w_cod-w_nam-w_tot)/n_gc; cw=[w_num,w_cod,w_nam]+[w_grd]*n_gc+[w_tot]
        t=Table(d,colWidths=cw,repeatRows=1); nr=len(d)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),RL_DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),6),
            ("ALIGN",(0,0),(-1,0),"CENTER"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white,RL_GRAY]),
            ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#CCCCCC")),
            ("ALIGN",(3,1),(-1,-1),"RIGHT"),("ALIGN",(2,1),(2,-1),"LEFT"),
            ("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1),
            ("BACKGROUND",(0,nr-1),(-1,nr-1),RL_LIGHT),
            ("FONTNAME",(0,nr-1),(-1,nr-1),"Helvetica-Bold"),
        ]))
        story.append(t)
    story.append(PageBreak())
    story.append(Paragraph("RESUMEN GENERAL",h2_s))
    d=[["Métrica","Valor"],
       ["Total de Escuelas",f"{sd['total_schools']:,}"],
       ["Total de Estudiantes",f"{sd['grand_total']:,}"],
       ["Total de Estudiantes B1",f"{sd['total_b1']:,}"],
       ["Total de Estudiantes B2",f"{sd['total_b2']:,}"]]
    story.append(rl_make_table(d,[W*0.7,W*0.3]))
    doc.build(story)
    print(f"    [OK] PDF Escuelas: {os.path.basename(path_out)}")

# ══════════════════════════════════════════════════════════════════════════════
# 10. PARTICIPATION BY SCHOOL
# ══════════════════════════════════════════════════════════════════════════════
def participacion_to_excel(df_csvs_valid, mes_label, path_out):
    if df_csvs_valid.empty:
        print(f"    [!] Sin datos CSV — omitiendo participación por escuela.")
        return
    wb=Workbook(); wb.remove(wb.active)
    for prefix,label in [("LEC","Lengua (LEC)"),("MAT","Matemática (MAT)")]:
        df_sub=df_csvs_valid[df_csvs_valid["Materia"]==("Lengua" if prefix=="LEC" else "Matemática")].copy()
        if df_sub.empty: continue
        pivot=(df_sub.groupby(["Nro de centro","Centro","Grado"])
               .size().reset_index(name="n")
               .pivot_table(index=["Nro de centro","Centro"],columns="Grado",
                            values="n",aggfunc="sum",fill_value=0).reset_index())
        pivot.columns.name=None
        known=[g for g in GRADE_ORDER_SHORT if g in pivot.columns]
        unknown=sorted([g for g in pivot.columns
                        if g not in GRADE_ORDER_SHORT and g not in ("Nro de centro","Centro")])
        grade_cols=known+unknown
        pivot=pivot[["Nro de centro","Centro"]+grade_cols]
        pivot["Total General"]=pivot[grade_cols].sum(axis=1)
        ws=wb.create_sheet(label); all_cols=["Nro de centro","Centro"]+grade_cols+["Total General"]
        n_cols=len(all_cols)
        
        for ci,col in enumerate(all_cols,start=1): 
            h_text = f"Grado {col}" if col not in ("Nro de centro","Centro","Total General") else col
            xl_hdr(ws.cell(1,ci), h_text, sz=10)
        ws.row_dimensions[1].height=26
        
        ALT=PatternFill("solid",fgColor="EBF5FB")
        for ri,(_,row) in enumerate(pivot.iterrows(),start=2):
            row_fill=ALT if ri%2==0 else None
            for ci,col in enumerate(all_cols,start=1):
                c=ws.cell(ri,ci,value=row[col]); c.border=B_THIN
                c.alignment=Alignment(horizontal="left" if ci==2 else "center",vertical="center")
                c.font=Font(size=10,name="Arial")
                if ci==n_cols:
                    c.font=Font(bold=True,size=10,name="Arial")
                    c.fill=PatternFill("solid",fgColor=XL_LIGHT_BLUE)
                elif row_fill: c.fill=row_fill
                
        tot_r=len(pivot)+2
        ws.cell(tot_r,1,value="TOTAL GENERAL").font=Font(bold=True,size=10,name="Arial")
        ws.cell(tot_r,1).alignment=Alignment(horizontal="center",vertical="center")
        ws.cell(tot_r,2).fill=PatternFill("solid",fgColor=XL_LIGHT_BLUE)
        for ci in range(3,n_cols+1):
            cl=get_column_letter(ci)
            c=ws.cell(tot_r,ci,value=f"=SUM({cl}2:{cl}{tot_r-1})")
            c.font=Font(bold=True,size=10,name="Arial")
            c.fill=PatternFill("solid",fgColor=XL_LIGHT_BLUE)
            c.alignment=Alignment(horizontal="center",vertical="center"); c.border=B_MED
            
        ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=46
        for ci in range(3,n_cols+1): ws.column_dimensions[get_column_letter(ci)].width=13
        ws.freeze_panes="C2"
    wb.save(path_out)
    print(f"    [OK] Excel Participación: {os.path.basename(path_out)}")

# ══════════════════════════════════════════════════════════════════════════════
# 11. MASTER COVERAGE EXCEL 
# ══════════════════════════════════════════════════════════════════════════════
def coverage_master_to_excel(coberturas_por_mes, path_out):
    if not coberturas_por_mes: return
    DARK=XL_DARK_BLUE; LIGHT=XL_LIGHT_BLUE; GRAY=XL_GRAY
    
    wb=Workbook(); wb.remove(wb.active)
    for mes_label,df_cov in coberturas_por_mes:
        ws=wb.create_sheet(title=mes_label[:31])
        
        headers = [
            "Grado", "Matrícula Esperada", 
            "MAT Participantes", "MAT % Cobertura", 
            "LEC Participantes", "LEC % Cobertura"
        ]
        for ci, h in enumerate(headers, start=1):
            xl_hdr(ws.cell(1, ci), h, bg=DARK, sz=10)
        ws.row_dimensions[1].height = 35

        DATA_START=2
        for r,row in enumerate(df_cov.itertuples(),start=DATA_START):
            bg=XL_WHITE if r%2==0 else GRAY
            xl_data(ws.cell(r,1),row.GRADO_DISPLAY,"left",bg=bg)
            xl_data(ws.cell(r,2),int(row.MATRICULA),"right",bg=bg,fmt="#,##0")
            xl_data(ws.cell(r,3),int(row.MAT_PART),"right",bg=bg,fmt="#,##0")
            xl_data(ws.cell(r,4),f"{row.MAT_PCT:.1f}%","right",bg=bg)
            xl_data(ws.cell(r,5),int(row.LEC_PART),"right",bg=bg,fmt="#,##0")
            xl_data(ws.cell(r,6),f"{row.LEC_PCT:.1f}%","right",bg=bg)
            
        tot_r=DATA_START+len(df_cov)
        tot_mat=int(df_cov["MATRICULA"].sum()); tot_matp=int(df_cov["MAT_PART"].sum())
        tot_lecp=int(df_cov["LEC_PART"].sum())
        tot_matp_pct=tot_matp/tot_mat*100 if tot_mat>0 else 0.0
        tot_lecp_pct=tot_lecp/tot_mat*100 if tot_mat>0 else 0.0
        
        for ci,(val,aln,fmt) in enumerate([
            ("Total General","left",None),(tot_mat,"right","#,##0"),
            (tot_matp,"right","#,##0"),(f"{tot_matp_pct:.1f}%","right",None),
            (tot_lecp,"right","#,##0"),(f"{tot_lecp_pct:.1f}%","right",None),
        ],start=1):
            c=ws.cell(tot_r,ci); c.value=val
            c.font=Font(bold=True,size=10,name="Calibri")
            c.alignment=Alignment(horizontal=aln,vertical="center")
            c.fill=PatternFill("solid",fgColor=LIGHT); c.border=B_THIN
            if fmt: c.number_format=fmt
            
        note_r=tot_r+1
        c=ws.cell(note_r,1,value="Nota: Matricula esperada segun registro oficial. Solo poblacion con pruebas validas.")
        c.font=Font(italic=True,size=8,color="555555",name="Calibri")
        
        ws.column_dimensions["A"].width=25; ws.column_dimensions["B"].width=18
        for col in ["C","D","E","F"]: ws.column_dimensions[col].width=18
        ws.freeze_panes="A2"
    wb.save(path_out)
    print(f"  [OK] Excel Maestro Cobertura ({len(coberturas_por_mes)} hojas): {os.path.basename(path_out)}")

# ══════════════════════════════════════════════════════════════════════════════
# 12. PER-SCHOOL PARTICIPATION MASTER EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def build_school_participation(df_mat, df_csvs_valid, df_csvs_all):
    enroll = (df_mat.groupby(["CODIGO","NOMBRE"])
              .size().reset_index(name="MATRICULA"))
    enroll["Nro de centro"] = enroll["CODIGO"].astype(str).str.strip()

    def count_unique(df_sub, subject):
        if df_sub.empty or "Nro de centro" not in df_sub.columns:
            return pd.DataFrame(columns=["Nro de centro","n"])
        mask = df_sub["Materia"] == ("Matemática" if subject=="MAT" else "Lengua")
        g = (df_sub[mask]
             .groupby("Nro de centro")["Documento"]
             .nunique().reset_index(name="n"))
        return g

    v_mat = count_unique(df_csvs_valid, "MAT").rename(columns={"n":"MAT_VALID"})
    v_lec = count_unique(df_csvs_valid, "LEC").rename(columns={"n":"LEC_VALID"})

    df_all_nv = df_csvs_all[~df_csvs_all.get("Nro de centro",
                pd.Series(dtype=str)).isin(VIRTUAL_CODES)].copy() \
                if not df_csvs_all.empty else df_csvs_all
    a_mat = count_unique(df_all_nv, "MAT").rename(columns={"n":"MAT_ALL"})
    a_lec = count_unique(df_all_nv, "LEC").rename(columns={"n":"LEC_ALL"})

    low_grades = pd.DataFrame(columns=["Nro de centro",
                                        "Menor_Partic_1","Menor_Partic_2",
                                        "Menor_Partic_3","Menor_Partic_4"])
    if not df_all_nv.empty and "Grado" in df_all_nv.columns and "Nro de centro" in df_all_nv.columns:
        gp = (df_all_nv
              .groupby(["Nro de centro","Grado"])["Documento"]
              .nunique().reset_index(name="n"))

        def lowest4(grp):
            s = grp.nsmallest(4, "n")["Grado"].tolist()
            s += [""] * 4
            return pd.Series({f"Menor_Partic_{j+1}": s[j] for j in range(4)})

        low_grades = gp.groupby("Nro de centro").apply(lowest4).reset_index()

    df = enroll.copy()
    for part_df in [v_mat, v_lec, a_mat, a_lec]:
        if not part_df.empty:
            df = df.merge(part_df, on="Nro de centro", how="left")
    if not low_grades.empty:
        df = df.merge(low_grades, on="Nro de centro", how="left")

    for col in ["MAT_VALID","LEC_VALID","MAT_ALL","LEC_ALL"]:
        if col not in df.columns: df[col] = 0
        df[col] = df[col].fillna(0).astype(int)

    df["MAT_VALID_PCT"] = df.apply(lambda r: r["MAT_VALID"]/r["MATRICULA"]*100 if r["MATRICULA"]>0 else 0, axis=1)
    df["LEC_VALID_PCT"] = df.apply(lambda r: r["LEC_VALID"]/r["MATRICULA"]*100 if r["MATRICULA"]>0 else 0, axis=1)
    df["MAT_ALL_PCT"]   = df.apply(lambda r: r["MAT_ALL"]  /r["MATRICULA"]*100 if r["MATRICULA"]>0 else 0, axis=1)
    df["LEC_ALL_PCT"]   = df.apply(lambda r: r["LEC_ALL"]  /r["MATRICULA"]*100 if r["MATRICULA"]>0 else 0, axis=1)

    return df.sort_values("NOMBRE").reset_index(drop=True)

def school_participation_sheet(wb, mes_label, df_sp):
    ws = wb.create_sheet(title=mes_label[:31])
    DARK=XL_DARK_BLUE; LIGHT=XL_LIGHT_BLUE; GRAY=XL_GRAY

    flat_headers = [
        "Nro de Centro", "Centro Educativo", "Matrícula Esperada",
        "MAT Válidos (Partic.)", "MAT Válidos (% Cob.)",
        "LEC Válidos (Partic.)", "LEC Válidos (% Cob.)",
        "MAT Total (Partic.)", "MAT Total (% Cob.)",
        "LEC Total (Partic.)", "LEC Total (% Cob.)",
        "Menor Partic. Grado 1", "Menor Partic. Grado 2",
        "Menor Partic. Grado 3", "Menor Partic. Grado 4"
    ]
    for ci, h in enumerate(flat_headers, start=1):
        xl_hdr(ws.cell(1, ci), h, bg=DARK, sz=9)
    ws.row_dimensions[1].height = 40

    low_cols = [c for c in df_sp.columns if c.startswith("Menor_Partic_")]

    for ri,row in enumerate(df_sp.itertuples(),start=2):
        bg = XL_WHITE if ri%2==0 else GRAY
        xl_data(ws.cell(ri,1),  str(row.CODIGO),            "center",bg=bg)
        xl_data(ws.cell(ri,2),  str(row.NOMBRE).strip(),    "left",  bg=bg)
        xl_data(ws.cell(ri,3),  int(row.MATRICULA),         "right", bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,4),  int(row.MAT_VALID),         "right", bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,5),  f"{row.MAT_VALID_PCT:.1f}%", "right", bg=bg)
        xl_data(ws.cell(ri,6),  int(row.LEC_VALID),         "right", bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,7),  f"{row.LEC_VALID_PCT:.1f}%", "right", bg=bg)
        xl_data(ws.cell(ri,8),  int(row.MAT_ALL),           "right", bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,9),  f"{row.MAT_ALL_PCT:.1f}%",   "right", bg=bg)
        xl_data(ws.cell(ri,10), int(row.LEC_ALL),           "right", bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,11), f"{row.LEC_ALL_PCT:.1f}%",   "right", bg=bg)
        for j,lc in enumerate(low_cols[:4]):
            val=getattr(row,lc,"") if lc in df_sp.columns else ""
            xl_data(ws.cell(ri,12+j),str(val) if val else "","center",bg=bg)

    tot_r = 2 + len(df_sp)
    tot_enroll = int(df_sp["MATRICULA"].sum())
    def pct_s(num,den): return f"{num/den*100:.1f}%" if den>0 else "—"
    for ci,(val,aln,fmt) in enumerate([
        ("TOTAL GENERAL","left",None),("","left",None),(tot_enroll,"right","#,##0"),
        (int(df_sp["MAT_VALID"].sum()),"right","#,##0"),
        (pct_s(df_sp["MAT_VALID"].sum(),tot_enroll),"right",None),
        (int(df_sp["LEC_VALID"].sum()),"right","#,##0"),
        (pct_s(df_sp["LEC_VALID"].sum(),tot_enroll),"right",None),
        (int(df_sp["MAT_ALL"].sum()),"right","#,##0"),
        (pct_s(df_sp["MAT_ALL"].sum(),tot_enroll),"right",None),
        (int(df_sp["LEC_ALL"].sum()),"right","#,##0"),
        (pct_s(df_sp["LEC_ALL"].sum(),tot_enroll),"right",None),
        ("","center",None),("","center",None),("","center",None),("","center",None),
    ],start=1):
        c=ws.cell(tot_r,ci); c.value=val
        c.font=Font(bold=True,size=9,name="Calibri")
        c.alignment=Alignment(horizontal=aln,vertical="center")
        c.fill=PatternFill("solid",fgColor=LIGHT); c.border=B_THIN
        if fmt: c.number_format=fmt

    ws.column_dimensions["A"].width=14
    ws.column_dimensions["B"].width=44
    ws.column_dimensions["C"].width=14
    for ci in range(4,12): ws.column_dimensions[get_column_letter(ci)].width=14
    for ci in range(12,16): ws.column_dimensions[get_column_letter(ci)].width=14
    ws.freeze_panes="C2"

# ══════════════════════════════════════════════════════════════════════════════
# 13. CENTRO VIRTUAL SHEET
# ══════════════════════════════════════════════════════════════════════════════
def virtual_sheet(wb, virtual_data):
    ws = wb.create_sheet(title="Centro Virtual")
    DARK=XL_DARK_BLUE; GRAY=XL_GRAY

    flat_headers = [
        "Mes", "MAT Virtual", "MAT Total Partic.", "MAT % Virtual",
        "LEC Virtual", "LEC Total Partic.", "LEC % Virtual",
        "Combinado Virtual", "Combinado % Virtual"
    ]
    for ci, h in enumerate(flat_headers, start=1):
        xl_hdr(ws.cell(1, ci), h, bg=DARK, sz=9)
    ws.row_dimensions[1].height = 25

    for ri,d in enumerate(virtual_data,start=2):
        bg=XL_WHITE if ri%2==0 else GRAY
        mat_pct=d["mat_virtual"]/d["mat_total"]*100 if d["mat_total"]>0 else 0
        lec_pct=d["lec_virtual"]/d["lec_total"]*100 if d["lec_total"]>0 else 0
        tot_v  =d["mat_virtual"]+d["lec_virtual"]
        tot_t  =d["mat_total"]+d["lec_total"]
        tot_pct=tot_v/tot_t*100 if tot_t>0 else 0
        
        xl_data(ws.cell(ri,1),d["label"],"left",bg=bg)
        xl_data(ws.cell(ri,2),d["mat_virtual"],"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,3),d["mat_total"],  "right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,4),f"{mat_pct:.2f}%","right",bg=bg)
        xl_data(ws.cell(ri,5),d["lec_virtual"],"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,6),d["lec_total"],  "right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,7),f"{lec_pct:.2f}%","right",bg=bg)
        xl_data(ws.cell(ri,8),tot_v,"right",bg=bg,fmt="#,##0")
        xl_data(ws.cell(ri,9),f"{tot_pct:.2f}%","right",bg=bg)
        
    note_r = len(virtual_data) + 2
    c=ws.cell(note_r, 1, value=(
        "Nota: % calculado respecto al total de participantes (incluyendo Centro Virtual) "
        "en cada mes. Conteo por Documento único por materia."))
    c.font=Font(italic=True,size=8,color="555555",name="Calibri")

    ws.column_dimensions["A"].width=18
    for ci in range(2,10): ws.column_dimensions[get_column_letter(ci)].width=18
    ws.freeze_panes="B2"

# ══════════════════════════════════════════════════════════════════════════════
# 14. GRADE × MONTH SHEET (UPDATED WITH PERCENTAGES)
# ══════════════════════════════════════════════════════════════════════════════
def grade_by_month_sheet(wb, grade_month_data):
    ws = wb.create_sheet(title="Grado x Mes (todos)")
    DARK=XL_DARK_BLUE; LIGHT=XL_LIGHT_BLUE; GRAY=XL_GRAY

    headers = ["Grado"]
    for mes_label, _, _ in grade_month_data:
        headers.append(f"{mes_label} MAT (todos)")
        headers.append(f"{mes_label} MAT (%)")
        headers.append(f"{mes_label} LEC (todos)")
        headers.append(f"{mes_label} LEC (%)")

    for ci, h in enumerate(headers, start=1):
        xl_hdr(ws.cell(1, ci), h, bg=DARK, sz=9)
    ws.row_dimensions[1].height = 40

    month_counts = []
    for mes_label, df_all, df_mat in grade_month_data:
        counts = {}
        # Get enrollment for this specific month for accurate percentage calculations
        mat_counts = df_mat.groupby("GRADO_SHORT").size().to_dict() if df_mat is not None and not df_mat.empty else {}
        
        if not df_all.empty and "Grado" in df_all.columns and "Documento" in df_all.columns:
            for gs in GRADE_ORDER_SHORT:
                sub  = df_all[df_all["Grado"]==gs]
                mat_n = int(sub[sub["Materia"]=="Matemática"]["Documento"].nunique())
                lec_n = int(sub[sub["Materia"]=="Lengua"]["Documento"].nunique())
                enroll = mat_counts.get(gs, 0)
                counts[gs] = (mat_n, lec_n, enroll)
        else:
            for gs in GRADE_ORDER_SHORT: 
                counts[gs] = (0, 0, mat_counts.get(gs, 0))
        month_counts.append(counts)

    for ri, gs in enumerate(GRADE_ORDER_SHORT, start=2):
        bg = XL_WHITE if ri%2==0 else GRAY
        xl_data(ws.cell(ri, 1), GRADE_DISPLAY.get(gs, gs), "left", bg=bg)
        
        col_idx = 2
        for mi, counts in enumerate(month_counts):
            mat_n, lec_n, enroll = counts.get(gs, (0, 0, 0))
            mat_pct = (mat_n / enroll * 100) if enroll > 0 else 0.0
            lec_pct = (lec_n / enroll * 100) if enroll > 0 else 0.0
            
            xl_data(ws.cell(ri, col_idx),   mat_n, "right", bg=bg, fmt="#,##0")
            xl_data(ws.cell(ri, col_idx+1), f"{mat_pct:.1f}%", "right", bg=bg)
            xl_data(ws.cell(ri, col_idx+2), lec_n, "right", bg=bg, fmt="#,##0")
            xl_data(ws.cell(ri, col_idx+3), f"{lec_pct:.1f}%", "right", bg=bg)
            col_idx += 4

    tot_r = 2 + len(GRADE_ORDER_SHORT)
    xl_data(ws.cell(tot_r, 1), "TOTAL GENERAL", "left", bold=True, bg=LIGHT)
    
    col_idx = 2
    for mi, counts in enumerate(month_counts):
        tot_mat = sum(v[0] for v in counts.values())
        tot_lec = sum(v[1] for v in counts.values())
        tot_enroll = sum(v[2] for v in counts.values())
        
        tot_mat_pct = (tot_mat / tot_enroll * 100) if tot_enroll > 0 else 0.0
        tot_lec_pct = (tot_lec / tot_enroll * 100) if tot_enroll > 0 else 0.0
        
        for ci, val, is_pct in [
            (col_idx,   tot_mat, False),
            (col_idx+1, f"{tot_mat_pct:.1f}%", True),
            (col_idx+2, tot_lec, False),
            (col_idx+3, f"{tot_lec_pct:.1f}%", True)
        ]:
            c = ws.cell(tot_r, ci)
            c.value = val
            c.font = Font(bold=True, size=9, name="Calibri")
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.fill = PatternFill("solid", fgColor=LIGHT)
            c.border = B_THIN
            if not is_pct:
                c.number_format = "#,##0"
        col_idx += 4

    ws.column_dimensions["A"].width = 24
    for ci in range(2, len(headers)+1): 
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.freeze_panes = "B2"

# ══════════════════════════════════════════════════════════════════════════════
# 15. MAIN LOOP
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print(f"\n{'='*65}")
    print("  COBERTURA Y MATRÍCULA POR MES — PRUEBA DE PROGRESO 2026")
    print(f"{'='*65}\n")

    coberturas_por_mes    = [] 
    school_part_by_mes    = [] 
    virtual_data          = [] 
    grade_month_data      = [] 

    for i, mes in enumerate(MESES, start=1):
        label = mes["label"]; short = mes["short"]
        print(f"\n{'─'*65}\n  PROCESANDO: {label}\n{'─'*65}")

        out_dir = os.path.join(PATH_OUTPUT_BASE, short)
        os.makedirs(out_dir, exist_ok=True)

        mat_path = mes["matricula"]
        if not os.path.exists(mat_path):
            local = os.path.join(
                r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System\Pedidos_Frecuentes_No_Mensual\13. Matricula Progreso Mes 1-Mes 4",
                f"Matricula_P{i}.xlsx")
            mat_path = local if os.path.exists(local) else mat_path
        if not os.path.exists(mat_path):
            print(f"  [!] Matrícula no encontrada: {mat_path} — saltando mes."); continue

        df_mat = cargar_matricula(mat_path, i)
        if df_mat["GRUPO"].isna().all():
            df_mat["GRUPO"] = "B1"
            print(f"    → GRUPO = B1 (archivo exclusivo B1): {len(df_mat):,} estudiantes")

        csv_dir = mes["csvs"]
        if not os.path.exists(csv_dir):
            print(f"  [!] Directorio CSVs no encontrado: {csv_dir}")
            df_csvs_valid = pd.DataFrame(); df_csvs_all = pd.DataFrame()
        else:
            df_csvs_valid, df_csvs_all = cargar_csvs_mes(csv_dir)

        print(f"  [*] Generando tabla de cobertura...")
        df_cov = build_coverage(df_mat, df_csvs_valid)
        coberturas_por_mes.append((label, df_cov))
        coverage_to_pdf(df_cov, label,
                        os.path.join(PATH_OUTPUT_COVERAGE, f"Cobertura_{short}.pdf"))
        coverage_to_excel(df_cov, label,
                          os.path.join(PATH_OUTPUT_COVERAGE, f"Cobertura_{short}.xlsx"))

        print(f"  [*] Generando reportes de escuelas...")
        sd = build_school_data(df_mat)
        school_reports_to_excel(sd, label,
            os.path.join(out_dir, f"Resumen_Escuelas_{short}.xlsx"))
        school_reports_to_pdf(sd, label,
            os.path.join(out_dir, f"Resumen_Escuelas_{short}.pdf"))

        print(f"  [*] Generando participación por escuela...")
        participacion_to_excel(df_csvs_valid, label,
            os.path.join(out_dir, f"Participacion_Escuelas_{short}.xlsx"))

        df_sp = build_school_participation(df_mat, df_csvs_valid, df_csvs_all)
        school_part_by_mes.append((label, df_sp))

        mat_virt = lec_virt = mat_tot = lec_tot = 0
        if not df_csvs_all.empty and "Nro de centro" in df_csvs_all.columns:
            virt_mask = df_csvs_all["Nro de centro"].isin(VIRTUAL_CODES)
            for subject, subject_label in [("Matemática","mat"), ("Lengua","lec")]:
                sub = df_csvs_all[df_csvs_all["Materia"]==subject]
                v_n = int(sub[virt_mask[sub.index]]["Documento"].nunique()) \
                      if "Documento" in sub.columns else int(len(sub[virt_mask[sub.index]]))
                t_n = int(sub["Documento"].nunique()) \
                      if "Documento" in sub.columns else int(len(sub))
                if subject_label == "mat": mat_virt=v_n; mat_tot=t_n
                else: lec_virt=v_n; lec_tot=t_n
        virtual_data.append({"label":label, "mat_virtual":mat_virt,
                              "lec_virtual":lec_virt, "mat_total":mat_tot,
                              "lec_total":lec_tot})

        # Changed to unpack the enrollment file `df_mat` per month so percent calculations are accurate
        grade_month_data.append((label, df_csvs_all, df_mat))

    if coberturas_por_mes:
        print(f"\n[*] Generando Excel Maestro de Cobertura...")
        coverage_master_to_excel(coberturas_por_mes,
            os.path.join(PATH_OUTPUT_COVERAGE,
                         "Cobertura_Maestro_Todos_Los_Meses.xlsx"))

    if school_part_by_mes:
        print(f"\n[*] Generando Excel Maestro de Participación por Escuela...")
        wb_master = Workbook(); wb_master.remove(wb_master.active)

        for label, df_sp in school_part_by_mes:
            school_participation_sheet(wb_master, label, df_sp)
            print(f"    Hoja '{label}': {len(df_sp)} escuelas")

        virtual_sheet(wb_master, virtual_data)
        print(f"    Hoja 'Centro Virtual': {len(virtual_data)} meses")

        grade_by_month_sheet(wb_master, grade_month_data)
        print(f"    Hoja 'Grado x Mes (todos)': {len(grade_month_data)} meses")

        master_path = os.path.join(PATH_OUTPUT_COVERAGE,
                                   "Participacion_Maestra_Por_Escuela.xlsx")
        wb_master.save(master_path)
        print(f"  [OK] Excel Maestro Participación: {os.path.basename(master_path)}")

    print(f"\n{'='*65}")
    print(f"  [OK] ¡Proceso completado!")
    print(f"       Cobertura maestro : {PATH_OUTPUT_COVERAGE}")
    print(f"       Reportes mensuales: {PATH_OUTPUT_BASE}")
    print(f"{'='*65}\n")