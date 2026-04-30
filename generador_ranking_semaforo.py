import os
import glob
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
INPUT_DIR_M1 = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados"
INPUT_DIR_M2 = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"

OUTPUT_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Final_Reports"

# Extraer nombres de los meses dinámicamente
def extraer_mes(ruta):
    carpeta = os.path.basename(os.path.dirname(os.path.dirname(ruta)))
    match = re.search(r'^0?(\d+)_.*_([A-Za-z]+)$', carpeta)
    return match.group(2) if match else "Desconocido"

nombre_m1 = extraer_mes(INPUT_DIR_M1)
nombre_m2 = extraer_mes(INPUT_DIR_M2)

# ==========================================
# 2. DEFINICIÓN DEL SEMÁFORO
# ==========================================
SEMAFORO_CONFIG = [
    (35, 'Crítico', '991B1B', 'FFFFFF'),
    (45, 'Bajo', 'FF8C2E', 'FFFFFF'),
    (55, 'Medio', 'FACC15', '000000'),
    (65, 'Bueno', '84CC16', '000000'),
    (100, 'Excelente', '065F46', 'FFFFFF')
]

def obtener_semaforo(score):
    if pd.isna(score): return 'Sin Datos', 'FFFFFF', '000000'
    score = float(score)
    for limite, nombre, bg, fg in SEMAFORO_CONFIG:
        if score <= limite: return nombre, bg, fg
    return 'Excelente', '065F46', 'FFFFFF'

# ==========================================
# 3. PROCESAMIENTO DE DATOS
# ==========================================
def procesar_mes(ruta_dir):
    archivos = glob.glob(os.path.join(ruta_dir, "*.csv"))
    lista_mat, lista_lec = [], []

    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            materia = 'MAT' if 'MAT' in os.path.basename(f).upper() else ('LEC' if 'LEC' in os.path.basename(f).upper() or 'LENG' in os.path.basename(f).upper() else None)
            if not materia: continue

            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            col_centro = next((c for c in df.columns if 'centro' in c.lower() and 'nro' not in c.lower()), None)
            col_cod = next((c for c in df.columns if 'nro de centro' in c.lower() or 'código' in c.lower() or 'codigo' in c.lower()), None)

            if not col_theta or not col_centro: continue
            if col_anular: df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')]

            df['Código'] = df[col_cod].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() if col_cod else "N/D"
            df['Centro Escolar'] = df[col_centro].astype(str).str.strip()
            df = df[df['Código'] != '99999']
            
            df['Puntaje'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=['Puntaje'])

            if df.empty: continue
            df_agrupado = df.groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index()
            
            if materia == 'MAT': lista_mat.append(df_agrupado)
            else: lista_lec.append(df_agrupado)
        except Exception:
            pass

    df_mat = pd.concat(lista_mat).groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index() if lista_mat else pd.DataFrame()
    df_lec = pd.concat(lista_lec).groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index() if lista_lec else pd.DataFrame()
    
    return df_mat, df_lec

def generar_datos_comparativos():
    print(f"[*] Extrayendo datos de {nombre_m1}...")
    df_mat_m1, df_lec_m1 = procesar_mes(INPUT_DIR_M1)
    
    print(f"[*] Extrayendo datos de {nombre_m2}...")
    df_mat_m2, df_lec_m2 = procesar_mes(INPUT_DIR_M2)
    
    def cruzar_y_formatear(df1, df2):
        if df1.empty and df2.empty: return pd.DataFrame()
        
        # Unir ambos meses (outer join para no perder escuelas que falten en un mes)
        df_cruce = pd.merge(df1, df2, on=['Código', 'Centro Escolar'], how='outer', suffixes=('_M1', '_M2'))
        
        # Calcular diferencia
        df_cruce['Diferencia'] = df_cruce['Puntaje_M2'] - df_cruce['Puntaje_M1']
        
        # Redondear y rellenar valores faltantes
        df_cruce['Puntaje_M1'] = df_cruce['Puntaje_M1'].round(2)
        df_cruce['Puntaje_M2'] = df_cruce['Puntaje_M2'].round(2)
        df_cruce['Diferencia'] = df_cruce['Diferencia'].round(2)

        # Asignar Niveles
        df_cruce[f'Nivel {nombre_m1}'] = df_cruce['Puntaje_M1'].apply(lambda x: obtener_semaforo(x)[0])
        df_cruce[f'Nivel {nombre_m2}'] = df_cruce['Puntaje_M2'].apply(lambda x: obtener_semaforo(x)[0])
        
        # Ordenar columnas limpias
        df_final = df_cruce[['Código', 'Centro Escolar', 
                             'Puntaje_M1', f'Nivel {nombre_m1}', 
                             'Puntaje_M2', f'Nivel {nombre_m2}', 
                             'Diferencia']].copy()
        
        df_final.rename(columns={'Puntaje_M1': f'Puntaje {nombre_m1}', 'Puntaje_M2': f'Puntaje {nombre_m2}'}, inplace=True)
        return df_final.sort_values(by=f'Puntaje {nombre_m2}', ascending=True, na_position='first')

    print("[*] Cruzando resultados y calculando evolución...")
    return cruzar_y_formatear(df_mat_m1, df_mat_m2), cruzar_y_formatear(df_lec_m1, df_lec_m2)

# ==========================================
# 4. CREACIÓN DEL EXCEL CON COLORES
# ==========================================
def crear_excel_comparativo(df_mat, df_lec):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ruta_excel = os.path.join(OUTPUT_DIR, f"Evolucion_Escuelas_{nombre_m1}_vs_{nombre_m2}.xlsx")
    
    wb = Workbook()
    if 'Sheet' in wb.sheetnames: wb.remove(wb['Sheet'])

    materias = [("Evolución_Matemática", df_mat), ("Evolución_Lengua", df_lec)]

    for nombre_hoja, df in materias:
        if df.empty: continue
        ws = wb.create_sheet(title=nombre_hoja)
        
        # Agregar cabeceras
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Diseño de Cabecera Principal
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.column_dimensions['A'].width = 15  # Código
        ws.column_dimensions['B'].width = 50  # Centro Escolar
        ws.column_dimensions['C'].width = 15  # Puntaje M1
        ws.column_dimensions['D'].width = 20  # Nivel M1
        ws.column_dimensions['E'].width = 15  # Puntaje M2
        ws.column_dimensions['F'].width = 20  # Nivel M2
        ws.column_dimensions['G'].width = 15  # Diferencia

        # Colorear Celdas de Datos
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=7):
            # Centrar textos numéricos
            row[2].alignment = Alignment(horizontal="center") # Puntaje M1
            row[4].alignment = Alignment(horizontal="center") # Puntaje M2
            row[6].alignment = Alignment(horizontal="center") # Diferencia

            # Colorear Nivel M1 (Columna D / Índice 3)
            val_m1 = row[2].value
            if val_m1 is not None and not pd.isna(val_m1):
                _, bg1, fg1 = obtener_semaforo(val_m1)
                row[3].fill = PatternFill(start_color=bg1, end_color=bg1, fill_type="solid")
                row[3].font = Font(color=fg1, bold=True)
                row[3].alignment = Alignment(horizontal="center", vertical="center")

            # Colorear Nivel M2 (Columna F / Índice 5)
            val_m2 = row[4].value
            if val_m2 is not None and not pd.isna(val_m2):
                _, bg2, fg2 = obtener_semaforo(val_m2)
                row[5].fill = PatternFill(start_color=bg2, end_color=bg2, fill_type="solid")
                row[5].font = Font(color=fg2, bold=True)
                row[5].alignment = Alignment(horizontal="center", vertical="center")

            # Colorear Diferencia (Columna G / Índice 6)
            dif = row[6].value
            if dif is not None and not pd.isna(dif):
                if dif > 0: row[6].font = Font(color="008000", bold=True) # Verde si subió
                elif dif < 0: row[6].font = Font(color="FF0000", bold=True) # Rojo si bajó

    wb.save(ruta_excel)
    print(f"\n[OK] Excel Comparativo generado exitosamente en:\n     {ruta_excel}")

if __name__ == "__main__":
    print("====================================================")
    print("   GENERADOR DE REPORTE COMPARATIVO DE ESCUELAS")
    print("====================================================")
    
    df_m, df_l = generar_datos_comparativos()
    
    if df_m.empty and df_l.empty:
        print("[!] No se encontraron datos para generar la comparativa.")
    else:
        crear_excel_comparativo(df_m, df_l)