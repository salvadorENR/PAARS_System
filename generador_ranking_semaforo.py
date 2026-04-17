import os
import glob
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS (HARDCODED)
# ==========================================
INPUT_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados"

# Guardaremos el Excel una carpeta atrás, en Final_Reports
OUTPUT_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Final_Reports"

# ==========================================
# 2. DEFINICIÓN DEL SEMÁFORO (COLORES HEX)
# ==========================================
# Formato: (Límite superior, Nombre, Color Fondo, Color Letra)
SEMAFORO_CONFIG = [
    (35, 'Crítico', '991B1B', 'FFFFFF'),   # Rojo oscuro, letra blanca
    (45, 'Bajo', 'FF8C2E', 'FFFFFF'),      # Naranja, letra blanca
    (55, 'Medio', 'FACC15', '000000'),     # Amarillo, letra negra
    (65, 'Bueno', '84CC16', '000000'),     # Verde claro, letra negra
    (100, 'Excelente', '065F46', 'FFFFFF') # Verde oscuro, letra blanca
]

def obtener_semaforo(score):
    """Devuelve la categoría y los colores según el puntaje."""
    if pd.isna(score):
        return 'Sin Datos', 'FFFFFF', '000000'
    score = float(score)
    for limite, nombre, bg_color, font_color in SEMAFORO_CONFIG:
        if score <= limite:
            return nombre, bg_color, font_color
    return 'Excelente', '065F46', 'FFFFFF'

# ==========================================
# 3. PROCESAMIENTO DE DATOS
# ==========================================
def cargar_datos_ranking():
    print(f"[*] Escaneando CSVs en:\n    {INPUT_DIR}")
    archivos = glob.glob(os.path.join(INPUT_DIR, "*.csv"))
    
    if not archivos:
        print("[!] No se encontraron archivos CSV en la ruta especificada.")
        return pd.DataFrame(), pd.DataFrame()

    lista_mat = []
    lista_lec = []

    for f in archivos:
        try:
            # Leer CSV tolerando errores de codificación
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            
            # Identificar la materia por el nombre del archivo
            nombre_archivo = os.path.basename(f).upper()
            if 'MAT' in nombre_archivo:
                materia = 'MAT'
            elif 'LEC' in nombre_archivo or 'LENG' in nombre_archivo:
                materia = 'LEC'
            else:
                continue

            # Identificar columnas clave
            col_theta = next((c for c in df.columns if '0-100' in c.lower()), None)
            col_anular = next((c for c in df.columns if 'anular' in c.lower()), None)
            col_centro = next((c for c in df.columns if 'centro' in c.lower() and 'nro' not in c.lower()), None)
            col_cod_centro = next((c for c in df.columns if 'nro de centro' in c.lower() or 'código' in c.lower() or 'codigo' in c.lower()), None)

            if not col_theta or not col_centro:
                continue

            # Filtrar pruebas anuladas
            if col_anular:
                df = df[df[col_anular].isna() | (df[col_anular].astype(str).str.strip() == '')]

            # Limpiar datos
            df['Código'] = df[col_cod_centro].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() if col_cod_centro else "N/D"
            df['Centro Escolar'] = df[col_centro].astype(str).str.strip()
            
            # Excluir centro virtual
            df = df[df['Código'] != '99999']
            
            # Convertir nota a numérico (manejando comas europeas si existen)
            df['Puntaje'] = pd.to_numeric(df[col_theta].astype(str).str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=['Puntaje'])

            if df.empty:
                continue

            # Agrupar por escuela y sacar el promedio
            df_agrupado = df.groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index()
            
            if materia == 'MAT':
                lista_mat.append(df_agrupado)
            else:
                lista_lec.append(df_agrupado)
                
        except Exception as e:
            print(f"  [!] Error leyendo {os.path.basename(f)}: {e}")

    # Consolidar todas las bases de cada materia
    df_final_mat = pd.concat(lista_mat, ignore_index=True) if lista_mat else pd.DataFrame()
    df_final_lec = pd.concat(lista_lec, ignore_index=True) if lista_lec else pd.DataFrame()

    # Si hay escuelas repetidas en varios CSVs de la misma materia, agruparlas de nuevo
    if not df_final_mat.empty:
        df_final_mat = df_final_mat.groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index()
        # ORDEN DE MENOR A MAYOR (ascending=True)
        df_final_mat = df_final_mat.sort_values(by='Puntaje', ascending=True)
        df_final_mat['Puntaje'] = df_final_mat['Puntaje'].round(2)
        df_final_mat['Nivel de Logro'] = df_final_mat['Puntaje'].apply(lambda x: obtener_semaforo(x)[0])

    if not df_final_lec.empty:
        df_final_lec = df_final_lec.groupby(['Código', 'Centro Escolar'])['Puntaje'].mean().reset_index()
        # ORDEN DE MENOR A MAYOR (ascending=True)
        df_final_lec = df_final_lec.sort_values(by='Puntaje', ascending=True)
        df_final_lec['Puntaje'] = df_final_lec['Puntaje'].round(2)
        df_final_lec['Nivel de Logro'] = df_final_lec['Puntaje'].apply(lambda x: obtener_semaforo(x)[0])

    return df_final_mat, df_final_lec

# ==========================================
# 4. CREACIÓN DEL EXCEL CON COLORES
# ==========================================
def crear_excel_semaforo(df_mat, df_lec):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ruta_excel = os.path.join(OUTPUT_DIR, "Ranking_Semaforo_Mes2_Abril.xlsx")
    
    wb = Workbook()
    
    # Remover la hoja por defecto
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    materias = [("Ranking_Matemática", df_mat), ("Ranking_Lengua", df_lec)]

    for nombre_hoja, df in materias:
        if df.empty:
            continue
            
        ws = wb.create_sheet(title=nombre_hoja)
        
        # Insertar los datos de Pandas a OpenPyXL
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Dar formato a la cabecera
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Ajustar ancho de columnas
        ws.column_dimensions['A'].width = 15  # Código
        ws.column_dimensions['B'].width = 60  # Centro Escolar
        ws.column_dimensions['C'].width = 20  # Puntaje
        ws.column_dimensions['D'].width = 25  # Nivel de Logro

        # Aplicar los colores del Semáforo a las filas de datos
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=4):
            puntaje_celda = row[2].value
            if puntaje_celda is None: continue
            
            # Obtener colores basados en el puntaje
            _, bg_hex, font_hex = obtener_semaforo(puntaje_celda)
            
            # Pintar la celda de "Nivel de Logro" (Columna D que es row[3])
            celda_nivel = row[3]
            celda_nivel.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
            celda_nivel.font = Font(color=font_hex, bold=True)
            celda_nivel.alignment = Alignment(horizontal="center", vertical="center")
            
            # Alinear el puntaje al centro
            row[2].alignment = Alignment(horizontal="center")

    wb.save(ruta_excel)
    print(f"\n[OK] Excel con Ranking y Semáforo generado exitosamente en:")
    print(f"     {ruta_excel}")

# ==========================================
# EJECUCIÓN
# ==========================================
if __name__ == "__main__":
    print("====================================================")
    print("   GENERADOR INDEPENDIENTE DE RANKING (MES 2)")
    print("====================================================")
    
    df_m, df_l = cargar_datos_ranking()
    
    if df_m.empty and df_l.empty:
        print("[!] No se encontraron datos válidos para generar el reporte.")
    else:
        crear_excel_semaforo(df_m, df_l)