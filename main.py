# main.py
import config
from src.pre_processor import ejecutar_conversion
from src.data_loader import cargar_datos_limpios
from src.conversor_excel_csv import convertir_todos_excel_a_csv

# --- NUEVO MÓDULO DE CONTROL DE CALIDAD (PASO 3) ---
# Lo importamos desde ahora para que esté listo cuando lo refactoricemos
try:
    from src.auditoria_calidad import auditoria_calidad
except ImportError:
    pass # Evita errores de importación temporalmente hasta completar el Paso 3

# --- MÁQUINAS DE PROGRESO ---
from src.stats_engine import (
    generar_estadisticas_basicas, 
    calcular_distribucion_niveles, 
    obtener_quintiles
)
from src.visualizer import generar_grafico_niveles, generar_grafico_quintiles
from src.report_profesores import process_section_reports
from src.report_directores import process_comparative_reports
from src.generador_rankings import generador_rankings_masivo 

# --- MÁQUINAS DE RESULTADOS (LaTeX) ---
from src.resultados_engine import calcular_metricas_resultados
from src.report_resultados_latex import generar_reporte_latex
from src.report_corto_latex import generar_reporte_corto_latex

def ejecutar_sistema_paars():
    
    print("\n====================================================")
    print(f"   SISTEMA PAARS - PROCESANDO: {config.MONTH_FOLDER}")
    print("====================================================\n")

    # ---------------------------------------------------------
    # BIFURCACIÓN DE AUDITORÍA (OPCIÓN 1 EN CUALQUIER FLUJO)
    # ---------------------------------------------------------
    if config.REPORT_CHOICE == 1:
        print("[*] Iniciando Módulo de Control de Calidad de Bases de Datos...")
        try:
            auditoria_calidad()
            print("\n[OK] Auditoría de calidad finalizada con éxito.")
        except NameError:
            print("\n[!] El módulo 'auditoria_calidad' aún no ha sido creado (Paso 3 pendiente).")
        return # Termina la ejecución, ya que solo se solicitó auditar la base.

    # ---------------------------------------------------------
    # PREPARACIÓN DE DATOS (PARA EL RESTO DE REPORTES)
    # ---------------------------------------------------------
    print("[*] Preparando archivos originales de Análisis Psicométrico...")
    convertir_todos_excel_a_csv()

    print("\n[*] Calculando el panorama general del país/estado...")
    df_master, items = cargar_datos_limpios()

    if df_master is None or df_master.empty:
        print("[!] Error: No hay datos suficientes para armar estadísticas.")
        return

    # ---------------------------------------------------------
    # FLUJO 1: PRUEBA DE PROGRESO
    # ---------------------------------------------------------
    if config.EXAM_TYPE == "PROGRESO":
        print("\n[-->] MODO: PRUEBA DE PROGRESO DETECTADO")
        
        # Gráficas ejecutivas generales (Siempre se generan como base)
        stats = generar_estadisticas_basicas(df_master)
        niveles = calcular_distribucion_niveles(df_master)
        quintiles = obtener_quintiles(df_master)
        
        print("\n[*] Dibujando gráficas ejecutivas generales...")
        generar_grafico_niveles(niveles)
        generar_grafico_quintiles(quintiles)

        # Regla Especial: Opciones 2 o 5 (Reporte de Estudiantes)
        if config.REPORT_CHOICE in [2, 5]:
            print("\n[!] AVISO: Informe de estudiantes está aún en diseño.")

        # Opción 3 o 5: Reporte de Profesores
        if config.REPORT_CHOICE in [3, 5]:
            print("\n[*] Generando Reportes Detallados por Sección (Para Profesores)...")
            process_section_reports()

        # Opción 4 o 5: Reporte de Directores
        if config.REPORT_CHOICE in [4, 5]:
            if config.PATH_PREV_INTERIM:
                print("\n[*] Generando Reportes Comparativos por Escuela (Para Directores)...")
                process_comparative_reports(df_master)
            else:
                print("\n[*] Omitiendo comparativa (No seleccionaste un mes anterior).")

        # Cierre estándar de Progreso (Siempre se generan los consolidados macro)
        print("\n[*] Generando Reporte Corto Institucional (Beamer LaTeX)...")
        generar_reporte_corto_latex(df_master)

        print("\n[*] Generando Rankings Macro e Histórico de Escuelas (Excel)...")
        try:
            generador_rankings_masivo()
        except Exception as e:
            print(f"  [!] Error al generar los rankings macro: {e}")

    # ---------------------------------------------------------
    # FLUJO 2: PRUEBA DE RESULTADOS
    # ---------------------------------------------------------
    elif config.EXAM_TYPE == "RESULTADOS":
        print("\n[-->] MODO: PRUEBA DE RESULTADOS DETECTADO (LaTeX)")
        
        # Opción 2 o 4: Generar Tablas y Gráficas (Reporte Formal Completo)
        if config.REPORT_CHOICE in [2, 4]:
            print("\n[*] Extrayendo estadísticas, quintiles y métricas de ítems para Reporte Formal...")
            report_data_dict = calcular_metricas_resultados(df_master, items)
            
            print("\n[*] Generando Reporte Institucional Formal en LaTeX...")
            if report_data_dict:
                generar_reporte_latex(report_data_dict, df_master)
            else:
                print("  [!] No se generó el reporte porque no hay datos válidos.")
                
        # Opción 3 o 4: Generar informes por escuelas (Reporte Corto Beamer)
        if config.REPORT_CHOICE in [3, 4]:
            print("\n[*] Generando Reporte Corto en LaTeX (Beamer)...")
            generar_reporte_corto_latex(df_master)

    # ---------------------------------------------------------
    # FINALIZACIÓN
    # ---------------------------------------------------------
    print("\n====================================================")
    print("      SISTEMA PAARS COMPLETADO CON ÉXITO!")
    print(f"   Todos los reportes están en: {config.PATH_REPORTS}")
    print("====================================================\n")

if __name__ == "__main__":
    ejecutar_sistema_paars()