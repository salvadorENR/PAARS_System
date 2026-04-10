# main.py
import config
from src.pre_processor import ejecutar_conversion
from src.data_loader import cargar_datos_limpios

# --- MÁQUINAS DE PROGRESO (HTML) ---
from src.stats_engine import (
    generar_estadisticas_basicas, 
    calcular_distribucion_niveles, 
    obtener_quintiles
)
from src.visualizer import generar_grafico_niveles, generar_grafico_quintiles
from src.report_profesores import process_section_reports
from src.report_directores import process_comparative_reports

# --- NUEVAS MÁQUINAS DE RESULTADOS (LaTeX) ---
from src.resultados_engine import calcular_metricas_resultados
from src.report_resultados_latex import generar_reporte_latex
from src.report_corto_latex import generar_reporte_corto_latex

def ejecutar_sistema_paars():
    # Nota: El menú interactivo ya se ejecutó arriba durante el 'import config'
    
    print("\n====================================================")
    print(f"   SISTEMA PAARS - PROCESANDO: {config.MONTH_FOLDER}")
    print("====================================================\n")

    # 1. PASO: Convertir Excel a CSV
    print("[1/5] Preparando archivos originales de Analisis Psicometrico...")
    
    # ⚠️ ¡ATENCIÓN! ⚠️
    # Esta línea está apagada (con #) temporalmente para que la simulación de Abril/Mayo funcione 
    # y no sobreescriba los datos del futuro con los de marzo.
    # CUANDO VAYAS A PROCESAR DATOS REALES NUEVOS, QUÍTALE EL "#" A LA SIGUIENTE LÍNEA:
    # ejecutar_conversion()

    # 2. PASO: Cargar e Integrar datos
    print("\n[2/5] Calculando el panorama general del pais/estado...")
    df_master, items = cargar_datos_limpios()

    if df_master is None or df_master.empty:
        print("[!] Error: No hay datos suficientes para armar estadisticas.")
        return

    # --- BIFURCACION DE FLUJO SEGUN EL TIPO DE EXAMEN ---
    if config.EXAM_TYPE == "PROGRESO":
        print("\n[-->] MODO: PRUEBA DE PROGRESO DETECTADO")
        
        # 3. PASO: Generar Estadisticas y Graficas Ejecutivas
        stats = generar_estadisticas_basicas(df_master)
        niveles = calcular_distribucion_niveles(df_master)
        quintiles = obtener_quintiles(df_master)
        
        print("\n[3/6] Dibujando graficas ejecutivas generales...")
        generar_grafico_niveles(niveles)
        generar_grafico_quintiles(quintiles)

        # 4. PASO: Generar Reportes por Seccion (Profesores)
        if config.REPORT_CHOICE in [1, 3]:
            print("\n[4/6] Generando Reportes Detallados por Seccion (Para Profesores)...")
            process_section_reports()
        else:
            print("\n[4/6] Omitiendo Reportes por Seccion (Opcion no seleccionada en el menu).")

        # 5. PASO: Generar Reportes Comparativos (Directores)
        if config.REPORT_CHOICE in [2, 3]:
            if config.PATH_PREV_INTERIM:
                print("\n[5/6] Generando Reportes Comparativos por Escuela (Para Directores)...")
                process_comparative_reports(df_master)
            else:
                print("\n[5/6] Omitiendo comparativa (No seleccionaste un mes anterior en el menu).")
        else:
            print("\n[5/6] Omitiendo Reportes Comparativos (Opcion no seleccionada en el menu).")

        # 6. PASO: Generar Reporte Corto (LaTeX)
        print("\n[6/6] Generando Reporte Corto Institucional (Beamer LaTeX)...")
        generar_reporte_corto_latex(df_master)

    else:
        print("\n[-->] MODO: PRUEBA DE RESULTADOS DETECTADO (LaTeX)")
        
        if config.REPORT_CHOICE in [1, 3]:
            # 3. PASO: Cálculos Matemáticos Avanzados
            print("\n[3/5] Extrayendo estadisticas, quintiles y metricas de items para Reporte Formal...")
            report_data_dict = calcular_metricas_resultados(df_master, items)
            
            # 4. PASO: Generación del Documento y Gráficas
            print("\n[4/5] Generando Reporte Institucional Formal en LaTeX...")
            if report_data_dict:
                generar_reporte_latex(report_data_dict, df_master)
            else:
                print("  [!] No se genero el reporte porque no hay datos validos.")
        else:
            print("\n[3/5] Omitiendo Reporte Formal (Opcion no seleccionada en el menu).")
            print("[4/5] Omitiendo Reporte Formal (Opcion no seleccionada en el menu).")
            
        if config.REPORT_CHOICE in [2, 3]:
            # 5. PASO: Generar Reporte Corto (LaTeX)
            print("\n[5/5] Generando Reporte Corto en LaTeX (Beamer)...")
            generar_reporte_corto_latex(df_master)
        else:
            print("\n[5/5] Omitiendo Reporte Corto (Opcion no seleccionada en el menu).")

    print("\n====================================================")
    print("      SISTEMA PAARS COMPLETADO CON EXITO!")
    print(f"  Todos los reportes estan en: {config.PATH_REPORTS}")
    print("====================================================")

if __name__ == "__main__":
    ejecutar_sistema_paars()