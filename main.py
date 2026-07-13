# main.py
import config
from src.data_loader import cargar_datos_limpios
from src.conversor_excel_csv import convertir_todos_excel_a_csv
from src.report_por_grados import generar_reporte_por_grados

try: 
    from src.auditoria_calidad import auditoria_calidad
except ImportError: 
    pass

def ejecutar_sistema_paars():
    print("\n====================================================")
    print(f"   SISTEMA PAARS - PROCESANDO: {config.MONTH_FOLDER}")
    print("====================================================\n")

    # ---------------------------------------------------------
    # 0. INTEGRACIÓN Y UNIFICACIÓN DE DATASETS CRUDOS (Opción 1)
    # ---------------------------------------------------------
    if config.REPORT_CHOICE == '1' and config.EXAM_TYPE in ["RESULTADOS", "PROGRESO"]:
        from src.integrador_datasets import unificar_datasets_crudos
        # AHORA APUNTA A PATH_REPORTS DIRECTAMENTE
        unificar_datasets_crudos(config.PATH_RAW, config.PATH_REPORTS)
        return

    # ---------------------------------------------------------
    # 1. BIFURCACIÓN DE AUDITORÍA (Opción 2)
    # ---------------------------------------------------------
    if config.REPORT_CHOICE == '2' and config.EXAM_TYPE in ["RESULTADOS", "PROGRESO"]:
        print("[*] Iniciando Módulo de Control de Calidad...")
        try:
            auditoria_calidad()
            print("\n[OK] Auditoría finalizada.")
        except NameError:
            print("\n[!] Módulo no creado.")
        return

    # ---------------------------------------------------------
    # FLUJO A Y B: RESULTADOS Y PROGRESO (Flujos de Reportes)
    # ---------------------------------------------------------
    if config.EXAM_TYPE in ["RESULTADOS", "PROGRESO"]:
        print("[*] Preparando archivos originales de Análisis Psicométrico...")
        convertir_todos_excel_a_csv()
        df_master, items = cargar_datos_limpios()

        if df_master is None or df_master.empty:
            print("[!] Error: No hay datos suficientes.")
            return

        if config.EXAM_TYPE == "RESULTADOS":
            if config.REPORT_CHOICE in ['3', '7']:
                from src.resultados_engine import calcular_metricas_resultados
                from src.report_resultados_latex import generar_reporte_latex
                generar_reporte_latex(calcular_metricas_resultados(df_master, items), df_master)
            if config.REPORT_CHOICE in ['4', '7']: generar_reporte_por_grados(df_master)
            if config.REPORT_CHOICE in ['5', '7']:
                pass 
            if config.REPORT_CHOICE in ['6', '7']:
                from src.generador_rankings import generador_rankings_masivo
                try: generador_rankings_masivo()
                except Exception as e: print(e)

        elif config.EXAM_TYPE == "PROGRESO":
            if config.REPORT_CHOICE in ['3', '8']:
                from src.report_profesores import process_section_reports
                process_section_reports()
            if config.REPORT_CHOICE in ['4', '8']:
                from src.report_directores import process_comparative_reports
                process_comparative_reports(df_master)
            if config.REPORT_CHOICE in ['5', '8']: generar_reporte_por_grados(df_master)
            if config.REPORT_CHOICE in ['6', '8']:
                from src.report_corto_latex import generar_reporte_corto_latex
                from src.analisis_comparativo_m1_m2 import ejecutar_comparativo
                generar_reporte_corto_latex(df_master)
                ejecutar_comparativo()
            if config.REPORT_CHOICE in ['7', '8']:
                from src.generador_rankings import generador_rankings_masivo
                try: generador_rankings_masivo()
                except Exception as e: print(e)

   # ---------------------------------------------------------
    # FLUJO C: PRUEBA DE FUNDAMENTO (Google Forms pipeline)
    # ---------------------------------------------------------
    elif config.EXAM_TYPE == "FUNDAMENTO":
        print("\n[-->] MODO: PRUEBA DE FUNDAMENTO DETECTADO")
        
        # --- Módulo: Generación de Reportes Clásicos ---
        if config.mod_choice == '2':
            if config.REPORT_CHOICE == '1':
                print("\n[*] Transformando Formulario Crudo en Datasets de Evaluación...")
                from src.sistematizador_fundamentos import procesar_y_generar_excel as procesar_fundamentos
                exito = procesar_fundamentos(config.CURRENT_MONTH_PATH)
                if not exito: return 
        
        # --- Módulo: Teoría de Respuesta al Ítem (TRI) ---
        elif config.mod_choice == '1':
            import subprocess
            import os
            
            # 1. Calibración: Modelo 2PL
            if config.REPORT_CHOICE == '1':
                # Definimos la nueva ruta específica para este modelo
                carpeta_codigos = os.path.join(os.getcwd(), "R codes", "TRI Fundamentos 2PL")
                
                scripts_a_ejecutar = [
                    "1. irt_fundamentos.R",
                    "2. Equate fundamentos.R",
                    "3. Reporte estudiantes fundamentos.R"
                ]
                
                print("\n[*] Iniciando Pipeline IRT 2PL (Fundamentos)...")
                
                for script in scripts_a_ejecutar:
                    ruta_script = os.path.join(carpeta_codigos, script)
                    print(f"\n[>] Ejecutando: {script} ...")
                    
                    try:
                        subprocess.run([config.R_EXE_PATH, ruta_script, config.CURRENT_MONTH_PATH], check=True)
                    except subprocess.CalledProcessError:
                        print(f"\n[!] ERROR CRÍTICO: Falló la ejecución de {script}. El pipeline se ha detenido.")
                        return # Detiene la ejecución para no correr el script 2 si el 1 falló
                    except FileNotFoundError:
                        print(f"\n[!] ERROR: No se encontró el script en: {ruta_script}")
                        return
                        
                print("\n[OK] Pipeline IRT 2PL (Fundamentos) completado con éxito.")
            
            # 2. Calibración: Modelo 3PL
            elif config.REPORT_CHOICE == '2':
                print("\n[!] El pipeline para Fundamentos 3PL está en desarrollo.")
    # ---------------------------------------------------------
    # FINALIZACIÓN
    # ---------------------------------------------------------
    print("\n====================================================")
    print("      SISTEMA PAARS COMPLETADO CON ÉXITO!")
    print("====================================================\n")

if __name__ == "__main__":
    ejecutar_sistema_paars()