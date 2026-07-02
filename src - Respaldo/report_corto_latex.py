# src/report_corto_latex.py
import os
import pandas as pd
import numpy as np
import config

LATEX_BEAMER_TEMPLATE = r"""\documentclass[aspectratio=169]{beamer}

\usepackage[utf8]{inputenc}
\usepackage[spanish]{babel}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{colortbl}
\usepackage{adjustbox}
\usepackage{tikz}
\usepackage{makecell}

% --- COLORES CORPORATIVOS ---
\definecolor{DeepNavy}{HTML}{033b6d}
\definecolor{RoyalBlue}{HTML}{094b93}
\definecolor{OceanBlue}{HTML}{3077b9}
\definecolor{SkyBlue}{HTML}{a7d2f2}
\definecolor{SteelGrey}{HTML}{666766}

% --- TEMA BEAMER ---
\setbeamercolor{structure}{fg=DeepNavy}
\setbeamercolor{frametitle}{fg=DeepNavy, bg=SkyBlue!30}
\setbeamercolor{title}{fg=DeepNavy, bg=SkyBlue!20}
\setbeamercolor{item}{fg=RoyalBlue}
\setbeamercolor{block title}{bg=RoyalBlue, fg=white}
\setbeamercolor{block body}{bg=SkyBlue!10, fg=SteelGrey}

\setbeamertemplate{navigation symbols}{}
\setbeamertemplate{footline}[frame number]

% --- PORTADA ---
\title{\textbf{Resultados de la aplicación}}
\subtitle{Prueba de @@TIPO_PRUEBA@@, 2026}
\author{Gerencia de Evaluación}
\date{@@MES_APLICACION@@ 2026}

\begin{document}

\begin{frame}
    \titlepage
\end{frame}

% --- TABLA 1: COBERTURA ---
\begin{frame}{Cobertura de la Prueba de @@TIPO_PRUEBA@@, 2026}
    \vspace{-0.3cm}
    \begin{center}
    \adjustbox{max width=\textwidth}{
        \renewcommand{\arraystretch}{1.3}
        \begin{tabular}{l c cc cc >{\color{RoyalBlue}\bfseries}c >{\color{RoyalBlue}\bfseries}c}
        \toprule
        \rowcolor{DeepNavy}
        \textcolor{white}{\textbf{Grado}} & 
        \textcolor{white}{\textbf{\makecell{Matrícula \\ B1}}} & 
        \textcolor{white}{\textbf{\makecell{Registros SOTE\\ Matemática}}} & 
        \textcolor{white}{\textbf{\makecell{Registros SOTE \\ Lengua}}} & 
        \textcolor{white}{\textbf{\makecell{Puntajes válidos \\ Matemática}}} & 
        \textcolor{white}{\textbf{\makecell{Puntajes válidos \\ Lengua}}} & 
        \textcolor{white}{\textbf{\makecell{Cobertura \\ Matemática}}} & 
        \textcolor{white}{\textbf{\makecell{Cobertura \\ Lengua}}} \\
        \midrule
@@FILAS_TABLA_COBERTURA@@
        \bottomrule
        \end{tabular}
    }
    \end{center}
    \vfill
    \tiny \textcolor{SteelGrey}{\textit{Nota: La cobertura se calcula con la fórmula $\frac{\text{Registros SOTE}}{\text{Matrícula B1}} \times 100$.}}
\end{frame}

% --- TABLA 2: ESTADÍSTICAS ---
\begin{frame}{Estadísticas Descriptivas por Grado}
    \vspace{-0.3cm}
    \begin{center}
    \adjustbox{max width=0.85\textwidth}{
        \renewcommand{\arraystretch}{1.3}
        \begin{tabular}{l ccc ccc}
        \toprule
        \rowcolor{RoyalBlue}
        \textcolor{white}{\textbf{Grado}} & 
        \multicolumn{3}{c}{\textcolor{white}{\textbf{Matemática (Escala 0-100)}}} & 
        \multicolumn{3}{c}{\textcolor{white}{\textbf{Lenguaje (Escala 0-100)}}} \\
        \cmidrule(lr){2-4} \cmidrule(lr){5-7}
        \rowcolor{RoyalBlue}
        & \textcolor{white}{\textbf{Media}} & \textcolor{white}{\textbf{Mediana}} & \textcolor{white}{\textbf{DS}} 
        & \textcolor{white}{\textbf{Media}} & \textcolor{white}{\textbf{Mediana}} & \textcolor{white}{\textbf{DS}} \\
        \midrule
@@FILAS_TABLA_ESTADISTICAS@@
        \bottomrule
        \end{tabular}
    }
    \end{center}
    \vfill
\end{frame}

% --- GRÁFICAS (MARCADORES PARA MAÑANA) ---
\begin{frame}{Distribución de Puntajes (3.$^\circ$, 6.$^\circ$ y 9.$^\circ$)}
    \centering
    \includegraphics[width=\textwidth,height=0.8\textheight,keepaspectratio]{Graficos_PPT/Boxplots_3_6_9.png}
\end{frame}

\begin{frame}{Estudiantes que alcanzaron 60 puntos o más}
    \centering
    \includegraphics[width=\textwidth,height=0.8\textheight,keepaspectratio]{Graficos_PPT/Pictogramas_Combinados_3_6_9.png}
\end{frame}

\begin{frame}{Distribución por Intervalos (Todos los Grados)}
    \centering
    \includegraphics[width=\textwidth,height=0.8\textheight,keepaspectratio]{Graficos_PPT/Barras_Intervalos_Completas.png}
\end{frame}

\end{document}
"""

def generar_reporte_corto_latex(df_master):
    print("  [Reporte Corto] Calculando tablas de Cobertura y Estadísticas...")
    
    # 1. Cargar Matrícula B1
    ruta_matricula = os.path.join(config.PATH_METADATA, "Matricula_B1.xlsx")
    dict_matricula = {}
    if os.path.exists(ruta_matricula):
        try:
            df_matr = pd.read_excel(ruta_matricula)
            for _, row in df_matr.iterrows():
                dict_matricula[int(row['Grado'])] = int(row['Matricula'])
        except Exception as e:
            print(f"  [!] Error leyendo Matricula_B1.xlsx: {e}")
    else:
        print(f"  [!] No se encontró {ruta_matricula}. Se usarán guiones en lugar de la matrícula.")

    df_master['Grado_Num'] = df_master['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
    grados_presentes = sorted([g for g in df_master['Grado_Num'].unique() if g > 0])
    
    filas_cobertura = ""
    filas_estadisticas = ""
    
    for grado in grados_presentes:
        matricula = dict_matricula.get(grado, 0)
        str_matricula = f"{matricula:,}" if matricula > 0 else "-"
        
        # Filtrar por materia
        df_mat = df_master[(df_master['Grado_Num'] == grado) & (df_master['Area temática'] == 'Matemática')].copy()
        df_len = df_master[(df_master['Grado_Num'] == grado) & (df_master['Area temática'] == 'Lengua')].copy()
        
        # Calcular SOTE (Total registros) y Válidos (Con Theta calculado)
        sote_mat = len(df_mat)
        sote_len = len(df_len)
        
        df_mat['theta.global (escala 0-100)'] = pd.to_numeric(df_mat['theta.global (escala 0-100)'], errors='coerce')
        df_len['theta.global (escala 0-100)'] = pd.to_numeric(df_len['theta.global (escala 0-100)'], errors='coerce')
        
        val_mat = df_mat['theta.global (escala 0-100)'].notna().sum()
        val_len = df_len['theta.global (escala 0-100)'].notna().sum()
        
        cob_mat = f"{(sote_mat / matricula * 100):.1f}\\%" if matricula > 0 and sote_mat > 0 else "-"
        cob_len = f"{(sote_len / matricula * 100):.1f}\\%" if matricula > 0 and sote_len > 0 else "-"
        
        # Fila Cobertura
        filas_cobertura += f"        {grado}.$^\\circ$ & {str_matricula} & {sote_mat:,} & {sote_len:,} & {val_mat:,} & {val_len:,} & {cob_mat} & {cob_len} \\\\\n"
        
        # Calcular Estadísticas
        med_mat = f"{df_mat['theta.global (escala 0-100)'].mean():.2f}" if val_mat > 0 else "-"
        mdn_mat = f"{df_mat['theta.global (escala 0-100)'].median():.2f}" if val_mat > 0 else "-"
        ds_mat = f"{df_mat['theta.global (escala 0-100)'].std(ddof=1):.2f}" if val_mat > 1 else "-"
        
        med_len = f"{df_len['theta.global (escala 0-100)'].mean():.2f}" if val_len > 0 else "-"
        mdn_len = f"{df_len['theta.global (escala 0-100)'].median():.2f}" if val_len > 0 else "-"
        ds_len = f"{df_len['theta.global (escala 0-100)'].std(ddof=1):.2f}" if val_len > 1 else "-"
        
        # Fila Estadísticas
        filas_estadisticas += f"        {grado}.$^\\circ$  & {med_mat} & {mdn_mat} & {ds_mat} & {med_len} & {mdn_len} & {ds_len} \\\\\n"

    # Preparar el archivo final
    tex_final = LATEX_BEAMER_TEMPLATE.replace("@@FILAS_TABLA_COBERTURA@@", filas_cobertura)
    tex_final = tex_final.replace("@@FILAS_TABLA_ESTADISTICAS@@", filas_estadisticas)
    tex_final = tex_final.replace("@@TIPO_PRUEBA@@", "PROGRESO" if config.EXAM_TYPE == "PROGRESO" else "RESULTADOS")
    
    # Extraer el mes del nombre de la carpeta (ej. 03_PROGRESO_Marzo -> Marzo)
    mes_str = config.MONTH_FOLDER.split('_')[-1]
    tex_final = tex_final.replace("@@MES_APLICACION@@", mes_str)

    # Crear carpeta y guardar
    output_dir = os.path.join(config.PATH_REPORTS, "Reporte_Corto_LaTeX")
    os.makedirs(os.path.join(output_dir, "Graficos_PPT"), exist_ok=True)
    
    tex_path = os.path.join(output_dir, f"Reporte_Corto_Beamer_{mes_str}.tex")
    with open(tex_path, 'w', encoding='utf-8') as f:
        f.write(tex_final)
        
    print(f"  [OK] ¡Reporte Corto Generado! Archivo LaTeX guardado en: {output_dir}")
    print("  [NOTA] Las gráficas de Beamer se implementarán en la siguiente fase de desarrollo.")