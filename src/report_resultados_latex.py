# src/report_resultados_latex.py
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import config

# --- 1. COLORES OFICIALES ---
COLOR_DEEP_NAVY = '#033b6d'
COLOR_ROYAL_BLUE = '#094b93'
COLOR_OCEAN_BLUE = '#3077b9'
COLOR_SKY_BLUE = '#a7d2f2'
COLOR_STEEL_GREY = '#666766'

# --- 2. PLANTILLAS LATEX ---
LATEX_HEADER = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[spanish,es-tabla]{babel}
\decimalpoint 
\usepackage{geometry}
\geometry{margin=2.5cm}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{float}
\usepackage{titlesec}
\usepackage{caption}
\usepackage{tabularx}
\usepackage{array}
\usepackage{threeparttable}

% --- IMAGE PATH SETUP ---
\graphicspath{{Graphs/}} 

% --- INTERACTIVITY SETUP ---
\usepackage[colorlinks=true, linkcolor=blue!70!black, citecolor=blue!70!black, urlcolor=blue!70!black]{hyperref}

% Formatting for section titles
\titleformat{\section}{\large\bfseries\color{blue!70!black}}{\thesection.}{1em}{}
\titleformat{\subsection}{\normalsize\bfseries}{\thesubsection.}{1em}{}

\title{\textbf{Resultados Descriptivos de la Aplicación} \\ Prueba de Resultados}
\date{Febrero 2026}
\author{Gerencia de Evaluación}

\begin{document}

\maketitle
\thispagestyle{empty} 
\clearpage            

\tableofcontents
\newpage
\listoffigures
\newpage
\listoftables
\newpage

\section*{Introducción}
\addcontentsline{toc}{section}{Introducción}
El presente informe técnico expone los resultados descriptivos derivados de la aplicación de la \textbf{Prueba de Resultados}.

\newpage
\section{Resumen Global de Distribución}
A continuación, la Figura~\ref{fig:boxplot_globales} presenta los diagramas de caja que muestran la distribución de los resultados obtenidos en Matemática y Lengua.

\begin{figure}[htbp]
    \centering
    \includegraphics[width=0.95\textwidth]{boxplot_globales}
    \caption{Distribución Global Estudiantil: Matemática y Lengua}
    \label{fig:boxplot_globales}
\end{figure}

En el panel, se evidencia la variabilidad de los datos representados en las cajas...
\newpage
"""

LATEX_GRADE_TEMPLATE = r"""
% ==========================================
% SECTION: GRADE @@GRADO@@
% ==========================================
\section{Informe del Grado @@GRADO@@ - Estadísticas y Quintiles}

\textbf{Tasa de Finalización del Grado (Válidos):} \\ 
LENGUA: Válidos: @@N_LEN@@ \\ 
MATEMÁTICA: Válidos: @@N_MAT@@ 

\subsection{Estadísticas Descriptivas y Quintiles}
La Tabla~\ref{tab:desc_grado@@GRADO@@} presenta los estadísticos descriptivos.

\begin{table}[htbp]
    \centering
    \begin{tabular}{lcccccc}
        \toprule
        \textbf{Materia} & \textbf{N} & \textbf{Media} & \textbf{Mediana} & \textbf{DS} & \textbf{Mínimo} & \textbf{Máximo} \\
        \midrule
        Lengua & @@N_LEN@@ & @@MED_LEN@@ & @@MDN_LEN@@ & @@DS_LEN@@ & @@MIN_LEN@@ & @@MAX_LEN@@ \\ 
        Matemática & @@N_MAT@@ & @@MED_MAT@@ & @@MDN_MAT@@ & @@DS_MAT@@ & @@MIN_MAT@@ & @@MAX_MAT@@ \\ 
        \bottomrule
    \end{tabular}
    \caption{Grado @@GRADO@@: Estadísticas Descriptivas}
    \label{tab:desc_grado@@GRADO@@}
\end{table}

\subsection{Gráficos de Distribución e Ítems}

Las Figuras~\ref{fig:hist_est_grado@@GRADO@@} y~\ref{fig:hist_esc_grado@@GRADO@@} presentan la distribución de los puntajes a nivel de estudiante y de centro educativo, respectivamente. 
\newpage

\begin{figure}[H]
    \centering
    \begin{minipage}[t]{0.48\textwidth}
        \centering
        \vspace{0pt}
        \includegraphics[width=\linewidth]{hist_estudiantes_grado_@@GRADO@@}
        \caption{Grado @@GRADO@@: Puntaje Global Estudiantes} 
        \label{fig:hist_est_grado@@GRADO@@}
    \end{minipage}\hfill
    \begin{minipage}[t]{0.48\textwidth}
        \centering
        \vspace{0pt}
        \includegraphics[width=\linewidth]{hist_escuelas_grado_@@GRADO@@}
        \caption{Grado @@GRADO@@: Puntaje Global Escuelas} 
        \label{fig:hist_esc_grado@@GRADO@@}
    \end{minipage}
\end{figure}

\begin{figure}[htbp]
    \centering
    \includegraphics[width=0.95\textwidth]{barras_dificultad_grado_@@GRADO@@}
    \caption{Grado @@GRADO@@: Porcentaje de aciertos por Ítem} 
    \label{fig:items_grado@@GRADO@@}
\end{figure}

@@COMENTARIO_DINAMICO@@
\newpage
"""

LATEX_FOOTER = r"\end{document}"

# --- 3. FUNCIONES DE DIBUJO ---
def dibujar_boxplot_global(df_master, graph_dir):
    plt.figure(figsize=(11, 8.5))
    df_plot = df_master.copy()
    df_plot['Grado_Num'] = df_plot['Grado'].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)
    df_plot['theta.global (escala 0-100)'] = pd.to_numeric(df_plot['theta.global (escala 0-100)'], errors='coerce')
    df_plot = df_plot.dropna(subset=['theta.global (escala 0-100)'])
    
    sns.boxplot(data=df_plot, x='Grado_Num', y='theta.global (escala 0-100)', hue='Area temática', 
                palette=[COLOR_OCEAN_BLUE, COLOR_ROYAL_BLUE])
    
    plt.title("Distribución Global (Escala 0-100)", fontweight='bold', color=COLOR_DEEP_NAVY)
    plt.ylabel("Puntaje Global")
    plt.ylim(0, 100)
    plt.legend(title='Materia')
    plt.savefig(os.path.join(graph_dir, "boxplot_globales.png"), bbox_inches='tight', dpi=300)
    plt.close()

def dibujar_graficos_grado(grado, data_mat, data_len, graph_dir):
    # 1. Histogramas Estudiantes
    fig, axes = plt.subplots(1, 2, figsize=(11, 4))
    if data_len:
        sns.histplot(data_len['df_estudiantes']['theta.global (escala 0-100)'], bins=20, color=COLOR_ROYAL_BLUE, edgecolor='white', ax=axes[0])
        axes[0].axvline(data_len['desc_stats']['Media'], color=COLOR_DEEP_NAVY, linestyle=':', linewidth=2)
        axes[0].set_title("Puntajes Estudiantes: Lengua", color=COLOR_ROYAL_BLUE)
    if data_mat:
        sns.histplot(data_mat['df_estudiantes']['theta.global (escala 0-100)'], bins=20, color=COLOR_OCEAN_BLUE, edgecolor='white', ax=axes[1])
        axes[1].axvline(data_mat['desc_stats']['Media'], color=COLOR_DEEP_NAVY, linestyle=':', linewidth=2)
        axes[1].set_title("Puntajes Estudiantes: Matemática", color=COLOR_OCEAN_BLUE)
    plt.tight_layout()
    plt.savefig(os.path.join(graph_dir, f"hist_estudiantes_grado_{grado}.png"), bbox_inches='tight', dpi=300)
    plt.close()

    # 2. Histogramas Escuelas
    fig, axes = plt.subplots(1, 2, figsize=(11, 4))
    if data_len and not data_len['df_escuelas'].empty:
        sns.histplot(data_len['df_escuelas']['theta.global (escala 0-100)'], bins=20, color=COLOR_ROYAL_BLUE, edgecolor='white', ax=axes[0])
        axes[0].set_title("Promedios Escuelas: Lengua", color=COLOR_ROYAL_BLUE)
    if data_mat and not data_mat['df_escuelas'].empty:
        sns.histplot(data_mat['df_escuelas']['theta.global (escala 0-100)'], bins=20, color=COLOR_OCEAN_BLUE, edgecolor='white', ax=axes[1])
        axes[1].set_title("Promedios Escuelas: Matemática", color=COLOR_OCEAN_BLUE)
    plt.tight_layout()
    plt.savefig(os.path.join(graph_dir, f"hist_escuelas_grado_{grado}.png"), bbox_inches='tight', dpi=300)
    plt.close()

    # 3. Barras de Dificultad (Ítems)
    fig, axes = plt.subplots(2, 1, figsize=(11, 8))
    if data_len and not data_len['items_df'].empty:
        sns.barplot(data=data_len['items_df'], x='Item', y='Pct', color=COLOR_ROYAL_BLUE, ax=axes[0])
        axes[0].set_title("Porcentaje de Aciertos por Ítem: Lengua", color=COLOR_ROYAL_BLUE, fontweight='bold')
        axes[0].tick_params(axis='x', rotation=90, labelsize=7)
        for container in axes[0].containers: axes[0].bar_label(container, fmt='%.1f%%', padding=3, rotation=90, size=7)
    if data_mat and not data_mat['items_df'].empty:
        sns.barplot(data=data_mat['items_df'], x='Item', y='Pct', color=COLOR_OCEAN_BLUE, ax=axes[1])
        axes[1].set_title("Porcentaje de Aciertos por Ítem: Matemática", color=COLOR_OCEAN_BLUE, fontweight='bold')
        axes[1].tick_params(axis='x', rotation=90, labelsize=7)
        for container in axes[1].containers: axes[1].bar_label(container, fmt='%.1f%%', padding=3, rotation=90, size=7)
    
    plt.tight_layout()
    plt.savefig(os.path.join(graph_dir, f"barras_dificultad_grado_{grado}.png"), bbox_inches='tight', dpi=300)
    plt.close()

# --- 4. EL ENSAMBLADOR PRINCIPAL ---
def generar_reporte_latex(report_data_dict, df_master):
    print("  [Generador LaTeX] Creando gráficas y ensamblando el código fuente...")
    
    # Crear carpetas
    output_dir = os.path.join(config.PATH_REPORTS, "Reporte_LaTeX")
    graph_dir = os.path.join(output_dir, "Graphs")
    os.makedirs(graph_dir, exist_ok=True)
    
    # 1. Dibujar gráfica global
    dibujar_boxplot_global(df_master, graph_dir)
    
    # 2. Iniciar el archivo LaTeX
    latex_final = LATEX_HEADER
    
    # 3. Iterar por cada grado
    grados_ordenados = sorted(list(report_data_dict.keys()))
    for grado in grados_ordenados:
        data_mat = report_data_dict[grado].get('Matemática')
        data_len = report_data_dict[grado].get('Lengua')
        
        if not data_mat and not data_len: continue
            
        # Dibujar gráficas del grado
        dibujar_graficos_grado(grado, data_mat, data_len, graph_dir)
        
        # Inyectar datos en la plantilla
        seccion_grado = LATEX_GRADE_TEMPLATE.replace("@@GRADO@@", str(grado))
        
        if data_len:
            st = data_len['desc_stats']
            seccion_grado = seccion_grado.replace("@@N_LEN@@", f"{st['N']:,}")
            seccion_grado = seccion_grado.replace("@@MED_LEN@@", f"{st['Media']:.2f}")
            seccion_grado = seccion_grado.replace("@@MDN_LEN@@", f"{st['Mediana']:.2f}")
            seccion_grado = seccion_grado.replace("@@DS_LEN@@", f"{st['DS']:.2f}")
            seccion_grado = seccion_grado.replace("@@MIN_LEN@@", f"{st['Mínimo']:.2f}")
            seccion_grado = seccion_grado.replace("@@MAX_LEN@@", f"{st['Máximo']:.2f}")
            txt_len = f"en Lengua, el porcentaje de aciertos está entre el {data_len['textos'].get('item_min_val', 0):.1f}\\% ({data_len['textos'].get('item_min_nombre', '')}) y el {data_len['textos'].get('item_max_val', 0):.1f}\\% ({data_len['textos'].get('item_max_nombre', '')})"
        else:
            seccion_grado = seccion_grado.replace("@@N_LEN@@", "-").replace("@@MED_LEN@@", "-").replace("@@MDN_LEN@@", "-").replace("@@DS_LEN@@", "-").replace("@@MIN_LEN@@", "-").replace("@@MAX_LEN@@", "-")
            txt_len = "no hay datos suficientes de Lengua"

        if data_mat:
            st = data_mat['desc_stats']
            seccion_grado = seccion_grado.replace("@@N_MAT@@", f"{st['N']:,}")
            seccion_grado = seccion_grado.replace("@@MED_MAT@@", f"{st['Media']:.2f}")
            seccion_grado = seccion_grado.replace("@@MDN_MAT@@", f"{st['Mediana']:.2f}")
            seccion_grado = seccion_grado.replace("@@DS_MAT@@", f"{st['DS']:.2f}")
            seccion_grado = seccion_grado.replace("@@MIN_MAT@@", f"{st['Mínimo']:.2f}")
            seccion_grado = seccion_grado.replace("@@MAX_MAT@@", f"{st['Máximo']:.2f}")
            txt_mat = f"para Matemática, el porcentaje de aciertos está entre el {data_mat['textos'].get('item_min_val', 0):.1f}\\% ({data_mat['textos'].get('item_min_nombre', '')}) y el {data_mat['textos'].get('item_max_val', 0):.1f}\\% ({data_mat['textos'].get('item_max_nombre', '')})"
        else:
            seccion_grado = seccion_grado.replace("@@N_MAT@@", "-").replace("@@MED_MAT@@", "-").replace("@@MDN_MAT@@", "-").replace("@@DS_MAT@@", "-").replace("@@MIN_MAT@@", "-").replace("@@MAX_MAT@@", "-")
            txt_mat = "no hay datos suficientes de Matemática"

        # Armar el comentario dinámico final
        comentario_dinamico = f"La Figura~\\ref{{fig:items_grado{grado}}} revela que, {txt_len}, mientras que, {txt_mat}."
        seccion_grado = seccion_grado.replace("@@COMENTARIO_DINAMICO@@", comentario_dinamico)
        
        latex_final += seccion_grado
        
    # 4. Cerrar y guardar
    latex_final += LATEX_FOOTER
    
    latex_path = os.path.join(output_dir, "Reporte_Resultados_2026.tex")
    with open(latex_path, 'w', encoding='utf-8') as f:
        f.write(latex_final)
        
    print(f"  [OK] ¡Generación completada! El código fuente y las imágenes están en: {output_dir}")