import os
import sys
import pandas as pd
import glob
import re
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go
import unicodedata
import numpy as np

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
DIRECTORIOS_MESES = [
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_PROGRESO_Marzo\Interim_CSVs\Resultados",  "Mes 1 (Marzo)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados",  "Mes 2 (Abril)"),
    (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Interim_CSVs\Resultados",   "Mes 3 (Mayo)")
]

PATH_METADATA     = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\MatriculaProgresoMes3.csv"
PATH_REQ027       = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\00_Metadata\REQ_027_1.xlsx"
PATH_OUTPUT_PLOTS = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\03_PROGRESO_Mayo\Final_Reports\Tablas_Graficas_ReportCorto"
os.makedirs(PATH_OUTPUT_PLOTS, exist_ok=True)

# ==========================================
# 2. CONFIGURACIÓN DE COLORES Y VARIABLES
# ==========================================
COL_SCORE_EXACTA = 'theta.global (escala 0-100)'

NIVELES          = ['Excelente', 'Bueno', 'Medio', 'Bajo', 'Crítico']
ORDEN_CATEGORIAS = ["Crítico", "Bajo", "Medio", "Bueno", "Excelente"]
COLOR_CATEGORIAS = {
    "Crítico":   "#991b1b",
    "Bajo":      "#ff8c2e",
    "Medio":     "#facc15",
    "Bueno":     "#84cc16",
    "Excelente": "#065f46"
}

COLORES_MESES = ["#80b1d3", "#fb8072", "#b3de69"]
COLORES_NODO  = ['#065f46', '#84cc16', '#facc15', '#ff8c2e', '#991b1b']
COLORES_LINK  = [
    'rgba(6, 95, 70, 0.4)',
    'rgba(132, 204, 22, 0.4)',
    'rgba(250, 204, 21, 0.4)',
    'rgba(255, 140, 46, 0.4)',
    'rgba(153, 27, 27, 0.4)'
]

GRUPOS_GRADOS = {
    "2° y 3° Grado":      [2, 3],
    "4°, 5° y 6° Grado":  [4, 5, 6],
    "7°, 8° y 9° Grado":  [7, 8, 9],
    "10° y 11° Grado":    [10, 11],
}

# ==========================================
# 3. CLASIFICACIÓN DE PUNTAJE
# ==========================================
def clasificar_puntaje(val):
    if pd.isna(val): return ''
    val = float(val)
    if val <= 35:   return 'Crítico'
    elif val <= 45: return 'Bajo'
    elif val <= 55: return 'Medio'
    elif val <= 65: return 'Bueno'
    else:           return 'Excelente'

def extraer_numero_grado(grado_str):
    if pd.isna(grado_str): return None
    match = re.search(r'\d+', str(grado_str).lower())
    return int(match.group()) if match else None

# ==========================================
# 4. LECTURA Y PROCESAMIENTO
# ==========================================
def preparar_entorno():
    sns.set_theme(style="whitegrid")
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14

def cargar_matricula():
    print(f"\n[*] Cargando Matrícula desde: {PATH_METADATA}")
    if not os.path.exists(PATH_METADATA):
        print("  [!] ERROR: No se encontró el archivo de matrícula.")
        return pd.DataFrame()
    try:
        df_mat = pd.read_csv(PATH_METADATA, sep=';', dtype=str,
                             encoding_errors='ignore')
        df_mat.columns = df_mat.columns.str.strip().str.upper()
        if 'NIE' not in df_mat.columns:
            print("  [!] ERROR: columna 'NIE' no encontrada.")
            return pd.DataFrame()
        df_mat['NIE_Limpio'] = (df_mat['NIE'].astype(str)
                                .str.replace(r'\.0$', '', regex=True)
                                .str.strip())
        df_mat['GRUPO_Limpio'] = (
            df_mat['GRUPO'].astype(str).str.strip().str.upper()
            if 'GRUPO' in df_mat.columns else 'DESCONOCIDO')
        return df_mat
    except Exception as e:
        print(f"  [!] Error leyendo matrícula: {e}")
        return pd.DataFrame()

def cargar_req027():
    print(f"\n[*] Cargando REQ_027_1 desde: {PATH_REQ027}")
    if not os.path.exists(PATH_REQ027):
        print("  [!] ADVERTENCIA: REQ_027_1.xlsx no encontrado — "
              "CÓDIGO_SECCIÓN quedará vacío.")
        return pd.Series(dtype=str)

    try:
        df = pd.read_excel(PATH_REQ027, dtype=str)
    except Exception as e:
        print(f"  [!] Error leyendo REQ_027_1.xlsx: {e}")
        return pd.Series(dtype=str)

    df.columns = df.columns.str.strip().str.upper()
    if 'GROUP_ID' not in df.columns:
        return pd.Series(dtype=str)

    nie_col = next(
        (c for c in df.columns
         if c in ('NIE', 'DOCUMENTO', 'ID_ESTUDIANTE',
                  'STUDENT_ID', 'DOC', 'NID')),
        None)

    if nie_col is None:
        candidates = [c for c in df.columns if c != 'GROUP_ID']
        nie_col    = candidates[0] if candidates else None

    if nie_col is None:
        return pd.Series(dtype=str)

    df['_KEY'] = (df[nie_col].astype(str)
                  .str.replace(r'\.0$',  '', regex=True)
                  .str.replace('"',      '', regex=False)
                  .str.replace("'",      '', regex=False)
                  .str.strip())

    df['GROUP_ID'] = (df['GROUP_ID'].astype(str)
                      .str.replace('"', '', regex=False)
                      .str.strip())

    df = df[df['_KEY'].str.len() > 0]
    df = df[df['GROUP_ID'].str.len() > 0]
    df = df[df['GROUP_ID'].str.upper() != 'NAN']

    lookup = (df.drop_duplicates(subset='_KEY', keep='last')
                .set_index('_KEY')['GROUP_ID'])
    return lookup

def _aplicar_group_id(df_out, lookup_req027):
    if lookup_req027.empty:
        return df_out

    nie_col = next(
        (c for c in df_out.columns
         if c.upper() in ('NIE', 'NIE_LIMPIO')), None)

    if nie_col is None:
        return df_out

    nie_key = (df_out[nie_col]
               .astype(str)
               .str.replace(r'\.0$', '', regex=True)
               .str.strip())

    group_ids = nie_key.map(lookup_req027)
    df_out['CÓDIGO_SECCIÓN'] = group_ids.fillna('').values
    return df_out

def extraer_csvs_puntajes():
    lista_dfs = []
    print("[*] Extrayendo puntajes de los CSVs...")
    for ruta, etiqueta_mes in DIRECTORIOS_MESES:
        if not os.path.exists(ruta): continue
        for f in glob.glob(os.path.join(ruta, "*.csv")):
            if 'legend' in os.path.basename(f).lower(): continue
            try:
                df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
                df.columns = df.columns.str.strip()
                materia = ('Matemática'
                           if 'MAT' in os.path.basename(f).upper()
                           else 'Lengua')
                if COL_SCORE_EXACTA not in df.columns: continue

                col_anular = next(
                    (c for c in df.columns if 'anular' in c.lower()), None)
                col_grado  = next(
                    (c for c in df.columns if 'grado'  in c.lower()), None)
                col_codigo = next(
                    (c for c in df.columns
                     if 'nro de centro' in c.lower()
                     or 'código'        in c.lower()
                     or 'codigo'        in c.lower()
                     or ('centro' in c.lower()
                         and 'id' in c.lower())), None)
                col_centro = next(
                    (c for c in df.columns
                     if c.lower() in ('centro',
                                      'nombre centro',
                                      'institución')), None)
                if not col_centro:
                    col_centro = next(
                        (c for c in df.columns
                         if 'centro'  in c.lower()
                         and 'nro'    not in c.lower()
                         and 'código' not in c.lower()), None)
                col_grupo  = next(
                    (c for c in df.columns
                     if 'grupo'   in c.lower()
                     or 'sección' in c.lower()
                     or 'seccion' in c.lower()), None)
                col_doc    = next(
                    (c for c in df.columns
                     if 'documento' in c.lower()
                     or 'nie'       in c.lower()
                     or c.lower() == 'id'), None)

                if not (col_grado and col_codigo and col_doc): continue

                if col_anular:
                    def excluir_por_tiempo(val):
                        if pd.isna(val): return False
                        v_norm = ''.join(
                            c for c in unicodedata.normalize(
                                'NFD', str(val).lower())
                            if unicodedata.category(c) != 'Mn')
                        return ('5 min'    in v_norm or
                                'demora'   in v_norm or
                                'duracion' in v_norm)
                    df = df[~df[col_anular]
                            .apply(excluir_por_tiempo)].copy()

                df[COL_SCORE_EXACTA] = pd.to_numeric(
                    df[COL_SCORE_EXACTA].astype(str)
                    .str.replace(',', '.'), errors='coerce')
                df = df.dropna(subset=[COL_SCORE_EXACTA])
                if df.empty: continue

                df['Grado_Num']     = df[col_grado].apply(extraer_numero_grado)
                df['Grado_Str']     = df[col_grado].astype(str).str.strip()
                df['Nro de Centro'] = (df[col_codigo].astype(str)
                                       .str.replace(r'\.0$', '',
                                                    regex=True)
                                       .str.strip())
                df['Centro']        = (df[col_centro].astype(str).str.strip()
                                       if col_centro else "Desconocido")
                df['Grupo']         = (df[col_grupo].astype(str).str.strip()
                                       if col_grupo  else "Desconocido")
                df['Documento']     = (df[col_doc].astype(str)
                                       .str.replace(r'\.0$', '',
                                                    regex=True)
                                       .str.strip())
                df['Nivel_Logro']   = df[COL_SCORE_EXACTA].apply(clasificar_puntaje)

                df_final = df[[COL_SCORE_EXACTA, 'Grado_Num', 'Grado_Str',
                               'Nivel_Logro', 'Nro de Centro', 'Centro',
                               'Grupo', 'Documento']].copy()
                df_final['Materia'] = materia
                df_final['Mes']     = etiqueta_mes
                df_final = df_final[df_final['Nro de Centro'] != '99999']
                df_final = df_final[
                    ~df_final['Centro'].str.contains(
                        'Centro Virtual', case=False, na=False)]
                lista_dfs.append(df_final)
            except: pass

    return (pd.concat(lista_dfs, ignore_index=True)
            if lista_dfs else pd.DataFrame())


# ==========================================
# 5. CONTEOS Y ALERTAS (SIN FILTROS)
# ==========================================
def generar_reportes_conteos_separados(df_base_total):
    print("\n[*] Generando Reportes de Totales de Estudiantes (Matriciales, Sin filtros)...")
    grados_esperados = list(range(2, 12))
    
    def crear_hoja_materia(df, materia):
        df_sub = df[df['Materia'] == materia]
        if df_sub.empty:
            return pd.DataFrame()
            
        pivot = df_sub.pivot_table(
            index='Mes', columns='Grado_Num', values='Documento', 
            aggfunc='nunique', fill_value=0
        )
        for col in grados_esperados:
            if col not in pivot.columns:
                pivot[col] = 0
                
        pivot = pivot[grados_esperados]
        pivot.columns = [f"{c}°" for c in pivot.columns]
        pivot['Total por prueba'] = pivot.sum(axis=1)
        return pivot

    def exportar_archivo(df_target, filename):
        path_output = os.path.join(PATH_OUTPUT_PLOTS, filename)
        with pd.ExcelWriter(path_output, engine='openpyxl') as writer:
            for materia in ['Matemática', 'Lengua']:
                df_hoja = crear_hoja_materia(df_target, materia)
                if not df_hoja.empty:
                    df_hoja.to_excel(writer, sheet_name=materia)
                else:
                    pd.DataFrame({'Aviso': [f'Sin datos para {materia}']}).to_excel(writer, sheet_name=materia, index=False)
        print(f"  [OK] Guardado: {filename}")

    df_b1 = df_base_total[df_base_total['GRUPO_Limpio'] == 'B1']
    df_b2 = df_base_total[df_base_total['GRUPO_Limpio'] == 'B2']

    exportar_archivo(df_base_total, "Conteo_Estudiantes_B1_y_B2.xlsx")
    exportar_archivo(df_b1, "Conteo_Estudiantes_Solo_B1.xlsx")
    exportar_archivo(df_b2, "Conteo_Estudiantes_Solo_B2.xlsx")

def alertar_b2_mes1_mes2(df_base_total, meses_labels):
    print("\n[*] Verificando escuelas B2 con estudiantes en Mes 1 o Mes 2...")
    if len(meses_labels) >= 2:
        meses_tempranos = meses_labels[:2] 
        df_b2_temprano = df_base_total[
            (df_base_total['GRUPO_Limpio'] == 'B2') & 
            (df_base_total['Mes'].isin(meses_tempranos))
        ]
        
        if not df_b2_temprano.empty:
            print("  [!] ATENCIÓN: Se encontraron escuelas del Grupo B2 con participación temprana:")
            conteo = df_b2_temprano.groupby(['Nro de Centro', 'Centro', 'Materia', 'Mes'])['Documento'].nunique().reset_index()
            for _, row in conteo.iterrows():
                print(f"      - {row['Nro de Centro']} - {row['Centro']} | {row['Materia']} | {row['Mes']}: {row['Documento']} estudiante(s)")
        else:
            print("  [OK] Ninguna escuela B2 fue evaluada en Mes 1 o Mes 2.")
    else:
        print("  [-] No hay suficientes meses definidos en los directorios para esta validación.")


# ==========================================
# 6. SANKEY
# ==========================================
def calc_y_centers(df_mes):
    totals   = [len(df_mes[df_mes['Nivel'] == n]) for n in NIVELES]
    pad_frac = 0.04
    usable   = 1.0 - pad_frac * (len(NIVELES) - 1)
    grand    = sum(totals) or 1
    heights  = [max((v / grand) * usable, 0.01) for v in totals]
    total_h  = sum(heights)
    if total_h > usable:
        heights = [h * usable / total_h for h in heights]
    y_centers, cursor = [], 0.0
    for h in heights:
        y_centers.append(round(cursor + h / 2, 4))
        cursor += h + pad_frac
    return y_centers

def dibujar_sankey_multimes(df_master, materia, meses_labels):
    print(f"  -> Dibujando Sankey para {materia} (Grupo B1)...")
    df_mat = df_master[df_master['Materia'] == materia].copy()
    if df_mat.empty: return

    meses_validos, nombres_validos = [], []
    for mes in meses_labels:
        df_m = df_mat[df_mat['Mes'] == mes]
        if not df_m.empty:
            promedios = (df_m.groupby(['Nro de Centro', 'Centro'])
                         [COL_SCORE_EXACTA].mean().reset_index())
            promedios['Nivel'] = promedios[COL_SCORE_EXACTA].apply(
                clasificar_puntaje)
            promedios = promedios.rename(
                columns={'Nro de Centro': 'Código'})
            meses_validos.append(promedios)
            nombres_validos.append(mes)
    if len(meses_validos) < 2: return

    labels_nodos, colores_nodos_lista, node_x, node_y = [], [], [], []
    sources, targets, values, link_colors = [], [], [], []

    df_stats = pd.merge(
        meses_validos[0][['Código', 'Nivel']],
        meses_validos[-1][['Código', 'Nivel']],
        on='Código', suffixes=('_Ini', '_Fin'), how='inner')
    n_centros = len(df_stats)
    mantienen = suben = bajan = 0
    if n_centros > 0:
        for _, row in df_stats.iterrows():
            if (row['Nivel_Ini'] not in NIVELES or
                    row['Nivel_Fin'] not in NIVELES): continue
            d = (NIVELES.index(row['Nivel_Fin']) -
                 NIVELES.index(row['Nivel_Ini']))
            if d == 0:  mantienen += 1
            elif d < 0: suben     += 1
            else:       bajan     += 1
        subtitulo_html = (
            f"<span style='font-size:13px;color:#888'>"
            f"Población base (Grupo B1): {n_centros} centros<br>"
            f"Mantienen: {mantienen} ({mantienen/n_centros*100:.1f}%) · "
            f"Suben: {suben} ({suben/n_centros*100:.1f}%) · "
            f"Bajan: {bajan} ({bajan/n_centros*100:.1f}%)</span>"
        )
    else:
        subtitulo_html = ""

    for i, (nombre, df_mes) in enumerate(
            zip(nombres_validos, meses_validos)):
        x_val     = 0.01 + i * (0.98 / max(1, len(nombres_validos) - 1))
        centros_y = calc_y_centers(df_mes)
        for j, n in enumerate(NIVELES):
            labels_nodos.append(f"{nombre}: {n}")
            colores_nodos_lista.append(COLORES_NODO[j])
            node_x.append(x_val)
            node_y.append(centros_y[j])

    for i in range(len(meses_validos) - 1):
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        for j in range(len(NIVELES)):
            sources.append(oo + j); targets.append(od + j)
            values.append(1e-9); link_colors.append('rgba(0,0,0,0)')

    for i in range(len(meses_validos) - 1):
        df_cruce = pd.merge(
            meses_validos[i][['Código', 'Centro', 'Nivel']],
            meses_validos[i + 1][['Código', 'Nivel']],
            on='Código', suffixes=('_A', '_S'), how='inner')
        if df_cruce.empty: continue

        if materia == "Matemática" and i == 0:
            alerta = df_cruce[
                (df_cruce['Nivel_A'] == 'Bueno') &
                (df_cruce['Nivel_S'] == 'Bajo')]
            if not alerta.empty:
                print(f"\n      [ALERTA] Escuelas Bueno→Bajo Matemática "
                      f"({nombres_validos[i]}→{nombres_validos[i+1]}):")
                for _, r in alerta.iterrows():
                    print(f"         - {r['Código']} | {r['Centro']}")

        flujos = (df_cruce.groupby(['Nivel_A', 'Nivel_S'])
                  .size().reset_index(name='Cantidad'))
        oo = i * len(NIVELES); od = (i + 1) * len(NIVELES)
        mo = {n: j + oo for j, n in enumerate(NIVELES)}
        md = {n: j + od for j, n in enumerate(NIVELES)}
        for _, row in flujos.iterrows():
            if row['Cantidad'] == 0: continue
            si = mo[row['Nivel_A']]; ti = md[row['Nivel_S']]
            sources.append(si); targets.append(ti)
            values.append(row['Cantidad'])
            link_colors.append(COLORES_LINK[si % len(NIVELES)])

    if not sources: return

    fig = go.Figure(data=[go.Sankey(
        arrangement='fixed',
        node=dict(pad=20, thickness=24,
                  line=dict(color="white", width=0.5),
                  label=labels_nodos,
                  color=colores_nodos_lista,
                  x=node_x, y=node_y),
        link=dict(source=sources, target=targets,
                  value=values, color=link_colors)
    )])
    fig.update_layout(
        title=dict(
            text=(f"<span style='font-size:22px;color:#1a2b4c;"
                  f"font-family:Arial,sans-serif;font-weight:bold;'>"
                  f"Trayectorias de niveles — {materia}</span><br>"
                  f"{subtitulo_html}"),
            x=0.01, y=0.95),
        font_size=12,
        width=max(900, 800 + len(meses_validos) * 150),
        height=600,
        margin=dict(l=20, r=20, t=100, b=20),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    fig.write_html(os.path.join(PATH_OUTPUT_PLOTS,
                                f"Sankey_Evolutivo_{materia}.html"))
    try:
        fig.write_image(
            os.path.join(PATH_OUTPUT_PLOTS,
                         f"Sankey_Evolutivo_{materia}.png"), scale=2)
    except: pass

# ==========================================
# 7. BARRAS APILADAS POR GRUPO DE GRADOS Y TODAS JUNTAS
# ==========================================
def generar_barras_por_grupo_grados(df_master, meses_labels, sufijo_archivo="B1"):
    print(f"\n[*] Generando Barras por Grupo de Grados ({sufijo_archivo})...")
    for nombre_grupo, grados in GRUPOS_GRADOS.items():
        df_grupo = df_master[df_master['Grado_Num'].isin(grados)].copy()
        if df_grupo.empty: continue
        df_grupo['Grado_Str'] = (df_grupo['Grado_Num'].astype(str) + "° Grado")
        grados_orden = [f"{g}° Grado"
                        for g in sorted(grados, reverse=True)
                        if g in df_grupo['Grado_Num'].unique()]

        for materia in df_grupo['Materia'].unique():
            df_mat = df_grupo[df_grupo['Materia'] == materia].copy()
            if df_mat.empty: continue
            meses_con_datos = [m for m in meses_labels
                               if not df_mat[df_mat['Mes'] == m].empty]
            if not meses_con_datos: continue

            n_meses   = len(meses_con_datos)
            fig, axes = plt.subplots(1, n_meses,
                                     figsize=(max(9, n_meses * 9), 7),
                                     sharey=True)
            if n_meses == 1: axes = [axes]

            for i, mes_key in enumerate(meses_con_datos):
                ax         = axes[i]
                df_mes_bar = df_mat[df_mat['Mes'] == mes_key]
                if df_mes_bar.empty:
                    ax.set_visible(False); continue
                ct = pd.crosstab(df_mes_bar['Grado_Str'],
                                 df_mes_bar['Nivel_Logro'],
                                 normalize='index') * 100
                for cat in ORDEN_CATEGORIAS:
                    if cat not in ct.columns: ct[cat] = 0
                ct = ct[ORDEN_CATEGORIAS].reindex(grados_orden)
                ct.plot(kind='barh', stacked=True, ax=ax,
                        color=[COLOR_CATEGORIAS[c]
                               for c in ORDEN_CATEGORIAS],
                        legend=False, edgecolor='white', width=0.6)
                prueba_num = meses_labels.index(mes_key) + 1
                mes_nombre = (mes_key.split('(')[-1]
                              .replace(')', '').strip())
                ax.set_title(f"Prueba {prueba_num}\n({mes_nombre})",
                             fontsize=16, color='#005288',
                             pad=14, fontweight='bold')
                ax.set_xlim(0, 100)
                ax.set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)
                ax.set_ylabel("Grado" if i == 0 else "")
                if i == n_meses - 1:
                    h, l = ax.get_legend_handles_labels()
                    ax.legend(h, l, title="Nivel de Logro",
                              bbox_to_anchor=(1.02, 1),
                              loc='upper left', fontsize=9)
                for p in ax.patches:
                    w = p.get_width()
                    if w > 4:
                        ax.text(p.get_x() + w / 2,
                                p.get_y() + p.get_height() / 2,
                                f"{w:.1f}%", ha='center', va='center',
                                color='white', fontsize=9,
                                fontweight='bold')

            plt.suptitle(
                f"Niveles de Logro — {sufijo_archivo} — {materia}\n"
                f"{nombre_grupo}",
                fontsize=16, fontweight='bold', color='#1a2b4c')
            plt.tight_layout(rect=[0, 0, 0.88, 1])
            nombre_seguro = (nombre_grupo.replace('°', '')
                             .replace(' ', '_').replace(',', ''))
            fname = (f"Barras_{sufijo_archivo}_{materia}_"
                     f"{nombre_seguro}.png")
            plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, fname),
                        dpi=300, bbox_inches='tight')
            plt.close()
            print(f"    [OK] {fname}")

def generar_barras_todos_los_grados(df_master, meses_labels, sufijo_archivo="B1"):
    print(f"\n[*] Generando Barras (Todos los Grados juntos) ({sufijo_archivo})...")
    for materia in df_master['Materia'].unique():
        df_mat = df_master[df_master['Materia'] == materia].copy()
        if df_mat.empty: continue
        
        meses_con_datos = [m for m in meses_labels if not df_mat[df_mat['Mes'] == m].empty]
        if not meses_con_datos: continue

        n_meses = len(meses_con_datos)
        # Hacemos la figura un poco más alta para acomodar de 2° a 11°
        fig, axes = plt.subplots(1, n_meses, figsize=(max(9, n_meses * 9), 9), sharey=True)
        if n_meses == 1: axes = [axes]

        # Configurar orden de grados de 11 a 2, más "Global" abajo del todo
        grados_presentes = sorted(df_mat['Grado_Num'].dropna().unique(), reverse=True)
        grados_orden = [f"{int(g)}° Grado" for g in grados_presentes]
        grados_orden.append("Global (Todos)")

        for i, mes_key in enumerate(meses_con_datos):
            ax = axes[i]
            df_mes_bar = df_mat[df_mat['Mes'] == mes_key].copy()
            if df_mes_bar.empty:
                ax.set_visible(False); continue
            
            df_mes_bar['Grado_Str'] = df_mes_bar['Grado_Num'].astype(str) + "° Grado"
            
            # Crosstab por cada grado
            ct_grados = pd.crosstab(df_mes_bar['Grado_Str'], df_mes_bar['Nivel_Logro'], normalize='index') * 100
            
            # Fila extra "Global" agrupando a todos los estudiantes de ese mes
            ct_global = df_mes_bar['Nivel_Logro'].value_counts(normalize=True).to_frame().T * 100
            ct_global.index = ["Global (Todos)"]
            
            # Combinar y ordenar
            ct = pd.concat([ct_grados, ct_global])
            for cat in ORDEN_CATEGORIAS:
                if cat not in ct.columns: ct[cat] = 0
            
            orden_actual = [g for g in grados_orden if g in ct.index]
            ct = ct[ORDEN_CATEGORIAS].reindex(orden_actual)
            
            ct.plot(kind='barh', stacked=True, ax=ax,
                    color=[COLOR_CATEGORIAS[c] for c in ORDEN_CATEGORIAS],
                    legend=False, edgecolor='white', width=0.7)
            
            prueba_num = meses_labels.index(mes_key) + 1
            mes_nombre = (mes_key.split('(')[-1].replace(')', '').strip())
            ax.set_title(f"Prueba {prueba_num}\n({mes_nombre})",
                         fontsize=16, color='#005288', pad=14, fontweight='bold')
            ax.set_xlim(0, 100)
            ax.set_xlabel("Porcentaje de Estudiantes (%)", fontsize=10)
            ax.set_ylabel("Grado" if i == 0 else "")
            
            if i == n_meses - 1:
                h, l = ax.get_legend_handles_labels()
                ax.legend(h, l, title="Nivel de Logro", bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)
            
            for p in ax.patches:
                w = p.get_width()
                if w > 4:
                    ax.text(p.get_x() + w / 2, p.get_y() + p.get_height() / 2,
                            f"{w:.1f}%", ha='center', va='center',
                            color='white', fontsize=9, fontweight='bold')

        plt.suptitle(f"Niveles de Logro — {sufijo_archivo} — {materia}\n(Vista Completa Todos los Grados)",
                     fontsize=16, fontweight='bold', color='#1a2b4c')
        plt.tight_layout(rect=[0, 0, 0.88, 1])
        fname = f"Barras_{sufijo_archivo}_{materia}_TodosLosGrados.png"
        plt.savefig(os.path.join(PATH_OUTPUT_PLOTS, fname), dpi=300, bbox_inches='tight')
        plt.close()
        print(f"    [OK] {fname}")

# ==========================================
# 8. KDE + BOXPLOT JITTER
# ==========================================
def generar_kde_y_boxplot(df_master, meses_labels, sufijo_archivo="B1", color_list=None):
    if color_list is None:
        color_list = COLORES_MESES
    palette_meses = {mes: color_list[i % len(color_list)]
                     for i, mes in enumerate(meses_labels)}
    print(f"\n[*] Generando KDE y Boxplot+Jitter ({sufijo_archivo})...")

    for materia in df_master['Materia'].unique():
        df_sub = df_master[df_master['Materia'] == materia].copy()
        print(f"    - {materia}...")

        # KDE
        plt.figure(figsize=(12, 7))
        sns.histplot(data=df_sub, x=COL_SCORE_EXACTA, hue='Mes',
                     kde=True, element="step", stat="probability",
                     common_norm=False, palette=palette_meses, alpha=0.6)
        titulo = " vs ".join(
            [m.split(' ')[-1].replace('(', '').replace(')', '')
             for m in meses_labels])
        plt.title(f"Evolución de Puntajes Theta: {materia}\n"
                  f"({titulo} — {sufijo_archivo})")
        plt.xlim(0, 100)
        plt.savefig(
            os.path.join(
                PATH_OUTPUT_PLOTS,
                f"Distribucion_Evolutiva_{sufijo_archivo}_{materia}.png"),
            dpi=300, bbox_inches='tight')
        plt.close()

        # Boxplot + jitter
        meses_con_datos = [m for m in meses_labels
                           if not df_sub[df_sub['Mes'] == m].empty]
        n_meses   = len(meses_con_datos)
        fig, axes = plt.subplots(n_meses, 1,
                                 figsize=(12, 4 * n_meses),
                                 sharex=True)
        if n_meses == 1: axes = [axes]

        for i, mes_key in enumerate(meses_con_datos):
            ax        = axes[i]
            color     = color_list[i % len(color_list)]
            datos_box = (df_sub[df_sub['Mes'] == mes_key]
                         [COL_SCORE_EXACTA].dropna())
            if datos_box.empty:
                ax.set_visible(False); continue
            ax.boxplot(
                datos_box, vert=False, patch_artist=True, widths=0.45,
                boxprops=dict(facecolor=color, alpha=0.45, linewidth=1.5),
                medianprops=dict(color='#1a2b4c', linewidth=2.5),
                whiskerprops=dict(linewidth=1.5, linestyle='--',
                                  color='#555555'),
                capprops=dict(linewidth=2, color='#555555'),
                flierprops=dict(marker='', linestyle='none'),
                showmeans=True,
                meanprops=dict(marker='D', markerfacecolor='white',
                               markeredgecolor='#1a2b4c', markersize=8))
            np.random.seed(42)
            jitter = np.random.uniform(-0.18, 0.18, size=len(datos_box))
            ax.scatter(datos_box, 1 + jitter, alpha=0.20, s=10,
                       color=color, edgecolors='none', zorder=3)
            prueba_num = meses_labels.index(mes_key) + 1
            mes_nombre = (mes_key.split('(')[-1]
                          .replace(')', '').strip())
            ax.set_ylabel(f"Prueba {prueba_num}\n({mes_nombre})",
                          fontsize=11, fontweight='bold',
                          color='#005288', labelpad=10)
            ax.set_xlim(0, 100); ax.set_yticks([])
            ax.grid(axis='x', linestyle='--', alpha=0.4)
            ax.spines[['top', 'right', 'left']].set_visible(False)
            ax.text(0.99, 0.82,
                    f"n={len(datos_box):,}  |  "
                    f"Mediana={datos_box.median():.1f}  |  "
                    f"Media={datos_box.mean():.1f}",
                    transform=ax.transAxes, ha='right', va='top',
                    fontsize=10, color='#333333',
                    bbox=dict(boxstyle='round,pad=0.35',
                              facecolor='white',
                              edgecolor='#cccccc', alpha=0.85))

        axes[-1].set_xlabel("Puntaje (0–100)", fontsize=12)
        fig.suptitle(
            f"Distribución por Prueba — {sufijo_archivo} — {materia}\n"
            f"(todos los grados)",
            fontsize=15, fontweight='bold', color='#1a2b4c', y=1.01)
        plt.tight_layout()
        plt.savefig(
            os.path.join(
                PATH_OUTPUT_PLOTS,
                f"Boxplot_Jitter_{sufijo_archivo}_{materia}.png"),
            dpi=300, bbox_inches='tight')
        plt.close()

# ==========================================
# 9. B1 CONSOLIDADO
# ==========================================
def exportar_base_consolidada_b1(df_master, df_matricula, lookup_req027, meses_labels):
    print("\n[*] Generando Base de Datos Consolidada (Grupo B1)...")
    try:
        df_pivot = df_master.pivot_table(
            index=['Nro de Centro', 'Centro', 'Grupo',
                   'Documento', 'Grado_Str'],
            columns=['Materia', 'Mes'],
            values=COL_SCORE_EXACTA,
            aggfunc='first'
        )

        score_cols_raw = sorted(
            df_pivot.columns.tolist(),
            key=lambda col: (
                0 if 'Matemática' in col[0] else 1,
                next((i for i, m in enumerate(meses_labels)
                      if m in col[1]), 99)
            )
        )

        flat_map = {
            col: (f"{COL_SCORE_EXACTA} "
                  f"(para {col[0].lower()} {col[1].lower()})")
            for col in df_pivot.columns
        }

        df_pivot.columns = [flat_map[c] for c in df_pivot.columns]
        df_pivot = df_pivot.reset_index()
        df_pivot = df_pivot.rename(columns={'Grado_Str': 'Grado'})

        interleaved = []
        for raw_col in score_cols_raw:
            score_col = flat_map[raw_col]
            nivel_col = f"Nivel {raw_col[0]} ({raw_col[1]})"
            if score_col in df_pivot.columns:
                df_pivot[nivel_col] = (df_pivot[score_col]
                                       .apply(clasificar_puntaje))
                interleaved.append(score_col)
                interleaved.append(nivel_col)

        id_cols   = [c for c in ['Nro de Centro', 'Centro',
                                 'Grupo', 'Documento', 'Grado']
                     if c in df_pivot.columns]
        placed    = set(id_cols + interleaved)
        remaining = [c for c in df_pivot.columns if c not in placed]
        df_pivot  = df_pivot[id_cols + interleaved + remaining]

        cols_mat_priority = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN',
                             'GRADO', 'NIE', 'PRIMER_NOMBRE',
                             'SEGUNDO_NOMBRE', 'PRIMER_APELLIDO',
                             'SEGUNDO_APELLIDO']
        cols_exist = [c for c in cols_mat_priority
                      if c in df_matricula.columns
                      and c != 'CÓDIGO_SECCIÓN']
        df_mat_sub = df_matricula[cols_exist + ['NIE_Limpio']].copy()
        df_out = pd.merge(df_mat_sub, df_pivot,
                          left_on='NIE_Limpio', right_on='Documento',
                          how='inner')
        df_out = df_out.drop(columns=['NIE_Limpio'], errors='ignore')
        df_out = _aplicar_group_id(df_out, lookup_req027)

        header = [c for c in cols_mat_priority if c in df_out.columns]
        rest   = [c for c in df_out.columns if c not in set(header)]
        df_out = df_out[header + rest]

        df_out.to_excel(
            os.path.join(PATH_OUTPUT_PLOTS,
                         "Base_Estudiantes_B1_Consolidada.xlsx"),
            index=False)
        print("  [OK] Excel consolidado B1 guardado.")
    except Exception as e:
        print(f"  [!] Error al exportar la base B1: {e}")

# ==========================================
# 10. B2 EXCEL + BARRAS
# ==========================================
def generar_excel_y_barras_b2(df_b2_mes3, df_matricula, lookup_req027):
    if df_b2_mes3.empty:
        print("\n  [!] No hay estudiantes en el Grupo B2 para el Mes 3.")
        return
    print("\n[*] Generando Excel y Gráficas para el Grupo B2...")

    df_pivot = df_b2_mes3.pivot_table(
        index='Documento', columns='Materia',
        values=COL_SCORE_EXACTA, aggfunc='first')

    df_pivot.columns = [f"Puntaje {c}" for c in df_pivot.columns]
    df_pivot = df_pivot.reset_index()

    for base_col in ['Puntaje Matemática', 'Puntaje Lengua']:
        if base_col not in df_pivot.columns:
            df_pivot[base_col] = pd.NA
        nivel_col = base_col.replace('Puntaje', 'Nivel')
        df_pivot[nivel_col] = df_pivot[base_col].apply(clasificar_puntaje)

    cols_sol   = ['CODIGO', 'NOMBRE', 'GRADO', 'NIE',
                  'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                  'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO']
    cols_exist = [c for c in cols_sol if c in df_matricula.columns]
    df_mat_exp = df_matricula[cols_exist + ['NIE_Limpio']].copy()
    df_excel   = pd.merge(df_mat_exp, df_pivot,
                          left_on='NIE_Limpio', right_on='Documento',
                          how='inner')
    df_excel   = df_excel.drop(
        columns=['NIE_Limpio', 'Documento'], errors='ignore')
    df_excel   = _aplicar_group_id(df_excel, lookup_req027)

    priority = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO',
                'Puntaje Matemática', 'Nivel Matemática',
                'Puntaje Lengua',     'Nivel Lengua']
    final_cols = [c for c in priority if c in df_excel.columns]
    df_excel[final_cols].to_excel(
        os.path.join(PATH_OUTPUT_PLOTS,
                     "Resultados_Mes3_Grupo_B2.xlsx"), index=False)
    print("  [OK] Excel del Grupo B2 guardado.")

    mes_b2 = [df_b2_mes3['Mes'].iloc[0]]
    # Barras por grupos divididos
    generar_barras_por_grupo_grados(
        df_b2_mes3, mes_b2, sufijo_archivo="B2")

# ==========================================
# 11. EXCEL MAESTRO — HOJA ÚNICA
# ==========================================
def construir_excel_unico(df_merged, df_matricula, meses_labels, lookup_req027):
    print("\n[*] Construyendo Excel único "
          "(una sola hoja, todos los datos)...")

    df_b1 = df_merged[df_merged['GRUPO_Limpio'] == 'B1'].copy()
    if len(meses_labels) >= 3:
        base  = (df_b1[df_b1['Mes'].isin(meses_labels[:2])]
                 [['Documento', 'Materia']].drop_duplicates())
        df_b1 = pd.merge(df_b1, base,
                         on=['Documento', 'Materia'], how='inner')
    df_b1['Grupo_Paars'] = 'B1'

    df_b2 = df_merged[df_merged['GRUPO_Limpio'] == 'B2'].copy()
    if len(meses_labels) >= 3:
        df_b2 = df_b2[df_b2['Mes'] == meses_labels[2]]
    df_b2['Grupo_Paars'] = 'B2'

    df_all = pd.concat([df_b1, df_b2], ignore_index=True)
    if df_all.empty:
        print("  [!] No hay datos para el Excel único.")
        return None

    df_pivot = df_all.pivot_table(
        index=['Documento', 'Nro de Centro', 'Centro',
               'Grado_Str', 'GRUPO_Limpio', 'Grupo_Paars'],
        columns=['Materia', 'Mes'],
        values=COL_SCORE_EXACTA,
        aggfunc='first'
    )

    score_cols_raw = sorted(
        df_pivot.columns.tolist(),
        key=lambda col: (
            0 if 'Matemática' in col[0] else 1,
            next((i for i, m in enumerate(meses_labels)
                  if m in col[1]), 99)
        )
    )
    flat_map = {
        col: f"Puntaje {col[0]} ({col[1]})"
        for col in df_pivot.columns
    }
    df_pivot.columns = [flat_map[c] for c in df_pivot.columns]
    df_pivot = df_pivot.reset_index()

    interleaved = []
    for raw_col in score_cols_raw:
        score_col = flat_map[raw_col]
        nivel_col = f"Nivel {raw_col[0]} ({raw_col[1]})"
        if score_col in df_pivot.columns:
            df_pivot[nivel_col] = (df_pivot[score_col]
                                   .apply(clasificar_puntaje))
            interleaved.append(score_col)
            interleaved.append(nivel_col)

    cols_mat_priority = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO',
                         'NIE', 'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                         'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO']
    cols_exist = [c for c in cols_mat_priority
                  if c in df_matricula.columns
                  and c != 'CÓDIGO_SECCIÓN']
    df_mat_sub = df_matricula[cols_exist + ['NIE_Limpio']].copy()

    df_out = pd.merge(
        df_mat_sub, df_pivot,
        left_on='NIE_Limpio', right_on='Documento',
        how='inner'
    )
    df_out = df_out.drop(columns=['NIE_Limpio'], errors='ignore')
    df_out = _aplicar_group_id(df_out, lookup_req027)

    filter_cols = ['CODIGO', 'NOMBRE', 'CÓDIGO_SECCIÓN', 'GRADO', 'NIE',
                   'PRIMER_NOMBRE', 'SEGUNDO_NOMBRE',
                   'PRIMER_APELLIDO', 'SEGUNDO_APELLIDO',
                   'Grupo_Paars',
                   'Nro de Centro', 'Centro',
                   'Grado_Str', 'GRUPO_Limpio']
    header = [c for c in filter_cols if c in df_out.columns]
    placed = set(header + interleaved)
    rest   = [c for c in df_out.columns if c not in placed]
    df_out = df_out[[c for c in header + interleaved + rest
                     if c in df_out.columns]]

    out_path = os.path.join(PATH_OUTPUT_PLOTS,
                            "Excel_Maestro_HojaUnica.xlsx")
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name="Todos los Datos",
                        index=False)
        ws = writer.sheets["Todos los Datos"]
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes    = "A2"

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"  [OK] Excel único guardado → {out_path}  "
          f"({len(df_out):,} filas, {size_mb:.1f} MB)")
    return out_path


# ==========================================
# 12. EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print(f"\n{'='*60}")
    print("  GENERADOR DE REPORTE EVOLUTIVO — GRUPOS B1 Y B2")
    print(f"{'='*60}")

    preparar_entorno()
    meses_labels = [etiqueta for _, etiqueta in DIRECTORIOS_MESES]

    # 1. Cargar la Matrícula Maestra
    df_matricula = cargar_matricula()
    if df_matricula.empty:
        print("\n[!] El proceso se detuvo por falta de matrícula.")
        sys.exit()

    # ---> Mapa oficial de Grupos por Escuela basado en la Matrícula
    print("  [*] Generando mapa oficial de grupos (B1/B2) por código de centro...")
    if 'CODIGO' in df_matricula.columns:
        df_grupos_oficiales = df_matricula[['CODIGO', 'GRUPO_Limpio']].dropna().drop_duplicates(subset=['CODIGO'])
        mapa_grupos_escuela = df_grupos_oficiales.set_index('CODIGO')['GRUPO_Limpio']
    else:
        print("  [!] ERROR: La columna 'CODIGO' no existe en MatriculaProgresoMes3.")
        sys.exit()

    lookup_req027 = cargar_req027()

    # 2. Extraer puntajes
    df_all_scores = extraer_csvs_puntajes()
    if df_all_scores.empty:
        print("\n[!] No se extrajeron puntajes válidos.")
        sys.exit()

    # 3. Cruce COMPLETO (Sin filtros temporales)
    print("\n[*] Cruzando puntajes con la matrícula...")
    df_merged_full = pd.merge(
        df_all_scores,
        df_matricula[['NIE_Limpio']], 
        left_on='Documento', right_on='NIE_Limpio', how='inner')

    # ---> Asignar el GRUPO estrictamente según el código de la escuela
    df_merged_full['GRUPO_Limpio'] = df_merged_full['Nro de Centro'].map(mapa_grupos_escuela)
    df_merged_full['GRUPO_Limpio'] = df_merged_full['GRUPO_Limpio'].fillna('DESCONOCIDO')


    # 4. ---> Reportes de Conteo y Alertas por consola
    generar_reportes_conteos_separados(df_merged_full)
    alertar_b2_mes1_mes2(df_merged_full, meses_labels)


    # ---> INICIO DEL PROCESO ORIGINAL CON FILTROS
    print("\n[*] Aplicando filtros de trayectoria para reportes y gráficas...")
    
    # ── GRUPO B1 ─────────────────────────────────────────────────
    df_b1 = df_merged_full[df_merged_full['GRUPO_Limpio'] == 'B1'].copy()
    if df_b1.empty:
        print("  [!] No hay estudiantes de B1.")
    else:
        if len(meses_labels) >= 3:
            base  = (df_b1[df_b1['Mes'].isin(meses_labels[:2])]
                     [['Documento', 'Materia']].drop_duplicates())
            df_b1 = pd.merge(df_b1, base,
                             on=['Documento', 'Materia'], how='inner')
        print(f"  [OK] Estudiantes válidos B1: "
              f"{len(df_b1['Documento'].unique())}")

        exportar_base_consolidada_b1(
            df_b1, df_matricula, lookup_req027, meses_labels)
        dibujar_sankey_multimes(df_b1, "Matemática", meses_labels)
        dibujar_sankey_multimes(df_b1, "Lengua",     meses_labels)
        generar_kde_y_boxplot(df_b1, meses_labels,
                              sufijo_archivo="B1",
                              color_list=COLORES_MESES)
        generar_barras_por_grupo_grados(
            df_b1, meses_labels, sufijo_archivo="B1")
        
        # ---> BARRAS TODOS LOS GRADOS JUNTOS (B1)
        generar_barras_todos_los_grados(df_b1, meses_labels, sufijo_archivo="B1")

    # ── GRUPO B2 ─────────────────────────────────────────────────
    df_b2 = df_merged_full[df_merged_full['GRUPO_Limpio'] == 'B2'].copy()
    if not df_b2.empty and len(meses_labels) >= 3:
        df_b2_mes3 = df_b2[df_b2['Mes'] == meses_labels[2]].copy()
        generar_excel_y_barras_b2(
            df_b2_mes3, df_matricula, lookup_req027)
        generar_kde_y_boxplot(df_b2_mes3, [meses_labels[2]],
                              sufijo_archivo="B2",
                              color_list=["#b3b3ff"])
        
        # ---> BARRAS TODOS LOS GRADOS JUNTOS (B2 MES 3)
        generar_barras_todos_los_grados(df_b2_mes3, [meses_labels[2]], sufijo_archivo="B2")
        
    else:
        print("\n  [!] No se encontraron datos para B2 en el Mes 3.")

    # ── EXCEL MAESTRO HOJA ÚNICA ──────────────────────────────────
    construir_excel_unico(
        df_merged_full, df_matricula, meses_labels, lookup_req027)

    print(f"\n[OK] ¡Proceso completado! Todo está en:\n"
          f"     {PATH_OUTPUT_PLOTS}\n")