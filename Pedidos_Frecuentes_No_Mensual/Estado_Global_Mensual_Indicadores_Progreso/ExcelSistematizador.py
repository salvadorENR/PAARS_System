"""
Estado_Global_Indicadores_Progreso.py
======================================
Lee los HTMLs de Reporte_Nacional_Grados y los CSVs de resultados para
construir un único Excel consolidado CON los ítems ancla incluidos.

Para el Mes 3 (Mayo), donde el HTML fue generado SIN ítems ancla:
  - Los ítems ancla se recalculan desde los CSVs de resultados
  - Se incluyen en el Excel con la columna "Ítem_Ancla" = "Sí"
  - Se genera un nuevo HTML con los ítems ancla reincorporados

Fuentes:
  HTMLs      → directorio DIR_HTML  (mismo que DIR_SALIDA)
  CSVs       → .../2026/0N_PROGRESO_Mes/Interim_CSVs/Resultados/*.csv
  Metadata   → .../00_Metadata/Items_Progreso_MesN_2026_procesado.xlsx
  Anclas     → .../00_Metadata/items_anclaje_progreso_mes_N.xlsx

Salida:
  DIR_SALIDA/Estado_Global_Indicadores_Progreso.xlsx
  DIR_SALIDA/Reporte_Nacional_Grados_03_PROGRESO_Mayo_con_Anclas.html
"""

import os, re, glob, sys, subprocess

# =============================================================================
# AUTO-INSTALADOR
# =============================================================================
_REQUIRED = {"bs4": "beautifulsoup4", "pandas": "pandas", "openpyxl": "openpyxl"}

def _ensure_packages():
    missing = []
    for imp, pip in _REQUIRED.items():
        try:
            __import__(imp)
        except ImportError:
            missing.append(pip)
    if missing:
        print(f"[Setup] Instalando: {', '.join(missing)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + missing)
        print("[Setup] Listo.\n")

_ensure_packages()

from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. RUTAS
# =============================================================================
ROOT = (r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis"
        r"\PAARS_Warehouse")
PATH_META = os.path.join(ROOT, "00_Metadata")

# HTMLs y salida en el mismo directorio
DIR_HTML = DIR_SALIDA = (
    r"C:\Users\USUARIO.DESKTOP-KJ9DUHE.000\Documents\PAARS_System"
    r"\Pedidos_Frecuentes_No_Mensual\Estado_Global_Mensual_Indicadores_Progreso"
)
ARCHIVO_SALIDA = os.path.join(DIR_SALIDA, "Estado_Global_Indicadores_Progreso.xlsx")

# Meses: (num, etiqueta, subcarpeta_2026, nombre_html)
MESES = [
    (1, "Marzo", "01_PROGRESO_Marzo",
     "Reporte_Nacional_Grados_01_PROGRESO_Marzo.html"),
    (2, "Abril", "02_PROGRESO_Abril",
     "Reporte_Nacional_Grados_02_PROGRESO_Abril.html"),
    (3, "Mayo",  "03_PROGRESO_Mayo",
     "Reporte_Nacional_Grados_03_PROGRESO_Mayo.html"),
]

# Meses donde los ítems ancla fueron REMOVIDOS del HTML y deben reconstruirse
# desde los CSVs de resultados. Añadir el número de mes aquí cuando corresponda.
MESES_CON_ANCLAS_REMOVIDAS = {2, 3}

# =============================================================================
# 2. COLORES
# =============================================================================
CAT_STYLES = {
    "Crítico":   {"fill": "FFCDD2", "font": "B71C1C"},
    "Bajo":      {"fill": "FFE0B2", "font": "E65100"},
    "Medio":     {"fill": "FFF9C4", "font": "F57F17"},
    "Bueno":     {"fill": "DCEDC8", "font": "33691E"},
    "Excelente": {"fill": "C8E6C9", "font": "1B5E20"},
    "N/D":       {"fill": "F5F5F5", "font": "757575"},
    "Ancla":     {"fill": "E8EAF6", "font": "283593"},   # blue-grey for anchor items
}

HTML_COLORS = {
    "Crítico":   ("#fef2f2", "#991b1b"),
    "Bajo":      ("#fff7ed", "#ff8c2e"),
    "Medio":     ("#fefce8", "#facc15"),
    "Bueno":     ("#f7fee7", "#84cc16"),
    "Excelente": ("#ecfdf5", "#065f46"),
    "N/D":       ("#f9fafb", "#9ca3af"),
}

def clasificar(pct):
    if pct is None: return "N/D"
    if pct <= 20:   return "Crítico"
    if pct <= 40:   return "Bajo"
    if pct <= 60:   return "Medio"
    if pct <= 80:   return "Bueno"
    return "Excelente"

def normalize_code(s):
    return re.sub(r'[^A-Z0-9]', '', str(s).upper()) if s else ""

# =============================================================================
# 3. CARGA DE METADATA
# =============================================================================
def cargar_metadata(mes_num):
    """
    Returns two dicts:
      by_item_code : { NORMALIZED_ITEM_CODE : {objetivo, clases, orden} }
      by_obj_code  : { 'X.Y' objective code  : clases_str }

    by_obj_code is used to enrich Clase when the HTML has no badge spans
    (e.g. Mes 2 was generated without badge spans but the metadata file
    contains the Clases_Sugeridas column keyed by objective code).
    """
    by_item_code = {}
    by_obj_code  = {}

    patterns = [
        f"Items_Progreso_Mes{mes_num}_*procesado*.xlsx",
        f"Items_Progreso_Mes{mes_num}*.xlsx",
        f"*Progreso*Mes*{mes_num}*procesado*.xlsx",
        f"*Progreso*{mes_num}*.xlsx",
    ]
    found = None
    for pat in patterns:
        hits = [f for f in glob.glob(os.path.join(PATH_META, pat))
                if not os.path.basename(f).startswith('~$')]
        if hits:
            found = hits[0]
            break

    if not found:
        print(f"  ⚠  No se encontró metadata para Mes {mes_num} en {PATH_META}")
        return by_item_code, by_obj_code

    print(f"  -> Metadata: {os.path.basename(found)}")
    try:
        df = pd.read_excel(found, dtype=str)
        df.columns = df.columns.str.strip()
        print(f"     Columnas disponibles: {list(df.columns)}")

        # Column detection mirrors report_profesores.py exactly.
        # Primary exact names: ItemCodigo, indicador_logro, Clases_Sugeridas, Orden_Clase
        # Fallback: broad keyword search for files with different naming conventions.
        col_item = next((c for c in df.columns
                         if c == 'ItemCodigo'), None) or \
                   next((c for c in df.columns
                         if 'itemcodigo' in c.lower().replace(' ','')
                         or c.lower() in ('item', 'código', 'codigo', 'itemcod')), None) or \
                   next((c for c in df.columns
                         if any(k in c.lower() for k in ('item', 'código', 'codigo', 'cod_'))), None)

        col_ind  = next((c for c in df.columns
                         if c == 'indicador_logro'), None) or \
                   next((c for c in df.columns
                         if 'indicador_logro' in c.lower().replace(' ','_')), None) or \
                   next((c for c in df.columns
                         if 'indicador' in c.lower() or 'logro' in c.lower()), None)

        col_clase = next((c for c in df.columns
                          if c == 'Clases_Sugeridas'), None) or \
                    next((c for c in df.columns
                          if 'clases_sugeridas' in c.lower().replace(' ','_')), None) or \
                    next((c for c in df.columns
                          if 'sugerida' in c.lower() or 'clase' in c.lower()), None)

        col_orden = next((c for c in df.columns
                          if c == 'Orden_Clase'), None) or \
                    next((c for c in df.columns
                          if 'orden_clase' in c.lower().replace(' ','_')
                          or 'orden' in c.lower()), None)

        print(f"     Mapeado → col_item={col_item}  col_ind={col_ind}  "
              f"col_clase={col_clase}  col_orden={col_orden}")

        if not col_ind:
            print(f"  ⚠  Columna de indicador no encontrada.")
            return by_item_code, by_obj_code

        for _, row in df.iterrows():
            # Skip rows where both item code and indicator are empty
            ind_text = str(row[col_ind]).strip() if pd.notna(row.get(col_ind)) else ""
            if not ind_text or ind_text == 'nan':
                continue

            clases = ""
            if col_clase and pd.notna(row.get(col_clase)):
                v = str(row[col_clase]).strip()
                if v and v != 'nan':
                    clases = v

            orden = 9999.0
            if col_orden and pd.notna(row.get(col_orden)):
                try:
                    orden = float(row[col_orden])
                except ValueError:
                    pass

            # Index 1: by normalized item code (e.g. MAT211 → for anchor lookup)
            if col_item and pd.notna(row.get(col_item)):
                code = normalize_code(row[col_item])
                if code:
                    by_item_code[code] = {
                        "objetivo": ind_text,
                        "clases":   clases,
                        "orden":    orden,
                    }

            # Index 2: by objective code extracted from indicator text (e.g. '5.4')
            # Used to enrich Clase when HTML has no badge spans
            obj_match = re.search(r'\b(\d+\.\d+)\b', ind_text)
            if obj_match:
                obj_key = obj_match.group(1)
                # If multiple items share the same obj code, concat their clases
                if obj_key in by_obj_code:
                    if clases and clases not in by_obj_code[obj_key]:
                        by_obj_code[obj_key] += f" | {clases}"
                elif clases:
                    by_obj_code[obj_key] = clases

        print(f"  -> {len(by_item_code)} entradas por código ítem, "
              f"{len(by_obj_code)} entradas por código objetivo")
    except Exception as e:
        print(f"  ⚠  Error leyendo metadata: {e}")
    return by_item_code, by_obj_code


def cargar_anclas(mes_num):
    """
    Loads the anchor-item exclusion list and returns:
        { grade_str: set_of_normalized_codes }
    e.g. { '2': {'LEC8','LEC9','MAT11',...}, '3': {...}, ... }

    Automatically detects and handles TWO different file formats:

    FORMAT A — Wide (Mes 3 style):
        One sheet per subject ('Matematica', 'Lectura').
        One COLUMN per grade ('2deg', '3deg', ...).
        Each cell = one item code.
        Detection: any column header contains a digit (grade number).

    FORMAT B — Long/Tall (Mes 2 style):
        Single sheet ('Hoja1' or similar).
        Two columns: 'PruebaTitulo' and 'ItemCodigo'.
        PruebaTitulo encodes both subject and grade:
            e.g. 'Prueba de Lectura 10° Mes 2', 'Prueba de Matemática 2° Mes 2'
        Detection: columns named PruebaTitulo / ItemCodigo (case-insensitive).

    File search order (both formats):
        1. items_anclaje_progreso_mes_N.xlsx  (canonical name)
        2. Any file matching *anclaje*mes*N*.xlsx
        3. Any file matching *anclaje*N*.xlsx  (looser fallback)
        4. Listado*ancla*N*.xlsx  (alternative naming convention like Mes 2)
        5. Listado*anclaje*N*.xlsx
    """
    anclas = {}

    patterns = [
        f"items_anclaje_progreso_mes_{mes_num}.xlsx",
        f"items_anclaje*mes*{mes_num}*.xlsx",
        f"items_anclaje*{mes_num}*.xlsx",
        f"*anclaje*mes*{mes_num}*.xlsx",
        f"*anclaje*{mes_num}*.xlsx",
        f"Listado*ancla*{mes_num}*.xlsx",
        f"Listado*anclaje*{mes_num}*.xlsx",
        f"*ancla*progreso*{mes_num}*.xlsx",
        f"*ancla*{mes_num}*.xlsx",
    ]
    found = None
    for pat in patterns:
        hits = [f for f in glob.glob(os.path.join(PATH_META, pat))
                if not os.path.basename(f).startswith('~$')]
        if hits:
            found = hits[0]
            break

    if not found:
        print(f"  ⚠  No se encontró archivo de ítems ancla para Mes {mes_num}")
        print(f"     Buscado en: {PATH_META}")
        print(f"     Formatos aceptados:")
        print(f"       • items_anclaje_progreso_mes_{mes_num}.xlsx  (formato wide, una col por grado)")
        print(f"       • Listado_ítems_ancla_progreso_{mes_num}.xlsx  (formato long, col PruebaTitulo+ItemCodigo)")
        return anclas

    print(f"  -> Anclas: {os.path.basename(found)}")
    try:
        xl  = pd.ExcelFile(found)

        for sheet_name in xl.sheet_names:
            df_s = xl.parse(sheet_name, dtype=str)
            df_s.columns = df_s.columns.str.strip()

            # ── Detect format ─────────────────────────────────────────────────
            col_titulo = next(
                (c for c in df_s.columns
                 if 'prueba' in c.lower() or 'titulo' in c.lower()
                 or 'title' in c.lower()), None
            )
            col_codigo = next(
                (c for c in df_s.columns
                 if 'itemcodigo' in c.lower().replace(' ','').replace('_','')
                 or c.lower() in ('itemcodigo', 'item', 'codigo', 'code', 'itemcode')), None
            )

            if col_titulo and col_codigo:
                # ── FORMAT B: Long/Tall ────────────────────────────────────────
                # PruebaTitulo = 'Prueba de Lectura 10° Mes 2'
                # Extract grade number from the degree symbol pattern: '10°'
                for _, row in df_s.dropna(subset=[col_titulo, col_codigo]).iterrows():
                    titulo = str(row[col_titulo])
                    code   = normalize_code(str(row[col_codigo]))
                    if not code:
                        continue
                    grade_m = re.search(r'(\d+)\s*°', titulo)
                    if not grade_m:
                        # fallback: last number before 'Mes' or in the string
                        grade_m = re.search(r'(\d+)\s*(?:°|Mes|$)', titulo)
                    if grade_m:
                        grade_key = grade_m.group(1)
                        if grade_key not in anclas:
                            anclas[grade_key] = set()
                        anclas[grade_key].add(code)

            else:
                # ── FORMAT A: Wide ─────────────────────────────────────────────
                # Each column header = grade label ('2deg', '10deg', ...)
                # Each cell = item code
                for col in df_s.columns:
                    grade_key = re.sub(r'\D', '', str(col).strip())
                    if not grade_key:
                        continue
                    if grade_key not in anclas:
                        anclas[grade_key] = set()
                    for val in df_s[col].dropna():
                        code = normalize_code(str(val))
                        if code:
                            anclas[grade_key].add(code)

        total = sum(len(v) for v in anclas.values())
        print(f"  -> {total} ítems ancla en {len(anclas)} grado(s)")
        for g in sorted(anclas, key=lambda x: int(x) if x.isdigit() else 99):
            lec = sorted(c for c in anclas[g] if c.startswith('LEC'))
            mat = sorted(c for c in anclas[g] if c.startswith('MAT'))
            print(f"     Grado {g}: {len(lec)} LEC  {len(mat)} MAT")

    except Exception as e:
        print(f"  ⚠  Error leyendo anclas: {e}")
    return anclas


def cargar_csvs(mes_num, subcarpeta):
    """Loads and concatenates all result CSVs for a month."""
    ruta = os.path.join(ROOT, "2026", subcarpeta, "Interim_CSVs", "Resultados")
    archivos = glob.glob(os.path.join(ruta, "*.csv"))
    if not archivos:
        print(f"  ⚠  No se encontraron CSVs en {ruta}")
        return pd.DataFrame()
    dfs = []
    for f in archivos:
        try:
            df = pd.read_csv(f, dtype=str, encoding_errors='ignore')
            df.columns = df.columns.str.strip()
            df['_area'] = 'Matemática' if 'MAT' in os.path.basename(f).upper() else 'Lengua'
            dfs.append(df)
        except Exception as e:
            print(f"  ⚠  Error leyendo CSV {os.path.basename(f)}: {e}")
    if dfs:
        combined = pd.concat(dfs, ignore_index=True)
        print(f"  -> CSVs cargados: {len(archivos)} archivos, {len(combined)} filas")
        return combined
    return pd.DataFrame()

# =============================================================================
# 4. HELPERS — LEC fingerprint (ported from report_por_grados.py)
# =============================================================================
def make_lec_fingerprint(texto):
    """
    Stable grouping key for a LEC indicator text.
    Mirrors report_por_grados.py exactly:
      - Priority: objective code (e.g. '4.3') → all items sharing the same
        X.Y code are merged into ONE indicator row.
      - Fallback: normalised stripped text (accent-free, alphanumeric only).
    This prevents duplicate rows like '6.3 Compara...' appearing 5 times
    with different percentages — they get pooled into a single averaged row.
    """
    import unicodedata as _ud
    if not texto or texto.startswith('Indicador no disponible'):
        return f'missing_{texto}'
    code_match = re.search(r'\b(\d+\.\d+)\b', texto)
    if code_match:
        return f"obj_{code_match.group(1)}"
    norm = _ud.normalize('NFD', texto.lower())
    norm = ''.join(c for c in norm if _ud.category(c) != 'Mn')
    norm = re.sub(r'[\W_]+', '', norm)
    return norm[:60]


# =============================================================================
# 5. PARSER DEL HTML
# =============================================================================
def parsear_html(ruta_html, mes_num, mes_nombre, by_obj_code=None):
    """
    Parses and consolidates indicator rows from the HTML.

    For LENGUA: multiple HTML items that share the same objective code (e.g.
    '6.3') are merged into ONE row using the same make_lec_fingerprint() logic
    as report_por_grados.py. The merged % is the simple average of the
    individual pcts (valid because all items in a grade have the same N).
    Clase badges are union-joined with ' | '.

    For MATEMÁTICA: each item is its own indicator (no merging).

    by_obj_code : dict { 'X.Y' → clases_str } used to fill Clase when the
                  HTML has no badge spans (e.g. Mes 2 was generated without them).
    """
    if by_obj_code is None:
        by_obj_code = {}

    raw_rows     = []   # one entry per HTML item-div (before merging)
    badges_found = 0

    with open(ruta_html, encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    for tab in soup.find_all('div', class_='tab-content'):
        tab_id     = tab.get('id', '')
        asignatura = 'Matemática' if tab_id == 'MAT' else 'Lengua'

        for acc in tab.find_all('details', class_='grade-accordion'):
            g_match = re.search(r'Grado\s+(\d+)', acc.find('summary').get_text())
            grado   = int(g_match.group(1)) if g_match else 0

            for item_div in acc.find_all('div', class_='indicator-item'):
                pct_text  = item_div.find('div', class_='pct-box').get_text(' ', strip=True)
                pct_match = re.search(r'([\d.]+)%', pct_text)
                pct       = float(pct_match.group(1)) if pct_match else None

                ind_div = item_div.find('div', class_='ind-text')
                badge   = ind_div.find('span', class_='badge-class')

                if badge:
                    clase_txt = badge.get_text(strip=True)
                    badge.decompose()
                    badges_found += 1
                else:
                    clase_txt = ""

                indicador = ind_div.get_text(strip=True)

                # No badge → enrich Clase from metadata by objective code
                if not clase_txt and by_obj_code:
                    obj_m = re.search(r'\b(\d+\.\d+)\b', indicador)
                    if obj_m:
                        clase_txt = by_obj_code.get(obj_m.group(1), "")

                raw_rows.append({
                    "asignatura": asignatura,
                    "grado":      grado,
                    "clase_txt":  clase_txt,
                    "indicador":  indicador,
                    "pct":        pct,
                })

    # ── Merge duplicate LEC rows that share the same objective code ──────────
    # Groups are keyed by (asignatura, grado, fingerprint).
    # MAT uses codigo_puro as fingerprint → no merging.
    # LEC uses make_lec_fingerprint() → items with the same X.Y code merge.

    groups = {}   # key → { display_text, clases, pcts, asignatura, grado }

    for r in raw_rows:
        if r['asignatura'] == 'Lengua':
            fp = (r['asignatura'], r['grado'],
                  make_lec_fingerprint(r['indicador']))
        else:
            # MAT: unique per indicator text (no merging)
            fp = (r['asignatura'], r['grado'], r['indicador'])

        if fp not in groups:
            groups[fp] = {
                'asignatura':    r['asignatura'],
                'grado':         r['grado'],
                'display_text':  r['indicador'],
                'clases':        set(),
                'pcts':          [],
            }

        g = groups[fp]
        if r['pct'] is not None:
            g['pcts'].append(r['pct'])
        if r['clase_txt']:
            # Split multi-badge strings and add each part separately
            for part in re.split(r'\s*\|\s*', r['clase_txt']):
                if part.strip():
                    g['clases'].add(part.strip())
        # Keep the first (lowest-pct, since HTML is sorted asc) display text
        # — or prefer the one without a period-dot issue
        if r['indicador'] and not g['display_text']:
            g['display_text'] = r['indicador']

    # ── Build final rows ─────────────────────────────────────────────────────
    rows         = []
    merged_count = 0

    for fp, g in groups.items():
        if not g['pcts']:
            continue

        n_items = len(g['pcts'])
        if n_items > 1:
            merged_count += 1

        # Pool: simple average (equivalent to raw-count pool when N is equal per item)
        pct_merged = round(sum(g['pcts']) / n_items, 1)
        cat        = clasificar(pct_merged)

        # Rebuild Clase string: sort badges naturally
        clases_sorted = sorted(g['clases'],
                               key=lambda x: [int(t) if t.isdigit() else t
                                              for t in re.split(r'(\d+)', x)])
        clase_str = " | ".join(clases_sorted)

        rows.append({
            "Mes_Num":     mes_num,
            "Mes":         mes_nombre,
            "Asignatura":  g['asignatura'],
            "Grado":       g['grado'],
            "Clase":       clase_str,
            "Indicador":   g['display_text'],
            "Pct_Acierto": pct_merged,
            "Categoria":   cat,
            "Item_Ancla":  "No",
            "_n_items_merged": n_items,   # internal, stripped before export
        })

    enrichment    = "con badges HTML" if badges_found > 0 else "SIN badges → enriquecido desde metadata"
    clases_filled = sum(1 for r in rows if r['Clase'])
    raw_n         = len(raw_rows)
    final_n       = len(rows)

    print(f"  -> HTML parseado: {raw_n} ítems HTML → {final_n} indicadores únicos ({enrichment})")
    if merged_count:
        print(f"     ⚠  {merged_count} grupo(s) de LEC con código objetivo compartido fueron "
              f"consolidados (promedio de pcts).")
    print(f"     Clase disponible en: {clases_filled}/{final_n} indicadores")
    return rows

# =============================================================================
# 5. RECONSTRUCCIÓN DE ÍTEMS ANCLA DESDE CSVs
# =============================================================================
def reconstruir_anclas(df_csv, by_item_code, anclas_por_grado, mes_num, mes_nombre):
    """
    For each anchor item code, calculates % correct from CSVs.

    LEC anchor items that share the same objective code fingerprint are MERGED
    (averaged) exactly like parsear_html does for regular LEC items, so the
    regenerated HTML and Excel don't show duplicate anchor rows or
    'Indicador no disponible' entries when metadata is available.

    If by_item_code is empty (metadata file not found on this machine), the
    label will show the item code clearly, e.g. "LEC315 — indicador pendiente
    de metadata" instead of the cryptic "Indicador no disponible (LEC315)".
    """
    rows = []
    if df_csv.empty:
        return rows

    col_grado = next((c for c in df_csv.columns if 'grado' in c.lower()), None)

    # ── Step 1: compute raw pct per anchor code ───────────────────────────────
    raw_anchor_rows = []   # {code, area, grado_num, pct, indicador, clase_txt}

    for grade_key, codes in anclas_por_grado.items():
        grado_num = int(grade_key)

        if col_grado:
            df_g = df_csv[
                df_csv[col_grado].astype(str).str.extract(r'(\d+)')[0] == grade_key
            ].copy()
        else:
            df_g = df_csv.copy()

        if df_g.empty:
            continue

        for code in sorted(codes):
            area    = 'Matemática' if code.startswith('MAT') else 'Lengua'
            df_area = df_g[df_g['_area'] == area]

            col_match = next(
                (c for c in df_area.columns if normalize_code(c) == code),
                None
            )
            if not col_match:
                continue

            serie = pd.to_numeric(df_area[col_match], errors='coerce').dropna()
            if serie.empty:
                continue

            pct = round((serie == 1).sum() / len(serie) * 100, 1)

            meta      = by_item_code.get(code, {})
            indicador = meta.get('objetivo', None)
            clase_txt = meta.get('clases', '')

            # If metadata not found, use a clear informative label
            if not indicador:
                indicador = f"{code} — indicador pendiente de metadata"

            raw_anchor_rows.append({
                'code':      code,
                'area':      area,
                'grado_num': grado_num,
                'pct':       pct,
                'indicador': indicador,
                'clase_txt': clase_txt,
            })

    # ── Step 2: merge LEC anchor items sharing the same objective fingerprint ─
    groups = {}   # (area, grado_num, fingerprint) → group dict

    for r in raw_anchor_rows:
        if r['area'] == 'Lengua':
            fp = (r['area'], r['grado_num'], make_lec_fingerprint(r['indicador']))
        else:
            fp = (r['area'], r['grado_num'], r['code'])   # MAT: one per code

        if fp not in groups:
            groups[fp] = {
                'area':         r['area'],
                'grado_num':    r['grado_num'],
                'display_text': r['indicador'],
                'clases':       set(),
                'pcts':         [],
            }

        g = groups[fp]
        g['pcts'].append(r['pct'])
        if r['clase_txt']:
            for part in re.split(r'\s*\|\s*', r['clase_txt']):
                if part.strip():
                    g['clases'].add(part.strip())
        # Prefer indicador text that is NOT a fallback code
        if ('pendiente de metadata' not in r['indicador']
                and 'pendiente de metadata' in g['display_text']):
            g['display_text'] = r['indicador']

    # ── Step 3: build final anchor rows ───────────────────────────────────────
    for fp, g in groups.items():
        if not g['pcts']:
            continue

        pct_merged = round(sum(g['pcts']) / len(g['pcts']), 1)
        cat        = clasificar(pct_merged)

        clases_sorted = sorted(
            g['clases'],
            key=lambda x: [int(t) if t.isdigit() else t
                           for t in re.split(r'(\d+)', x)]
        )
        clase_str = " | ".join(clases_sorted)

        rows.append({
            "Mes_Num":     mes_num,
            "Mes":         mes_nombre,
            "Asignatura":  g['area'],
            "Grado":       g['grado_num'],
            "Clase":       clase_str,
            "Indicador":   g['display_text'],
            "Pct_Acierto": pct_merged,
            "Categoria":   cat,
            "Item_Ancla":  "Sí",
        })

    n_meta_missing = sum(1 for r in rows
                         if 'pendiente de metadata' in r['Indicador'])
    print(f"  -> Ítems ancla reconstruidos: {len(rows)} indicadores únicos")
    if n_meta_missing:
        print(f"     ⚠  {n_meta_missing} ítem(s) ancla sin metadata "
              f"(el archivo Items_Progreso_Mes3_2026_procesado.xlsx "
              f"debe estar en {PATH_META})")
    return rows

# =============================================================================
# 6. REGENERAR HTML CON ÍTEMS ANCLA REINCORPORADOS
# =============================================================================
def regenerar_html_con_anclas(ruta_html_original, filas_ancla, ruta_salida):
    """
    Injects anchor items back into the Mayo HTML accordion sections.
    Anchor items are visually distinct (dashed border, blue-grey background).
    """
    with open(ruta_html_original, encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    # Add anchor item CSS to <style>
    style_tag = soup.find('style')
    if style_tag:
        style_tag.string = (style_tag.string or '') + """
            .indicator-item.ancla {
                border-left: 6px dashed #3949AB !important;
                background-color: #E8EAF6 !important;
                opacity: 0.88;
            }
            .ancla-badge {
                display: inline-block;
                background: #3949AB;
                color: white;
                font-size: 10px;
                font-weight: bold;
                padding: 2px 7px;
                border-radius: 3px;
                margin-right: 6px;
                letter-spacing: 0.4px;
                vertical-align: middle;
            }
        """

    # Update header to note anchor items included
    h1 = soup.find('h1')
    if h1:
        h1.string = h1.get_text(strip=True) + " (con Ítems Ancla)"

    # Build lookup: (tab_id, grado) → list of anchor rows
    from collections import defaultdict
    ancla_lookup = defaultdict(list)
    for row in filas_ancla:
        tab_id = 'MAT' if row['Asignatura'] == 'Matemática' else 'LEC'
        ancla_lookup[(tab_id, row['Grado'])].append(row)

    for tab in soup.find_all('div', class_='tab-content'):
        tab_id = tab.get('id', '')
        for acc in tab.find_all('details', class_='grade-accordion'):
            g_match = re.search(r'Grado\s+(\d+)', acc.find('summary').get_text())
            grado   = int(g_match.group(1)) if g_match else 0
            content = acc.find('div', class_='content')
            if not content:
                continue

            anclas_para_este = ancla_lookup.get((tab_id, grado), [])
            if not anclas_para_este:
                continue

            # Sort anchor items by pct ascending (same as existing items)
            anclas_para_este.sort(key=lambda x: x['Pct_Acierto'] or 0)

            for arow in anclas_para_este:
                pct = arow['Pct_Acierto']
                cat = arow['Categoria']
                bg, border = HTML_COLORS.get(cat, ("#E8EAF6", "#3949AB"))
                clase_txt  = arow['Clase']
                badge_html = (f'<span class="badge-class">{clase_txt}</span> '
                              if clase_txt else '')
                new_div = BeautifulSoup(f"""
                <div class="indicator-item ancla"
                     style="background-color: {bg}; border-left: 6px dashed #3949AB;">
                    <div class="pct-box">
                        <strong>{pct:.1f}%</strong><br>
                        <small>{cat}</small>
                    </div>
                    <div class="ind-text">
                        <span class="ancla-badge">ANCLA</span>
                        {badge_html}{arow['Indicador']}
                    </div>
                </div>""", 'html.parser')
                content.append(new_div)

    # Update legend note
    legend = soup.find('div', class_='legend')
    if legend:
        note = soup.new_tag('span',
                            style="margin-left:15px;color:#3949AB;font-weight:bold;")
        note.string = "· 🔵 Ítems con borde azul punteado = Ítems Ancla"
        legend.append(note)

    with open(ruta_salida, 'w', encoding='utf-8') as f:
        f.write(str(soup))

    print(f"  -> HTML regenerado: {os.path.basename(ruta_salida)}")

# =============================================================================
# 7. GENERADOR EXCEL
# =============================================================================
def make_border():
    s = Side(style='thin', color='DDDDDD')
    return Border(left=s, right=s, top=s, bottom=s)

def cell_style(cell, fill_hex, font_hex, bold=False, size=10,
               halign='center', wrap=False):
    cell.fill      = PatternFill(start_color=fill_hex, end_color=fill_hex,
                                  fill_type='solid')
    cell.font      = Font(color=font_hex, bold=bold, size=size)
    cell.alignment = Alignment(horizontal=halign, vertical='center',
                                wrap_text=wrap)
    cell.border    = make_border()

def aplicar_cat(ws, row_num, col_num, categoria, is_ancla=False):
    key   = 'Ancla' if is_ancla else categoria
    style = CAT_STYLES.get(key, CAT_STYLES['N/D'])
    cell  = ws.cell(row=row_num, column=col_num)
    cell.fill      = PatternFill(start_color=style['fill'],
                                  end_color=style['fill'], fill_type='solid')
    cell.font      = Font(color=style['font'], bold=True, size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border    = make_border()


def generar_excel(df_all):
    os.makedirs(DIR_SALIDA, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    HDR_DARK   = "1F3864"
    HDR_MAT    = "1A5276"
    HDR_LEC    = "6E2F4A"
    HDR_ANCLA  = "283593"
    WHITE_BOLD = Font(color="FFFFFF", bold=True, size=10)
    CENTER     = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LEFT_WRAP  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    THIN       = make_border()

    # ── Hoja 1: Datos Completos ───────────────────────────────────────────────
    ws1 = wb.create_sheet("Datos_Completos")
    ws1.freeze_panes = "A2"
    hdrs  = ["Mes_Num","Mes","Asignatura","Grado","Clase",
             "Indicador","% Acierto","Categoría","Ítem Ancla"]
    widths= [8, 10, 13, 7, 18, 80, 11, 12, 11]

    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.fill = PatternFill(start_color=HDR_DARK, end_color=HDR_DARK,
                              fill_type='solid')
        c.font = WHITE_BOLD; c.alignment = CENTER; c.border = THIN
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.row_dimensions[1].height = 32

    for ri, (_, row) in enumerate(df_all.iterrows(), 2):
        is_ancla = str(row.get('Item_Ancla', 'No')).strip() == 'Sí'
        vals = [row['Mes_Num'], row['Mes'], row['Asignatura'], row['Grado'],
                row['Clase'], row['Indicador'],
                row['Pct_Acierto'], row['Categoria'],
                row.get('Item_Ancla', 'No')]
        for ci, val in enumerate(vals, 1):
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.border    = THIN
            cell.alignment = LEFT_WRAP if ci == 6 else CENTER
            cell.font      = Font(size=10)
            if ci == 7 and val is not None:
                cell.number_format = "0.0"
            if is_ancla:
                cell.fill = PatternFill(start_color="E8EAF6",
                                         end_color="E8EAF6", fill_type='solid')
        aplicar_cat(ws1, ri, 8, row['Categoria'], is_ancla)

    ws1.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}1"

    # ── Hoja 2 & 3: Pivot por asignatura ─────────────────────────────────────
    for asig, sheet_nm, hdr_color in [
        ("Matemática", "Progreso_MAT_por_Indicador", HDR_MAT),
        ("Lengua",     "Progreso_LEC_por_Indicador", HDR_LEC),
    ]:
        df_s = df_all[df_all['Asignatura'] == asig].copy()
        if df_s.empty:
            continue

        ws_p = wb.create_sheet(sheet_nm)
        ws_p.freeze_panes = "D2"

        meses_pres = sorted(df_s['Mes_Num'].unique())
        mes_label  = {r['Mes_Num']: r['Mes']
                      for _, r in df_s[['Mes_Num','Mes']].drop_duplicates().iterrows()}

        df_s['_key'] = (df_s['Grado'].astype(str) + '||'
                        + df_s['Clase'] + '||'
                        + df_s['Indicador'] + '||'
                        + df_s.get('Item_Ancla', 'No').astype(str))
        pivot_rows = (df_s[['_key','Grado','Clase','Indicador','Item_Ancla']]
                      .drop_duplicates('_key')
                      .sort_values(['Grado','Item_Ancla','Clase','Indicador'])
                      .reset_index(drop=True))

        pct_map = {(r['_key'], r['Mes_Num']): r['Pct_Acierto']
                   for _, r in df_s.iterrows()}
        cat_map = {(r['_key'], r['Mes_Num']): r['Categoria']
                   for _, r in df_s.iterrows()}

        fixed_hdrs = ["Grado", "Clase", "Indicador", "Ítem Ancla"]
        fixed_ws   = [7, 18, 80, 11]
        mes_hdrs   = [f"Mes {mn} — {mes_label[mn]}" for mn in meses_pres]
        all_hdrs   = fixed_hdrs + mes_hdrs
        all_ws     = fixed_ws + [16] * len(meses_pres)

        hdr_fill = PatternFill(start_color=hdr_color, end_color=hdr_color,
                                fill_type='solid')
        for ci, (h, w) in enumerate(zip(all_hdrs, all_ws), 1):
            c = ws_p.cell(row=1, column=ci, value=h)
            c.fill = hdr_fill; c.font = WHITE_BOLD
            c.alignment = CENTER; c.border = THIN
            ws_p.column_dimensions[get_column_letter(ci)].width = w
        ws_p.row_dimensions[1].height = 32

        for ri, prow in pivot_rows.iterrows():
            er       = ri + 2
            key      = prow['_key']
            is_ancla = str(prow.get('Item_Ancla','No')).strip() == 'Sí'
            row_bg   = PatternFill(start_color="E8EAF6", end_color="E8EAF6",
                                    fill_type='solid') if is_ancla else None

            fixed_vals = [prow['Grado'], prow['Clase'],
                          prow['Indicador'], prow.get('Item_Ancla','No')]
            for ci, val in enumerate(fixed_vals, 1):
                cell = ws_p.cell(row=er, column=ci, value=val)
                cell.border    = THIN
                cell.alignment = LEFT_WRAP if ci == 3 else CENTER
                cell.font      = Font(size=10,
                                       color="283593" if is_ancla else "000000",
                                       italic=is_ancla)
                if row_bg:
                    cell.fill = row_bg

            for m_off, mn in enumerate(meses_pres):
                ci   = len(fixed_hdrs) + m_off + 1
                pct  = pct_map.get((key, mn))
                cat  = cat_map.get((key, mn), 'N/D')
                cell = ws_p.cell(row=er, column=ci,
                                  value=pct if pct is not None else '—')
                if pct is not None:
                    cell.number_format = "0.0"
                aplicar_cat(ws_p, er, ci, cat, is_ancla)

        ws_p.auto_filter.ref = f"A1:{get_column_letter(len(all_hdrs))}1"

    # ── Hoja 4: Resumen por categoría ────────────────────────────────────────
    ws_r = wb.create_sheet("Resumen_Categorias")
    ws_r.freeze_panes = "A2"
    cats     = ["Crítico","Bajo","Medio","Bueno","Excelente"]
    res_hdrs = ["Mes","Asignatura","Grado","Ítem Ancla"] + cats + ["Total"]

    hdr_fill_r = PatternFill(start_color=HDR_DARK, end_color=HDR_DARK,
                              fill_type='solid')
    for ci, h in enumerate(res_hdrs, 1):
        c = ws_r.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill_r; c.font = WHITE_BOLD
        c.alignment = CENTER; c.border = THIN
        ws_r.column_dimensions[get_column_letter(ci)].width = (
            12 if h in ["Mes","Grado","Ítem Ancla"] else
            14 if h in cats else 18
        )
    ws_r.row_dimensions[1].height = 30

    summary = (df_all.groupby(
                   ['Mes','Asignatura','Grado',
                    df_all.get('Item_Ancla', pd.Series(['No']*len(df_all),
                                                        name='Item_Ancla')),
                    'Categoria'])
               .size().unstack(fill_value=0).reset_index())
    for cat in cats:
        if cat not in summary.columns:
            summary[cat] = 0
    summary['Total'] = summary[cats].sum(axis=1)
    summary = (summary
               .merge(df_all[['Mes_Num','Mes']].drop_duplicates(), on='Mes', how='left')
               .sort_values(['Mes_Num','Asignatura','Grado'])
               .reset_index(drop=True))

    for ri, row in summary.iterrows():
        er       = ri + 2
        is_ancla = str(row.get('Item_Ancla','No')).strip() == 'Sí'
        row_vals = ([row['Mes'], row['Asignatura'], row['Grado'],
                     row.get('Item_Ancla','No')]
                    + [row.get(c, 0) for c in cats]
                    + [row['Total']])
        for ci, val in enumerate(row_vals, 1):
            cell = ws_r.cell(row=er, column=ci, value=val)
            cell.border    = THIN
            cell.alignment = CENTER
            cell.font      = Font(size=10)
        for cat_i, cat in enumerate(cats):
            ci  = 5 + cat_i
            val = row.get(cat, 0)
            if val > 0:
                aplicar_cat(ws_r, er, ci, cat, is_ancla)

    ws_r.auto_filter.ref = f"A1:{get_column_letter(len(res_hdrs))}1"

    wb.save(ARCHIVO_SALIDA)
    print(f"\n✔  Excel guardado en:\n   {ARCHIVO_SALIDA}")
    print("   Hojas:")
    for ws in wb.worksheets:
        print(f"     · {ws.title}")


# =============================================================================
# 8b. PATCH — corregir "Indicador no disponible" en Excel ya generado
# =============================================================================
def patch_excel_anclas(ruta_excel=None):
    """
    Reads an existing Estado_Global_Indicadores_Progreso.xlsx, looks up the
    metadata file for Mes 3, and replaces every 'Indicador no disponible (CODE)'
    cell in ALL sheets with the real indicator text + its Clase.

    Run this standalone after getting the metadata file:
        python Estado_Global_Indicadores_Progreso.py --patch
    """
    from openpyxl import load_workbook

    ruta = ruta_excel or ARCHIVO_SALIDA
    if not os.path.exists(ruta):
        print(f"✘  Archivo no encontrado: {ruta}")
        return

    print(f"[Patch] Cargando metadata para Mes {MES_CON_ANCLAS_REMOVIDAS}...")
    by_item_code, by_obj_code = cargar_metadata(MES_CON_ANCLAS_REMOVIDAS)

    if not by_item_code:
        print(f"✘  Metadata no encontrada. Asegúrese de tener "
              f"Items_Progreso_Mes{MES_CON_ANCLAS_REMOVIDAS}_2026_procesado.xlsx "
              f"en {PATH_META}")
        return

    print(f"[Patch] Abriendo Excel: {os.path.basename(ruta)}")
    wb = load_workbook(ruta)

    pattern = re.compile(r'Indicador no disponible \(([A-Z0-9]+)\)')
    total_fixed = 0
    total_clase_fixed = 0

    for ws in wb.worksheets:
        print(f"  Hoja: {ws.title}")
        # Find header row to locate Indicador and Clase columns
        header = {cell.value: cell.column for cell in ws[1]}
        col_ind   = header.get('Indicador')
        col_clase = header.get('Clase')

        if not col_ind:
            print(f"    ⚠  Columna 'Indicador' no encontrada — saltando")
            continue

        fixed_this_sheet = 0
        for row in ws.iter_rows(min_row=2):
            ind_cell = row[col_ind - 1]
            if not ind_cell.value:
                continue
            cell_str = str(ind_cell.value)

            # Match BOTH old formats:
            #   "Indicador no disponible (LEC315)"
            #   "LEC315 — indicador pendiente de metadata"
            m = re.search(r'\b((?:LEC|MAT)\d+)\b', cell_str)
            if not m:
                continue
            # Only patch rows that actually have the fallback text
            if ('no disponible' not in cell_str.lower()
                    and 'pendiente' not in cell_str.lower()
                    and 'sin texto' not in cell_str.lower()):
                continue

            code = m.group(1)
            meta = by_item_code.get(code, {})
            new_text = meta.get('objetivo', '')

            # Fallback: look up by objective code via by_obj_code
            # (finds the text if the same X.Y appears in a regular item)
            if not new_text and by_obj_code:
                # Try to find any entry in by_item_code whose indicador text
                # has the same fingerprint as another entry mapped to this code
                for icode, imeta in by_item_code.items():
                    if icode == code:
                        new_text = imeta.get('objetivo', '')
                        break

            if new_text:
                ind_cell.value = new_text
                total_fixed += 1
                fixed_this_sheet += 1

                # Fix Clase column
                if col_clase:
                    clase_cell = row[col_clase - 1]
                    clase_val  = str(clase_cell.value) if clase_cell.value else ''
                    new_clase  = meta.get('clases', '')
                    if new_clase and clase_val in ('', 'None', 'nan'):
                        clase_cell.value = new_clase
                        total_clase_fixed += 1

        print(f"    ✔  {fixed_this_sheet} celdas corregidas")

    if total_fixed > 0:
        wb.save(ruta)
        print(f"\n✔  Patch completado: {total_fixed} indicadores y "
              f"{total_clase_fixed} clases actualizados en {os.path.basename(ruta)}")
    else:
        print("\n  ℹ  No se encontraron celdas con 'Indicador no disponible' — "
              "el archivo ya está correcto o la metadata no tiene esos códigos.")


# =============================================================================
# 9. CROSS-MERGE: fuse anchor rows with regular rows sharing the same fingerprint
# =============================================================================
def _cross_merge_anchors(df_all):
    """
    After all rows (regular + anchor) are collected, merges anchor items into
    their regular counterparts when they share the same objective fingerprint.

    This mirrors exactly what report_profesores.py does: ALL items — whether
    regular or anchor — pass through make_lec_fingerprint(), so LEC315 whose
    metadata says "6.3 Compara..." automatically groups with the regular item
    that also starts with "6.3".

    Rules:
    - Only applies to Lengua rows (MAT anchor items stay separate).
    - Only merges within the same (Mes_Num, Mes, Grado) group.
    - When a regular row and an anchor row share a fingerprint:
        * Pct_Acierto → simple average of all pcts in the group
        * Categoria   → re-classified from the new average
        * Clase       → union of all Clase values
        * Indicador   → prefer the non-fallback text (real indicator description)
        * Item_Ancla  → 'Fusionado (incluye ancla)' to flag the merge
    - Rows whose text is still a fallback ('no disponible', 'pendiente de metadata')
      AND have no regular sibling to merge into: kept as-is with cleaner label.
    - MAT rows and non-Lengua rows pass through unchanged.
    """
    if df_all.empty:
        return df_all

    # Separate Lengua from everything else
    mask_lec = df_all['Asignatura'] == 'Lengua'
    df_lec   = df_all[mask_lec].copy()
    df_other = df_all[~mask_lec].copy()

    if df_lec.empty:
        return df_all

    merged_rows = []
    fused_count = 0

    # Group by (Mes_Num, Mes, Grado) — same scope as parsear_html's per-grade loop
    for group_key, grp in df_lec.groupby(['Mes_Num', 'Mes', 'Grado']):
        mes_num, mes_nombre, grado = group_key

        # Build fingerprint for every row in this group
        fp_groups = {}   # fingerprint → list of row indices

        for idx, row in grp.iterrows():
            fp = make_lec_fingerprint(str(row['Indicador']))
            if fp not in fp_groups:
                fp_groups[fp] = []
            fp_groups[fp].append(idx)

        for fp, indices in fp_groups.items():
            sub = grp.loc[indices]

            if len(indices) == 1:
                # No merging needed — pass through as-is
                # Clean up fallback label for lone anchor items
                row = sub.iloc[0].copy()
                if ('pendiente de metadata' in str(row['Indicador'])
                        or 'no disponible' in str(row['Indicador']).lower()):
                    code_m = re.search(r'((?:LEC|MAT)\d+)', str(row['Indicador']))
                    if code_m:
                        row['Indicador'] = f"{code_m.group(1)} — sin texto en metadata"
                merged_rows.append(row.to_dict())
                continue

            # Multiple rows share this fingerprint → merge them
            fused_count += 1
            pcts    = sub['Pct_Acierto'].dropna().tolist()
            pct_avg = round(sum(pcts) / len(pcts), 1) if pcts else None

            # Best indicator text: prefer non-fallback
            texts = sub['Indicador'].tolist()
            real_texts = [t for t in texts
                          if 'pendiente de metadata' not in str(t)
                          and 'no disponible' not in str(t).lower()
                          and not re.match(r'^(LEC|MAT)\d+\b', str(t))]
            best_text = real_texts[0] if real_texts else texts[0]

            # Merge Clase values
            clases = set()
            for v in sub['Clase'].dropna():
                for part in re.split(r'\s*\|\s*', str(v)):
                    if part.strip() and part.strip() != 'nan':
                        clases.add(part.strip())
            clases_sorted = sorted(
                clases,
                key=lambda x: [int(t) if t.isdigit() else t
                               for t in re.split(r'(\d+)', x)]
            )

            has_anchor  = (sub['Item_Ancla'] == 'Sí').any()
            ancla_flag  = 'Fusionado (incluye ancla)' if has_anchor else 'No'

            merged_rows.append({
                'Mes_Num':     mes_num,
                'Mes':         mes_nombre,
                'Asignatura':  'Lengua',
                'Grado':       grado,
                'Clase':       ' | '.join(clases_sorted),
                'Indicador':   best_text,
                'Pct_Acierto': pct_avg,
                'Categoria':   clasificar(pct_avg),
                'Item_Ancla':  ancla_flag,
            })

    if fused_count:
        print(f"  [Cross-merge] {fused_count} grupo(s) LEC fusionados "
              f"(ancla + regular con mismo código objetivo)")

    df_lec_merged = pd.DataFrame(merged_rows)
    return pd.concat([df_lec_merged, df_other], ignore_index=True)


# =============================================================================
# 10. MAIN
# =============================================================================
def main():
    # --patch mode: fix an existing Excel without re-running the full pipeline
    if '--patch' in sys.argv:
        patch_excel_anclas()
        return

    print("=" * 65)
    print("  Estado_Global_Indicadores_Progreso.py")
    print("  Excel consolidado + HTMLs con ítems ancla")
    print("=" * 65)

    all_rows         = []
    html_por_mes     = {}   # mes_num → ruta_html (for HTML regeneration)
    ancla_rows_por_mes = {} # mes_num → list of anchor rows

    for mes_num, mes_nombre, subcarpeta, html_name in MESES:
        print(f"\n{'─'*55}")
        print(f"  Mes {mes_num} — {mes_nombre}")
        print(f"{'─'*55}")

        # Locate HTML
        ruta_html = os.path.join(DIR_HTML, html_name)
        if not os.path.exists(ruta_html):
            print(f"  ⚠  HTML no encontrado: {ruta_html}")
            print(f"  → Saltando este mes.")
            continue

        # Load metadata (Clase enrichment + anchor text lookup)
        by_item_code, by_obj_code = cargar_metadata(mes_num)

        # Parse HTML indicators, enriching Clase from metadata where badges are absent
        rows = parsear_html(ruta_html, mes_num, mes_nombre, by_obj_code=by_obj_code)
        all_rows.extend(rows)

        # Reconstruct anchor items for any month where they were removed from HTML
        if mes_num in MESES_CON_ANCLAS_REMOVIDAS:
            html_por_mes[mes_num] = ruta_html
            print(f"\n  [Reconstrucción de ítems ancla — Mes {mes_num}]")
            anclas = cargar_anclas(mes_num)
            df_csv = cargar_csvs(mes_num, subcarpeta)

            if anclas and not df_csv.empty:
                ancla_rows = reconstruir_anclas(
                    df_csv, by_item_code, anclas, mes_num, mes_nombre
                )
                ancla_rows_por_mes[mes_num] = ancla_rows
                all_rows.extend(ancla_rows)
                print(f"  ✔  {len(ancla_rows)} ítems ancla añadidos")
            else:
                print("  ⚠  No se pudo reconstruir los ítems ancla:")
                if not anclas:
                    print("       - Archivo de anclas no encontrado")
                    print(f"        Formatos aceptados:")
                    print(f"          items_anclaje_progreso_mes_{mes_num}.xlsx  (wide)")
                    print(f"          Listado_ítems_ancla_progreso_{mes_num}.xlsx  (long)")
                if df_csv.empty:
                    print("       - CSVs de resultados no encontrados en Interim_CSVs/Resultados/")
                if not by_item_code:
                    print(f"       ℹ  Metadata no encontrada — textos de indicadores quedarán "
                          f"como código. Asegúrese de tener "
                          f"Items_Progreso_Mes{mes_num}_2026_procesado.xlsx en {PATH_META}")
                print(f"  → Mes {mes_num} incluido sin ítems ancla.")

    if not all_rows:
        print("\n✘  No se encontraron datos. Verifica las rutas.")
        return

    df_all = pd.DataFrame(all_rows)
    if 'Item_Ancla' not in df_all.columns:
        df_all['Item_Ancla'] = 'No'
    df_all = df_all.drop(columns=['_n_items_merged'], errors='ignore')

    # Cross-merge anchor LEC rows with regular rows sharing the same obj fingerprint
    df_all = _cross_merge_anchors(df_all)

    print(f"\n{'─'*55}")
    print(f"  RESUMEN TOTAL: {len(df_all)} filas")
    print(df_all.groupby(['Mes','Asignatura','Item_Ancla'])
          .size().rename('N').to_string())

    # Generate Excel
    print(f"\n[Excel] Generando...")
    generar_excel(df_all)

    # Regenerate HTML for every month where anchors were reconstructed
    for mes_num, ruta_html in html_por_mes.items():
        ancla_rows = ancla_rows_por_mes.get(mes_num, [])
        if ancla_rows:
            mes_nombre = next(m[1] for m in MESES if m[0] == mes_num)
            subcarpeta = next(m[2] for m in MESES if m[0] == mes_num)
            html_out   = os.path.join(
                DIR_SALIDA,
                f"Reporte_Nacional_Grados_{mes_num:02d}_PROGRESO_{mes_nombre}_con_Anclas.html"
            )
            print(f"\n[HTML] Regenerando HTML Mes {mes_num} con ítems ancla...")
            regenerar_html_con_anclas(ruta_html, ancla_rows, html_out)
            print(f"  ✔  {os.path.basename(html_out)}")
        else:
            print(f"\n  ℹ  HTML Mes {mes_num} no regenerado (sin ítems ancla disponibles).")

    print("\n" + "=" * 65)
    print("  COMPLETADO")
    print("=" * 65)


if __name__ == "__main__":
    main()