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

# Mes 3 necesita reconstrucción con ítems ancla
MES_CON_ANCLAS_REMOVIDAS = 3

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
        print(f"     Columnas: {list(df.columns)}")

        col_item  = next((c for c in df.columns
                          if 'item' in c.lower() or 'código' in c.lower()
                          or 'cod' in c.lower()), None)
        col_ind   = next((c for c in df.columns if 'indicador' in c.lower()), None)
        col_clase = next((c for c in df.columns
                          if 'clase' in c.lower() or 'sugerida' in c.lower()), None)
        col_orden = next((c for c in df.columns if 'orden' in c.lower()), None)
        col_obj   = next((c for c in df.columns
                          if 'objetivo' in c.lower() or 'obj' in c.lower()
                          or 'codigo_obj' in c.lower()), None)

        print(f"     col_item={col_item}  col_ind={col_ind}  "
              f"col_clase={col_clase}  col_obj={col_obj}")

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
    """Returns dict { grade_str: set_of_normalized_codes }"""
    anclas = {}
    patterns = [
        f"items_anclaje*mes*{mes_num}*.xlsx",
        f"items_anclaje*{mes_num}*.xlsx",
        f"*anclaje*mes*{mes_num}*.xlsx",
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
        return anclas

    print(f"  -> Anclas: {os.path.basename(found)}")
    try:
        xl = pd.ExcelFile(found)
        for sheet in xl.sheet_names:
            df_s = xl.parse(sheet, dtype=str)
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
            m = pattern.match(str(ind_cell.value))
            if not m:
                continue

            code = m.group(1)
            meta = by_item_code.get(code, {})

            # Try by item code first; fallback to obj_code lookup
            new_text = meta.get('objetivo', '')
            if not new_text and by_obj_code:
                # Try fingerprint match: strip code prefix, look for X.Y in metadata
                # Anchor codes don't have obj text embedded, so try all keys
                # that could map to this code via their item column
                pass   # best effort: metadata file resolves this

            if new_text:
                ind_cell.value = new_text
                total_fixed += 1
                fixed_this_sheet += 1

                # Also fix Clase column if it's empty/NaN
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
# 9. MAIN
# =============================================================================
def main():
    # --patch mode: fix an existing Excel without re-running the full pipeline
    if '--patch' in sys.argv:
        patch_excel_anclas()
        return

    print("=" * 65)
    print("  Estado_Global_Indicadores_Progreso.py")
    print("  Excel consolidado + HTML Mes 3 con ítems ancla")
    print("=" * 65)

    all_rows    = []
    html_mayo   = None   # path to Mayo HTML for regeneration
    ancla_rows_mayo = []

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

        # Load metadata for this month (needed for Clase enrichment + anchor reconstruction)
        by_item_code, by_obj_code = cargar_metadata(mes_num)

        # Parse present indicators from HTML, enriching Clase from metadata when needed
        rows = parsear_html(ruta_html, mes_num, mes_nombre, by_obj_code=by_obj_code)
        all_rows.extend(rows)

        # For the month where anchor items were removed: reconstruct them
        if mes_num == MES_CON_ANCLAS_REMOVIDAS:
            html_mayo = ruta_html
            print(f"\n  [Reconstrucción de ítems ancla para Mes {mes_num}]")
            anclas  = cargar_anclas(mes_num)
            df_csv  = cargar_csvs(mes_num, subcarpeta)

            # Gate: only need CSVs + anchor list (metadata optional — pcts still
            # calculated without it; indicator texts become 'pendiente de metadata')
            if anclas and not df_csv.empty:
                ancla_rows = reconstruir_anclas(
                    df_csv, by_item_code, anclas, mes_num, mes_nombre
                )
                ancla_rows_mayo = ancla_rows
                all_rows.extend(ancla_rows)
                print(f"  ✔  {len(ancla_rows)} ítems ancla añadidos al dataset")
            else:
                print("  ⚠  No se pudo reconstruir los ítems ancla:")
                if not anclas:    print("       - Archivo de anclas no encontrado")
                if df_csv.empty:  print("       - CSVs de resultados no encontrados")
                if not by_item_code:
                    print("       ℹ  Metadata no encontrada: los textos de indicadores "
                          f"se mostrarán como 'CÓDIGO — pendiente de metadata'. "
                          f"Asegúrese de tener Items_Progreso_Mes{mes_num}_2026_procesado.xlsx "
                          f"en {PATH_META}")
                print("  → El Excel incluirá el Mes 3 sin ítems ancla.")

    if not all_rows:
        print("\n✘  No se encontraron datos. Verifica las rutas.")
        return

    df_all = pd.DataFrame(all_rows)
    if 'Item_Ancla' not in df_all.columns:
        df_all['Item_Ancla'] = 'No'
    # Drop internal helper column (not for export)
    df_all = df_all.drop(columns=['_n_items_merged'], errors='ignore')

    print(f"\n{'─'*55}")
    print(f"  RESUMEN TOTAL: {len(df_all)} filas")
    print(df_all.groupby(['Mes','Asignatura','Item_Ancla'])
          .size().rename('N').to_string())

    # Generate Excel
    print(f"\n[Excel] Generando...")
    generar_excel(df_all)

    # Regenerate Mayo HTML with anchor items
    if html_mayo and ancla_rows_mayo:
        print(f"\n[HTML] Regenerando HTML Mes 3 con ítems ancla...")
        ruta_html_nueva = os.path.join(
            DIR_SALIDA,
            "Reporte_Nacional_Grados_03_PROGRESO_Mayo_con_Anclas.html"
        )
        regenerar_html_con_anclas(html_mayo, ancla_rows_mayo, ruta_html_nueva)
        print(f"  ✔  {os.path.basename(ruta_html_nueva)}")
    elif html_mayo and not ancla_rows_mayo:
        print("\n  ℹ  HTML Mes 3 no regenerado (no se encontraron ítems ancla).")

    print("\n" + "=" * 65)
    print("  COMPLETADO")
    print("=" * 65)


if __name__ == "__main__":
    main()