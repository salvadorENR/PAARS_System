import os
import glob
import sys
import re
import pandas as pd
import numpy as np
import warnings
import json
from pandas.errors import PerformanceWarning

warnings.filterwarnings('ignore', category=PerformanceWarning)

# ==========================================
# 1. CONFIGURACIÓN COMÚN
# ==========================================
ROOT_DIR = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse"
YEAR_DIR = os.path.join(ROOT_DIR, "2026")
PATH_METADATA = os.path.join(ROOT_DIR, "00_Metadata")

# ==========================================
# 2. CLAVE DE RESPUESTAS OFICIAL
# ==========================================
# Fuente autoritativa: PRUEBA_DIAGNÓSTICA_LENGUAJE_FINALÍSIMO.docx
# Verificado contra: BluePrint-Fundamentos.xlsx y 1_Resultados_Fundamento_Respuestas_Crudas.xlsx
#
# FORMATO DE RESPUESTAS EN EL FORMULARIO:
#   - Items 1–45  : letra suelta (ej. 'a', 'e', 'rr')
#   - Items 46–47 : numero con paréntesis (ej. '2)', '1)') — las opciones son palabras de audio
#   - Item  48    : '1) s - o - l'  (texto completo de la opción)
#   - Items 49–50 : '3) pan', '2) l'
#   - Items 51–56 : texto completo de la opción (ej. '1) pa', '3) sol', '1) mufe', '2) palo',
#                   '3) imagen de un sol brillando', '1) camino')  — NOTA: item 55 usa texto de imagen
#   - Items 57–64 : solo número con paréntesis (ej. '2)', '1)') — opciones son imágenes/frases de audio
#   - Items 65–68 : texto completo de la opción

CLAVE_RESPUESTAS = {
    # ── BLOQUE 1: Reconocimiento de Letras (Nombre) ── ítems 1–27 ──
    # Respuestas del formulario: letra suelta (valor real de la opción correcta)
    'L_item_1':  'a',   # B) a       — Letra A
    'L_item_2':  'e',   # C) e       — Letra E
    'L_item_3':  'i',   # B) i       — Letra I
    'L_item_4':  'o',   # C) o       — Letra O
    'L_item_5':  'u',   # B) u       — Letra U
    'L_item_6':  'm',   # C) m       — Letra M
    'L_item_7':  'p',   # B) p       — Letra P
    'L_item_8':  's',   # C) s       — Letra S
    'L_item_9':  'l',   # C) l       — Letra L
    'L_item_10': 't',   # B) t       — Letra T
    'L_item_11': 'n',   # C) n       — Letra N
    'L_item_12': 'd',   # C) d       — Letra D
    'L_item_13': 'r',   # B) r       — Letra R
    'L_item_14': 'b',   # B) b       — Letra B
    'L_item_15': 'c',   # B) c       — Letra C
    'L_item_16': 'f',   # B) f       — Letra F
    'L_item_17': 'g',   # C) g       — Letra G
    'L_item_18': 'h',   # A) h       — Letra H
    'L_item_19': 'j',   # C) j       — Letra J
    'L_item_20': 'k',   # B) k       — Letra K
    'L_item_21': 'ñ',   # B) ñ       — Letra Ñ
    'L_item_22': 'q',   # C) q       — Letra Q
    'L_item_23': 'v',   # B) v       — Letra V
    'L_item_24': 'w',   # C) w       — Letra W
    'L_item_25': 'x',   # B) x       — Letra X
    'L_item_26': 'y',   # C) y       — Letra Y
    'L_item_27': 'z',   # C) z       — Letra Z

    # ── BLOQUE 2: Correspondencia Sonido–Letra ── ítems 28–45 ──
    'L_item_28': 'a',   # C) a       — Fonema /a/
    'L_item_29': 'e',   # C) e       — Fonema /e/
    'L_item_30': 'i',   # B) i       — Fonema /i/
    'L_item_31': 'o',   # C) o       — Fonema /o/
    'L_item_32': 'u',   # A) u       — Fonema /u/
    'L_item_33': 'm',   # C) m       — Fonema /m/
    'L_item_34': 'p',   # B) p       — Fonema /p/
    'L_item_35': 's',   # A) s       — Fonema /s/
    'L_item_36': 'l',   # B) l       — Fonema /l/
    'L_item_37': 't',   # A) t       — Fonema /t/
    'L_item_38': 'n',   # C) n       — Fonema /n/
    'L_item_39': 'd',   # B) d       — Fonema /d/
    'L_item_40': 'rr',  # B) rr      — Fonema /rr/ vibrante múltiple
    'L_item_41': 'r',   # A) r       — Fonema /r/ vibrante simple
    'L_item_42': 'g',   # B) g       — Fonema /g/
    'L_item_43': 'f',   # B) f       — Fonema /f/
    'L_item_44': 'b',   # C) b       — Fonema /b/
    'L_item_45': 'k',   # B) k       — Fonema /k/

    # ── BLOQUE 3: Conciencia Fonológica ── ítems 46–50 ──
    # El formulario guarda texto completo de la opción seleccionada
    'L_item_46': '2)',              # B) sopa       — sonido inicial
    'L_item_47': '1)',              # A) taza       — rima  ← CORREGIDO (era '1)', ahora '1)' con opción A)
    'L_item_48': '1) s - o - l',   # A) /s/–/o/–/l/ — segmentación  ← CORREGIDO (era '1) s - o - l')
    'L_item_49': '3) pan',         # C) pan        — fusión (blending)
    'L_item_50': '2) l',           # B) l          — sonido final

    # ── BLOQUE 4: Lectura de Sílabas y Palabras sin Sentido ── ítems 51–53 ──
    'L_item_51': '1) pa',          # A) pa         — sílaba directa CV
    'L_item_52': '3) sol',         # C) sol        — palabra monosílaba CVC
    'L_item_53': '1) mufe',        # A) mufe       — pseudopalabra CV-CV

    # ── BLOQUE 5: Lectura de Palabras Familiares ── ítems 54–56 ──
    'L_item_54': '2) palo',                    # B) palo       — reconocimiento de palabra bisílaba
    'L_item_55': '3)',  # C) imagen sol — frase corta  ← CORREGIDO (texto real del formulario)
    'L_item_56': '1) camino',                  # A) camino     — palabra trisílaba

    # ── BLOQUE 6: Comprensión Auditiva ── ítems 57–64 ──
    # El formulario guarda solo número con paréntesis ('1)', '2)', '3)')
    'L_item_57': '2)',   # B) un perro       — literal personaje  ← CORREGIDO (código tenía '1)')
    'L_item_58': '1)',   # A) salta          — literal acción
    'L_item_59': '1)',   # A) porque va a llover — inferencial causa
    'L_item_60': '2)',   # B) un niño leyendo   — inferencial experiencia
    'L_item_61': '3)',   # C) camina al parque  — secuencia narración
    'L_item_62': '2)',   # B) el perro          — info explícita narración
    'L_item_63': '1)',   # A) triste            — inferencia emocional
    'L_item_64': '3)',   # C) porque empieza a llover — causa en narración

    # ── BLOQUE 7: Comprensión Lectora ── ítems 65–68 ──
    'L_item_65': '2) Nube',                    # B) Nube              — literal
    'L_item_66': '2) Es un perro',             # B) Es un perro       — reorganización
    'L_item_67': '1) Está lloviendo',          # A) Está lloviendo    — inferencial
    'L_item_68': '3) Actuó de forma honesta',  # C) Actuó de forma honesta — crítico

    # ── MATEMÁTICA ── ítems 1–18 ── (sin cambios, ya estaba correcto)
    'M_item_1':  '3) 9',
    'M_item_2':  '2) 15',
    'M_item_3':  '2) 18 es mayor que 15',
    'M_item_4':  '2) 25',
    'M_item_5':  '1) 4 decenas y 8 unidades',
    'M_item_6':  '1) 34',
    'M_item_7':  '1) 67 < 76',
    'M_item_8':  '2) 25, 40 y 52',
    'M_item_9':  '2) Decenas',
    'M_item_10': '3) 205',
    'M_item_11': '2) 7',
    'M_item_12': '1) 4',
    'M_item_13': '3) 57',
    'M_item_14': '2) 33',
    'M_item_15': '3) 23',
    'M_item_16': '2) 19',
    'M_item_17': '3) 16',
    'M_item_18': '2) 12',
}

# ==========================================
# 3. MAPEOS PEDAGÓGICOS
# ==========================================
# Bloques alineados con BluePrint-Fundamentos.xlsx y el documento de la prueba.
# CORRECCIONES respecto a la versión anterior:
#   - Bloque 1 renombrado (era Bloque0)
#   - Bloque 4 y Bloque 5 separados (antes se unían en Bloque2b)
#   - Bloque 6 unificado con todos los 8 ítems 57–64 (antes dividido en Bloque3b + Bloque4)
#   - Bloque 7 renombrado (era Bloque5)

L_SECCIONES = {
    'L_Bloque1_Reconocimiento_de_Letras':              [f'L_item_{i}' for i in range(1,  28)],  # 27 ítems
    'L_Bloque2_Correspondencia_Sonido_Letra':          [f'L_item_{i}' for i in range(28, 46)],  # 18 ítems
    'L_Bloque3_Conciencia_Fonologica':                 [f'L_item_{i}' for i in range(46, 51)],  #  5 ítems
    'L_Bloque4_Lectura_Silabas_Palabras_sin_Sentido':  [f'L_item_{i}' for i in range(51, 54)],  #  3 ítems
    'L_Bloque5_Lectura_Palabras_Familiares':           [f'L_item_{i}' for i in range(54, 57)],  #  3 ítems
    'L_Bloque6_Comprension_Auditiva':                  [f'L_item_{i}' for i in range(57, 65)],  #  8 ítems
    'L_Bloque7_Comprension_Lectora':                   [f'L_item_{i}' for i in range(65, 69)],  #  4 ítems
}

M_SECCIONES = {
    'M_Bloque1_Conteo_y_Sentido_Numerico':             [f'M_item_{i}' for i in range(1,  5)],   #  4 ítems
    'M_Bloque2_Valor_Posicional_y_Sistema_Decimal':    [f'M_item_{i}' for i in range(5,  11)],  #  6 ítems
    'M_Bloque3_Operaciones_Aditivas_y_Sustractivas':   [f'M_item_{i}' for i in range(11, 19)],  #  8 ítems
}

TODAS_SECCIONES = {**L_SECCIONES, **M_SECCIONES}

# Proceso por ítem — alinhado con BluePrint columna Tarea_Descripción
L_PROCESS_MAP = {
    # Bloque 1
    'L_item_1':  'Reconocer el nombre del grafema: Letra A',
    'L_item_2':  'Reconocer el nombre del grafema: Letra E',
    'L_item_3':  'Reconocer el nombre del grafema: Letra I',
    'L_item_4':  'Reconocer el nombre del grafema: Letra O',
    'L_item_5':  'Reconocer el nombre del grafema: Letra U',
    'L_item_6':  'Reconocer el nombre del grafema: Letra M',
    'L_item_7':  'Reconocer el nombre del grafema: Letra P',
    'L_item_8':  'Reconocer el nombre del grafema: Letra S',
    'L_item_9':  'Reconocer el nombre del grafema: Letra L',
    'L_item_10': 'Reconocer el nombre del grafema: Letra T',
    'L_item_11': 'Reconocer el nombre del grafema: Letra N',
    'L_item_12': 'Reconocer el nombre del grafema: Letra D',
    'L_item_13': 'Reconocer el nombre del grafema: Letra R',
    'L_item_14': 'Reconocer nombre ante distractor visual: Letra B (b/d/p)',
    'L_item_15': 'Reconocer nombre ante distractor visual: Letra C (c/o/s)',
    'L_item_16': 'Reconocer nombre ante distractor visual: Letra F (f/t/e)',
    'L_item_17': 'Reconocer nombre ante distractor visual: Letra G (g/j/p)',
    'L_item_18': 'Reconocer nombre ante distractor visual: Letra H (h/n/l)',
    'L_item_19': 'Reconocer nombre ante distractor visual: Letra J (j/i/|)',
    'L_item_20': 'Reconocer nombre ante distractor visual: Letra K (k/x/h)',
    'L_item_21': 'Reconocer nombre ante distractor visual: Letra Ñ (ñ/n/m)',
    'L_item_22': 'Reconocer nombre ante distractor visual: Letra Q (q/p/o)',
    'L_item_23': 'Reconocer nombre ante distractor visual: Letra V (v/u/y)',
    'L_item_24': 'Reconocer nombre ante distractor visual: Letra W (w/m/v)',
    'L_item_25': 'Reconocer nombre ante distractor visual: Letra X (x/k/y)',
    'L_item_26': 'Reconocer nombre ante distractor visual: Letra Y (y/v/i)',
    'L_item_27': 'Reconocer nombre ante distractor visual: Letra Z (z/s/c)',
    # Bloque 2
    'L_item_28': 'Asociar fonema a grafema: /a/',
    'L_item_29': 'Asociar fonema a grafema: /e/',
    'L_item_30': 'Asociar fonema a grafema: /i/',
    'L_item_31': 'Asociar fonema a grafema: /o/',
    'L_item_32': 'Asociar fonema a grafema: /u/',
    'L_item_33': 'Asociar fonema a grafema: /m/',
    'L_item_34': 'Asociar fonema a grafema: /p/',
    'L_item_35': 'Asociar fonema a grafema: /s/',
    'L_item_36': 'Asociar fonema a grafema: /l/',
    'L_item_37': 'Asociar fonema a grafema: /t/',
    'L_item_38': 'Asociar fonema a grafema: /n/',
    'L_item_39': 'Asociar fonema a grafema: /d/',
    'L_item_40': 'Asociar fonema a grafema: /rr/ (vibrante múltiple)',
    'L_item_41': 'Asociar fonema a grafema: /r/ (vibrante simple)',
    'L_item_42': 'Asociar fonema a grafema: /g/ (sonido suave)',
    'L_item_43': 'Asociar fonema a grafema: /f/',
    'L_item_44': 'Asociar fonema a grafema: /b/',
    'L_item_45': 'Asociar fonema a grafema: /k/',
    # Bloque 3
    'L_item_46': 'Identificar el sonido inicial compartido entre palabras (onset)',
    'L_item_47': 'Reconocer palabras que riman (conciencia de rima)',
    'L_item_48': 'Segmentar una palabra en sus fonemas individuales',
    'L_item_49': 'Fusionar fonemas aislados para formar una palabra (blending)',
    'L_item_50': 'Identificar el grafema que representa el sonido final de una palabra',
    # Bloque 4
    'L_item_51': 'Decodificar sílaba directa de estructura CV',
    'L_item_52': 'Decodificar palabra monosílaba de estructura CVC',
    'L_item_53': 'Decodificar pseudopalabra CV-CV sin apoyo léxico',
    # Bloque 5
    'L_item_54': 'Reconocer palabra bisílaba frecuente por discriminación auditiva',
    'L_item_55': 'Identificar el referente de una frase corta de dos palabras',
    'L_item_56': 'Reconocer palabra trisílaba frecuente ante distractores',
    # Bloque 6
    'L_item_57': 'Identificar el referente (personaje) mencionado en una frase oral simple',
    'L_item_58': 'Identificar la acción descrita en una frase oral simple',
    'L_item_59': 'Inferir la causa implícita de una situación descrita en una frase oral',
    'L_item_60': 'Relacionar una acción descrita oralmente con una experiencia conocida',
    'L_item_61': 'Identificar la primera acción en una secuencia narrada oralmente',
    'L_item_62': 'Identificar el personaje responsable de un hecho (pista implícita)',
    'L_item_63': 'Inferir el estado emocional de un personaje a partir de indicios físicos',
    'L_item_64': 'Recuperar la causa explícita de una acción narrada en un texto oral breve',
    # Bloque 7
    'L_item_65': 'Recuperar información explícita de un texto escrito: nombre del personaje',
    'L_item_66': 'Clasificar y relacionar información explícita distribuida en el texto',
    'L_item_67': 'Deducir información implícita a partir de indicios del texto escrito',
    'L_item_68': 'Emitir un juicio valorativo sobre la conducta de un personaje',
}

M_PROCESS_MAP = {
    'M_item_1':  'Conteo visual',
    'M_item_2':  'Secuencia numérica',
    'M_item_3':  'Comparación numérica',
    'M_item_4':  'Patrón numérico',
    'M_item_5':  'Descomposición DU',
    'M_item_6':  'Convertir DU a número',
    'M_item_7':  'Comparación con signos',
    'M_item_8':  'Ordenar números',
    'M_item_9':  'Valor posicional',
    'M_item_10': 'Construcción C-D-U',
    'M_item_11': 'Suma básica',
    'M_item_12': 'Resta básica',
    'M_item_13': 'Suma de dos cifras',
    'M_item_14': 'Resta sin préstamo',
    'M_item_15': 'Resta con préstamo',
    'M_item_16': 'Problema aditivo',
    'M_item_17': 'Problema sustractivo',
    'M_item_18': 'Multiplicación inicial',
}

TODO_PROCESS_MAP = {**L_PROCESS_MAP, **M_PROCESS_MAP}

# ==========================================
# 4. JERARQUÍAS PARA EL HTML
# ==========================================
# Alineadas con BluePrint: grupos (Grupo_Cod / Grupo_Descripción)

LENGUA_JERARQUIA = [
    {
        'bloque_col': 'L_Bloque1_Reconocimiento_de_Letras',
        'titulo':     'Bloque 1 — Reconocimiento de Letras (Nombre) · B01LEC01',
        'grupos': [
            {'label': 'Grupo 1: Vocales — G01LEC01',
             'items': [f'L_item_{i}' for i in range(1, 6)]},
            {'label': 'Grupo 2: Consonantes de Alta Frecuencia — G01LEC02',
             'items': [f'L_item_{i}' for i in range(6, 14)]},
            {'label': 'Grupo 3: Consonantes de Confusión Visual y Baja Frecuencia — G01LEC03',
             'items': [f'L_item_{i}' for i in range(14, 28)]},
        ]
    },
    {
        'bloque_col': 'L_Bloque2_Correspondencia_Sonido_Letra',
        'titulo':     'Bloque 2 — Correspondencia Sonido–Letra · B01LEC02',
        'grupos': [
            {'label': 'Grupo 1: Vocales — G01LEC04',
             'items': [f'L_item_{i}' for i in range(28, 33)]},
            {'label': 'Grupo 2: Consonantes de Alta Frecuencia — G01LEC05',
             'items': [f'L_item_{i}' for i in range(33, 40)]},
            {'label': 'Grupo 3: Consonantes de Menor Frecuencia o Mayor Dificultad — G01LEC06',
             'items': [f'L_item_{i}' for i in range(40, 46)]},
        ]
    },
    {
        'bloque_col': 'L_Bloque3_Conciencia_Fonologica',
        'titulo':     'Bloque 3 — Conciencia Fonológica · B01LEC03',
        'grupos': [
            {'label': 'Conciencia Fonológica (audio-only) — G01LEC07',
             'items': [f'L_item_{i}' for i in range(46, 51)]},
        ]
    },
    {
        'bloque_col': 'L_Bloque4_Lectura_Silabas_Palabras_sin_Sentido',
        'titulo':     'Bloque 4 — Lectura de Sílabas y Palabras sin Sentido · B01LEC04',
        'grupos': [
            {'label': 'Decodificación silábica y léxica — G01LEC08',
             'items': [f'L_item_{i}' for i in range(51, 54)]},
        ]
    },
    {
        'bloque_col': 'L_Bloque5_Lectura_Palabras_Familiares',
        'titulo':     'Bloque 5 — Lectura de Palabras Familiares · B01LEC05',
        'grupos': [
            {'label': 'Reconocimiento automático de palabras — G01LEC09',
             'items': [f'L_item_{i}' for i in range(54, 57)]},
        ]
    },
    {
        'bloque_col': 'L_Bloque6_Comprension_Auditiva',
        'titulo':     'Bloque 6 — Comprensión Auditiva · B01LEC06',
        'grupos': [
            {'label': 'Grupo 1: Comprensión de frases simples — nivel literal — G01LEC10',
             'items': ['L_item_57', 'L_item_58']},
            {'label': 'Grupo 2: Comprensión de frases simples — nivel inferencial — G01LEC11',
             'items': ['L_item_59', 'L_item_60']},
            {'label': 'Grupo 3: Comprensión de narraciones cortas — literal e inferencial — G01LEC12',
             'items': ['L_item_61', 'L_item_62', 'L_item_63', 'L_item_64']},
        ]
    },
    {
        'bloque_col': 'L_Bloque7_Comprension_Lectora',
        'titulo':     'Bloque 7 — Comprensión Lectora · B01LEC07',
        'grupos': [
            {'label': 'Grupo 1: Nivel literal — G01LEC13',
             'items': ['L_item_65']},
            {'label': 'Grupo 2: Nivel de reorganización — G01LEC14',
             'items': ['L_item_66']},
            {'label': 'Grupo 3: Nivel inferencial — G01LEC15',
             'items': ['L_item_67']},
            {'label': 'Grupo 4: Nivel crítico-valorativo — G01LEC16',
             'items': ['L_item_68']},
        ]
    },
]

MATEMATICA_JERARQUIA = [
    {
        'bloque_col': 'M_Bloque1_Conteo_y_Sentido_Numerico',
        'titulo':     'Bloque 1 — Conteo y Sentido Numérico · B00MAT01',
        'grupos': [
            {'label': 'Ítems 1–4 — G00MAT01',
             'items': [f'M_item_{i}' for i in range(1, 5)]},
        ]
    },
    {
        'bloque_col': 'M_Bloque2_Valor_Posicional_y_Sistema_Decimal',
        'titulo':     'Bloque 2 — Valor Posicional y Sistema Decimal · B00MAT02',
        'grupos': [
            {'label': 'Decenas y Unidades (D-U) — G00MAT02',
             'items': [f'M_item_{i}' for i in range(5, 8)]},
            {'label': 'Centenas, Decenas y Unidades (C-D-U) — G00MAT03',
             'items': [f'M_item_{i}' for i in range(8, 11)]},
        ]
    },
    {
        'bloque_col': 'M_Bloque3_Operaciones_Aditivas_y_Sustractivas',
        'titulo':     'Bloque 3 — Operaciones Aditivas y Sustractivas · B00MAT03',
        'grupos': [
            {'label': 'Operaciones básicas (una cifra) — G00MAT04',
             'items': ['M_item_11', 'M_item_12']},
            {'label': 'Operaciones de dos cifras — G00MAT05',
             'items': ['M_item_13', 'M_item_14', 'M_item_15']},
            {'label': 'Problemas verbales — G00MAT06',
             'items': ['M_item_16', 'M_item_17', 'M_item_18']},
        ]
    },
]

# Etiquetas de ítems para el explorador dinámico
ITEM_LABELS = {
    # Bloque 1
    'L_item_1':  'Letra A',  'L_item_2':  'Letra E',  'L_item_3':  'Letra I',
    'L_item_4':  'Letra O',  'L_item_5':  'Letra U',
    'L_item_6':  'Letra M',  'L_item_7':  'Letra P',  'L_item_8':  'Letra S',
    'L_item_9':  'Letra L',  'L_item_10': 'Letra T',  'L_item_11': 'Letra N',
    'L_item_12': 'Letra D',  'L_item_13': 'Letra R',
    'L_item_14': 'Letra B',  'L_item_15': 'Letra C',  'L_item_16': 'Letra F',
    'L_item_17': 'Letra G',  'L_item_18': 'Letra H',  'L_item_19': 'Letra J',
    'L_item_20': 'Letra K',  'L_item_21': 'Letra Ñ',  'L_item_22': 'Letra Q',
    'L_item_23': 'Letra V',  'L_item_24': 'Letra W',  'L_item_25': 'Letra X',
    'L_item_26': 'Letra Y',  'L_item_27': 'Letra Z',
    # Bloque 2
    'L_item_28': 'Fonema /a/', 'L_item_29': 'Fonema /e/', 'L_item_30': 'Fonema /i/',
    'L_item_31': 'Fonema /o/', 'L_item_32': 'Fonema /u/',
    'L_item_33': 'Fonema /m/', 'L_item_34': 'Fonema /p/', 'L_item_35': 'Fonema /s/',
    'L_item_36': 'Fonema /l/', 'L_item_37': 'Fonema /t/', 'L_item_38': 'Fonema /n/',
    'L_item_39': 'Fonema /d/', 'L_item_40': 'Fonema /rr/', 'L_item_41': 'Fonema /r/',
    'L_item_42': 'Fonema /g/', 'L_item_43': 'Fonema /f/', 'L_item_44': 'Fonema /b/',
    'L_item_45': 'Fonema /k/',
    # Bloque 3
    'L_item_46': 'Ítem 46 — Sonido inicial',
    'L_item_47': 'Ítem 47 — Rima',
    'L_item_48': 'Ítem 48 — Segmentación fonémica',
    'L_item_49': 'Ítem 49 — Fusión (blending)',
    'L_item_50': 'Ítem 50 — Sonido final',
    # Bloque 4
    'L_item_51': 'Ítem 51 — Sílaba CV',
    'L_item_52': 'Ítem 52 — Palabra CVC',
    'L_item_53': 'Ítem 53 — Pseudopalabra CV-CV',
    # Bloque 5
    'L_item_54': 'Ítem 54 — Reconoc. bisílaba (palo)',
    'L_item_55': 'Ítem 55 — Frase corta (El sol brilla)',
    'L_item_56': 'Ítem 56 — Reconoc. trisílaba (camino)',
    # Bloque 6
    'L_item_57': 'Ítem 57 — Literal: personaje',
    'L_item_58': 'Ítem 58 — Literal: acción',
    'L_item_59': 'Ítem 59 — Inferencial: causa',
    'L_item_60': 'Ítem 60 — Inferencial: experiencia',
    'L_item_61': 'Ítem 61 — Secuencia: primera acción',
    'L_item_62': 'Ítem 62 — Narración: info explícita',
    'L_item_63': 'Ítem 63 — Narración: inferencia emocional',
    'L_item_64': 'Ítem 64 — Narración: causa explícita',
    # Bloque 7
    'L_item_65': 'Ítem 65 — Comp. Lectora: nivel literal',
    'L_item_66': 'Ítem 66 — Comp. Lectora: reorganización',
    'L_item_67': 'Ítem 67 — Comp. Lectora: inferencial',
    'L_item_68': 'Ítem 68 — Comp. Lectora: crítico-valorativo',
    # Matemática
    'M_item_1':  'Ítem 1 — Conteo visual',
    'M_item_2':  'Ítem 2 — Secuencia numérica',
    'M_item_3':  'Ítem 3 — Comparación numérica',
    'M_item_4':  'Ítem 4 — Patrón numérico',
    'M_item_5':  'Ítem 5 — Descomposición DU',
    'M_item_6':  'Ítem 6 — Convertir DU a número',
    'M_item_7':  'Ítem 7 — Comparación con signos',
    'M_item_8':  'Ítem 8 — Ordenar números',
    'M_item_9':  'Ítem 9 — Valor posicional',
    'M_item_10': 'Ítem 10 — Construcción C-D-U',
    'M_item_11': 'Ítem 11 — Suma básica',
    'M_item_12': 'Ítem 12 — Resta básica',
    'M_item_13': 'Ítem 13 — Suma 2 cifras',
    'M_item_14': 'Ítem 14 — Resta sin préstamo',
    'M_item_15': 'Ítem 15 — Resta con préstamo',
    'M_item_16': 'Ítem 16 — Problema aditivo',
    'M_item_17': 'Ítem 17 — Problema sustractivo',
    'M_item_18': 'Ítem 18 — Multiplicación inicial',
}

# Nombres cortos para renombrar columnas en los reportes Excel
RENAME_MAP_BLOQUES = {
    'L_Bloque1_Reconocimiento_de_Letras_puntaje_0_100':             'L-B1: Reconoc. Letras',
    'L_Bloque2_Correspondencia_Sonido_Letra_puntaje_0_100':         'L-B2: Corr. Sonido-Letra',
    'L_Bloque3_Conciencia_Fonologica_puntaje_0_100':                'L-B3: Conc. Fonológica',
    'L_Bloque4_Lectura_Silabas_Palabras_sin_Sentido_puntaje_0_100': 'L-B4: Lectura Sílabas',
    'L_Bloque5_Lectura_Palabras_Familiares_puntaje_0_100':          'L-B5: Palabs. Familiares',
    'L_Bloque6_Comprension_Auditiva_puntaje_0_100':                 'L-B6: Comp. Auditiva',
    'L_Bloque7_Comprension_Lectora_puntaje_0_100':                  'L-B7: Comp. Lectora',
    'M_Bloque1_Conteo_y_Sentido_Numerico_puntaje_0_100':            'M-B1: Conteo',
    'M_Bloque2_Valor_Posicional_y_Sistema_Decimal_puntaje_0_100':   'M-B2: Valor Posicional',
    'M_Bloque3_Operaciones_Aditivas_y_Sustractivas_puntaje_0_100':  'M-B3: Operaciones',
    'L_total_puntaje_0_100':  'L: Total',
    'M_total_puntaje_0_100':  'M: Total',
    'Nombre del centro':      'Centro Educativo',
}

# ==========================================
# 5. FUNCIÓN DE PUNTUACIÓN PORCENTUAL
# ==========================================
def calcular_puntaje_porcentual(suma_correcta_serie, total_items):
    if total_items == 0:
        return pd.Series([0.0] * len(suma_correcta_serie), index=suma_correcta_serie.index)
    return (suma_correcta_serie / total_items * 100).clip(0, 100)


# ==========================================
# 6. GENERADOR DE EXCEL CON DISTRACTORES
# ==========================================
def exportar_distractores(dict_dfs, ruta_salida, clave_respuestas):
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    yellow_fill  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_fill  = PatternFill(start_color="033b6d", end_color="033b6d", fill_type="solid")
    white_font   = Font(color="FFFFFF", bold=True)
    thin_border  = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )

    for sheet_name, df in dict_dfs.items():
        ws = wb.create_sheet(title=sheet_name)
        for col_num, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_num, value=col_name)
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(col_num)].width = (
                15 if col_name in ['Asignatura', 'Ítem'] else 32
            )

        for row_num, row_data in enumerate(df.to_dict('records'), 2):
            item_name = row_data['Ítem']
            correcta  = clave_respuestas.get(item_name, "").lower().strip()
            for col_num, col_name in enumerate(df.columns, 1):
                val  = row_data.get(col_name)
                cell = ws.cell(row=row_num, column=col_num, value=val)
                cell.border    = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_name.startswith('Opción') and pd.notna(val) and isinstance(val, str):
                    if ' (' in val:
                        opt_text = val.rsplit(' (', 1)[0].strip().lower()
                        if opt_text == correcta and correcta != "":
                            cell.fill = yellow_fill

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}1"

    wb.save(ruta_salida)


# ==========================================
# 7. GENERADOR DE HTML INTERACTIVO
# ==========================================
def generar_html_estatico(df_scored, df_mapping, ruta_salida):
    try:
        import plotly.express as px
        import plotly.graph_objects as go
    except ImportError:
        print("    [!] Para generar el HTML necesitas Plotly (pip install plotly).")
        return

    C_DEEP_NAVY  = '#033b6d'
    C_ROYAL_BLUE = '#094b93'
    C_OCEAN_BLUE = '#3077b9'
    C_SKY_BLUE   = '#a7d2f2'
    C_STEEL_GREY = '#666766'

    def _lm(labels, px_per_char=8, minimum=180):
        return max(max((len(str(l)) for l in labels), default=0) * px_per_char, minimum)

    def _bloque_box(bloque_col, btitle, color, include_plotlyjs=False):
        col  = f'{bloque_col}_puntaje_0_100'
        if col not in df_scored.columns: return ''
        data = pd.to_numeric(df_scored[col], errors='coerce').dropna()
        fig  = go.Figure()
        fig.add_trace(go.Box(
            y=data, boxpoints='all', jitter=0.3, pointpos=-1.8,
            fillcolor=f'rgba({int(color[1:3],16)},{int(color[3:5],16)},{int(color[5:7],16)},0.3)',
            line_color=color, marker=dict(color=color, size=5, opacity=0.6), name='Distribución',
        ))
        fig.update_layout(
            title=f'Distribución del puntaje — {btitle}',
            yaxis=dict(title='Puntaje (0–100)', range=[-5, 105],
                       gridcolor=C_SKY_BLUE, tickfont=dict(color=C_DEEP_NAVY)),
            xaxis=dict(showticklabels=False), plot_bgcolor='white', paper_bgcolor='white',
            font_color=C_DEEP_NAVY, height=320, margin=dict(l=60, r=60, t=55, b=20),
            showlegend=False,
        )
        fig.add_hline(y=data.mean(), line_dash='dot', line_color=C_DEEP_NAVY, opacity=0.6,
                      annotation_text=f'Media: {data.mean():.1f}',
                      annotation_position='top right')
        return fig.to_html(full_html=False, include_plotlyjs=False)

    def _bar_items(items, labels, vals, title, color):
        num_items = len(items)
        height    = max(400, num_items * 38 + 120)
        fig = px.bar(x=vals, y=labels, orientation='h', title=title,
                     text=vals, color_discrete_sequence=[color])
        fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside',
                          marker_color=color)
        fig.update_layout(
            xaxis_range=[0, 118], height=height,
            margin=dict(l=_lm(labels), r=80, t=50, b=30),
            plot_bgcolor='white', paper_bgcolor='white', font_color=C_DEEP_NAVY,
            yaxis=dict(tickfont=dict(size=11, color=C_DEEP_NAVY),
                       autorange='reversed', fixedrange=True,
                       tickmode='linear', dtick=1),
            xaxis=dict(title='% de Aciertos', gridcolor=C_SKY_BLUE, ticksuffix='%'),
        )
        return fig.to_html(full_html=False, include_plotlyjs=False)

    ORDEN_GRADOS = ['Primer Grado','Segundo Grado','Tercer Grado','Cuarto Grado',
                    'Quinto Grado','Sexto Grado','Séptimo Grado','Octavo Grado',
                    'Noveno Grado','1er Año','2do Año']

    df_validos = df_scored[df_scored['Nombre del centro'] != 'SIN REGISTRO'].copy()

    df_pivot_l = df_validos.pivot_table(
        index='Nombre del centro', columns='Grado',
        values='L_total_puntaje_0_100', aggfunc='mean')
    cols_l = [g for g in ORDEN_GRADOS if g in df_pivot_l.columns]
    df_pivot_l = df_pivot_l[cols_l].reset_index().round(1).fillna('-')
    df_pivot_l.rename(columns={'Nombre del centro': 'Centro Educativo'}, inplace=True)

    df_pivot_m = df_validos.pivot_table(
        index='Nombre del centro', columns='Grado',
        values='M_total_puntaje_0_100', aggfunc='mean')
    cols_m = [g for g in ORDEN_GRADOS if g in df_pivot_m.columns]
    df_pivot_m = df_pivot_m[cols_m].reset_index().round(1).fillna('-')
    df_pivot_m.rename(columns={'Nombre del centro': 'Centro Educativo'}, inplace=True)

    all_l_items = [f'L_item_{i}' for i in range(1, 69) if f'L_item_{i}' in df_scored.columns]
    all_m_items = [f'M_item_{i}' for i in range(1, 19) if f'M_item_{i}' in df_scored.columns]
    all_items   = all_l_items + all_m_items

    df_item_means   = df_validos.groupby(['Nombre del centro','Grado'])[all_items].mean() * 100
    df_item_means   = df_item_means.round(1)
    df_student_cnt  = df_validos.groupby(['Nombre del centro','Grado']).size()

    dynamic_data = {}
    dynamic_data["-- Todas las escuelas --"] = {}
    df_global_means = df_validos.groupby('Grado')[all_items].mean() * 100
    df_global_cnt   = df_validos.groupby('Grado').size()
    for grade, row in df_global_means.iterrows():
        d = row.round(1).to_dict()
        d['N'] = int(df_global_cnt.loc[grade])
        dynamic_data["-- Todas las escuelas --"][grade] = d

    for (school, grade), row in df_item_means.iterrows():
        if school not in dynamic_data:
            dynamic_data[school] = {}
        try:
            n = int(df_student_cnt.loc[school, grade])
        except KeyError:
            n = 0
        d = row.to_dict()
        d['N'] = n
        dynamic_data[school][grade] = d

    json_data    = json.dumps(dynamic_data)
    json_labels  = json.dumps(ITEM_LABELS)
    json_items_l = json.dumps(all_l_items)
    json_items_m = json.dumps(all_m_items)

    df_tabla = df_scored.copy()
    cols_show = (['NIE','Nombre del centro','Grado','Grupo',
                  'L_total_puntaje_0_100','M_total_puntaje_0_100']
                 + [c for c in df_tabla.columns if 'puntaje_0_100' in c and 'total' not in c])
    df_tabla = df_tabla[[c for c in cols_show if c in df_tabla.columns]]
    for c in df_tabla.select_dtypes(include=['float64']).columns:
        df_tabla[c] = df_tabla[c].round(1)
    df_tabla = df_tabla.rename(columns=RENAME_MAP_BLOQUES)
    tabla_html = df_tabla.to_html(classes='display wrap', table_id='tablaEstudiantes', index=False)

    html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<title>Reporte de Fundamentos</title>
<link rel='stylesheet' href='https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css'>
<script src='https://cdn.plot.ly/plotly-2.24.1.min.js'></script>
<style>
  body{{font-family:'Segoe UI',sans-serif;margin:0;padding:20px;background:#f0f4f8;}}
  h1{{color:{C_DEEP_NAVY};text-align:center;border-bottom:3px solid {C_ROYAL_BLUE};padding-bottom:10px;margin-bottom:20px;font-size:22px;}}
  h2{{color:{C_ROYAL_BLUE};margin:40px 0 16px;background:{C_SKY_BLUE}30;border-left:5px solid {C_OCEAN_BLUE};padding:10px 15px;border-radius:4px;font-size:17px;}}
  .tab{{overflow:hidden;background:#fff;border-radius:8px 8px 0 0;border:1px solid #ccc;display:flex;}}
  .tab button{{background:inherit;border:none;outline:none;cursor:pointer;padding:14px 24px;transition:.3s;font-size:15px;font-weight:bold;color:{C_STEEL_GREY};flex:1;}}
  .tab button:hover{{background:{C_SKY_BLUE}50;}}
  .tab button.active{{background:{C_ROYAL_BLUE};color:#fff;}}
  .tabcontent{{display:none;padding:25px;border:1px solid #ccc;border-top:none;background:#fff;border-radius:0 0 8px 8px;}}
  .container{{background:#fff;padding:25px;border-radius:8px;box-shadow:0 4px 12px rgba(3,59,109,.10);margin-bottom:24px;}}
  .metrics{{display:flex;justify-content:space-around;text-align:center;margin-bottom:10px;}}
  .metric-box{{background:{C_SKY_BLUE}25;padding:20px;border-radius:8px;width:30%;border-top:4px solid {C_ROYAL_BLUE};}}
  .metric-box.l{{border-top-color:{C_OCEAN_BLUE};}} .metric-box.m{{border-top-color:{C_DEEP_NAVY};}}
  .metric-title{{font-size:12px;color:{C_STEEL_GREY};text-transform:uppercase;font-weight:700;}}
  .metric-value{{font-size:30px;font-weight:700;color:{C_DEEP_NAVY};margin-top:8px;}}
  .subject-header{{padding:10px 16px;border-radius:6px;color:#fff;font-size:15px;font-weight:700;margin-bottom:16px;}}
  .sh-l{{background:{C_ROYAL_BLUE};}} .sh-m{{background:{C_DEEP_NAVY};}}
  .bloque-card{{border:1px solid {C_SKY_BLUE};border-radius:8px;margin-bottom:24px;overflow:hidden;}}
  .bloque-title{{padding:10px 16px;font-size:13px;font-weight:700;color:{C_DEEP_NAVY};background:{C_SKY_BLUE}40;border-bottom:1px solid {C_SKY_BLUE};}}
  .bloque-body{{padding:16px;}}
  .grupo-label{{font-size:12px;font-weight:600;color:{C_STEEL_GREY};text-transform:uppercase;margin:14px 0 6px;}}
  table.dataTable{{border-collapse:collapse;width:100%!important;}}
  table.dataTable th,table.dataTable td{{text-align:center;vertical-align:middle;font-size:12px;white-space:nowrap;}}
  table.dataTable th{{background:{C_DEEP_NAVY};color:#fff;font-size:11px;padding:8px 10px;}}
  table.dataTable tr:hover{{background:{C_SKY_BLUE}40!important;}}
  #tablaEstudiantes th:nth-child(2),#tablaEstudiantes td:nth-child(2){{text-align:left;min-width:250px;white-space:normal;}}
  #tablaPivotL th:first-child,#tablaPivotL td:first-child,
  #tablaPivotM th:first-child,#tablaPivotM td:first-child{{text-align:left;min-width:250px;white-space:normal;}}
</style></head><body>
<h1>📊 Reporte de resultados — Prueba Diagnóstica de Fundamentos</h1>
<div class="tab">
  <button class="tablinks active" onclick="openTab(event,'Global')">Reporte General Consolidado</button>
  <button class="tablinks" onclick="openTab(event,'Explorador')">Explorador Dinámico</button>
</div>
<div id="Global" class="tabcontent" style="display:block;">
  <div class='metrics'>
    <div class='metric-box'><div class='metric-title'>Total Estudiantes Evaluados</div>
      <div class='metric-value'>{len(df_scored)}</div></div>
    <div class='metric-box l'><div class='metric-title'>Promedio General — Lengua</div>
      <div class='metric-value'>{df_scored.get('L_total_puntaje_0_100',pd.Series([0])).mean():.1f} / 100</div></div>
    <div class='metric-box m'><div class='metric-title'>Promedio General — Matemática</div>
      <div class='metric-value'>{df_scored.get('M_total_puntaje_0_100',pd.Series([0])).mean():.1f} / 100</div></div>
  </div>
  <h2>Distribución Global de Puntajes por Asignatura y Grado</h2>
"""

    cols_melt = [c for c in ['L_total_puntaje_0_100','M_total_puntaje_0_100']
                 if c in df_scored.columns]
    if cols_melt:
        df_m = df_scored.melt(value_vars=cols_melt, var_name='Asignatura', value_name='Puntaje')
        df_m['Asignatura'] = df_m['Asignatura'].map({
            'L_total_puntaje_0_100': 'Lengua',
            'M_total_puntaje_0_100': 'Matemática',
        })
        fig_g = px.box(df_m, x='Asignatura', y='Puntaje', color='Asignatura', points='all',
                       color_discrete_map={'Lengua': C_OCEAN_BLUE, 'Matemática': C_DEEP_NAVY})
        fig_g.update_layout(yaxis_range=[-5, 105], plot_bgcolor='white',
                            paper_bgcolor='white', font_color=C_DEEP_NAVY, height=450)
        html += f"<div>{fig_g.to_html(full_html=False, include_plotlyjs='cdn')}</div>"

    for subj_col, subj_name, subj_color in [
        ('L_total_puntaje_0_100', 'Lengua',     C_OCEAN_BLUE),
        ('M_total_puntaje_0_100', 'Matemática', C_DEEP_NAVY),
    ]:
        if subj_col in df_scored.columns and 'Grado' in df_scored.columns:
            df_sl = df_scored.dropna(subset=[subj_col, 'Grado'])
            order = [g for g in ORDEN_GRADOS if g in df_sl['Grado'].unique()]
            if not df_sl.empty:
                fig2 = px.box(df_sl, x='Grado', y=subj_col, points='all',
                              title=f'Distribución de Puntajes por Grado — {subj_name}',
                              labels={subj_col: 'Puntaje (0–100)'},
                              color_discrete_sequence=[subj_color],
                              category_orders={"Grado": order})
                fig2.update_layout(yaxis_range=[-5, 105], plot_bgcolor='white',
                                   paper_bgcolor='white', font_color=C_DEEP_NAVY, height=450)
                html += f"<div>{fig2.to_html(full_html=False, include_plotlyjs=False)}</div>"

    html += f"<h2>Análisis Exhaustivo — Lengua</h2><div class='subject-header sh-l'>Prueba Diagnóstica de Lectura · Nivel 1</div>"
    for bloque in LENGUA_JERARQUIA:
        bcol = bloque['bloque_col']
        pcol = f"{bcol}_puntaje_0_100"
        if pcol in df_scored.columns:
            html += f"<div class='bloque-card'><div class='bloque-title'>{bloque['titulo']}</div><div class='bloque-body'>"
            html += _bloque_box(bcol, bloque['titulo'], C_OCEAN_BLUE)
            for grupo in bloque['grupos']:
                gitems = [it for it in grupo['items'] if it in df_scored.columns]
                if gitems:
                    g_vals = [df_scored[it].mean() * 100 for it in gitems]
                    g_labs = [ITEM_LABELS.get(it, it) for it in gitems]
                    html += f"<div class='grupo-label'>{grupo['label']}</div>"
                    html += _bar_items(gitems, g_labs, g_vals,
                                       f"% Acierto — {grupo['label']}", C_ROYAL_BLUE)
            html += "</div></div>"

    html += f"<h2>Análisis Exhaustivo — Matemática</h2><div class='subject-header sh-m'>Prueba de Fundamentos Matemáticas</div>"
    for bloque in MATEMATICA_JERARQUIA:
        bcol = bloque['bloque_col']
        pcol = f"{bcol}_puntaje_0_100"
        if pcol in df_scored.columns:
            html += f"<div class='bloque-card'><div class='bloque-title'>{bloque['titulo']}</div><div class='bloque-body'>"
            html += _bloque_box(bcol, bloque['titulo'], C_DEEP_NAVY)
            for grupo in bloque['grupos']:
                gitems = [it for it in grupo['items'] if it in df_scored.columns]
                if gitems:
                    g_vals = [df_scored[it].mean() * 100 for it in gitems]
                    g_labs = [ITEM_LABELS.get(it, it) for it in gitems]
                    html += f"<div class='grupo-label'>{grupo['label']}</div>"
                    html += _bar_items(gitems, g_labs, g_vals,
                                       f"% Acierto — {grupo['label']}", C_STEEL_GREY)
            html += "</div></div>"

    html += f"""
<h2>Comparativo por Escuela y Grado</h2>
<div style='overflow-x:auto;background:#fff;padding:20px;border-radius:8px;border:1px solid {C_SKY_BLUE};'>
  <h3 style='color:{C_OCEAN_BLUE};font-size:15px;border-bottom:1px solid {C_SKY_BLUE};padding-bottom:5px;'>
    Promedio General — Lengua</h3>
  {df_pivot_l.to_html(classes='display wrap', table_id='tablaPivotL', index=False)}
  <br><br>
  <h3 style='color:{C_DEEP_NAVY};font-size:15px;border-bottom:1px solid {C_SKY_BLUE};padding-bottom:5px;'>
    Promedio General — Matemática</h3>
  {df_pivot_m.to_html(classes='display wrap', table_id='tablaPivotM', index=False)}
</div>
<h2>Base de Datos de Estudiantes</h2>
<div style='overflow-x:auto;background:#fff;padding:20px;border-radius:8px;'>
  <p style='color:{C_STEEL_GREY};font-size:13px;margin-bottom:12px;'>
    <i>Escribe en las cajas debajo de cada columna para filtrar la información.</i></p>
  {tabla_html}
</div>
</div>"""

    html += f"""
<div id="Explorador" class="tabcontent">
  <h2>Explorador Interactivo por Ítem y Grado</h2>
  <p style='color:{C_STEEL_GREY};font-size:14px;margin-bottom:20px;'>
    <i>Seleccione "Todas las escuelas" para ver el desempeño nacional por grado, o un Centro
    Educativo específico. Posicione el cursor sobre cualquier barra para ver el porcentaje
    exacto y el N (estudiantes evaluados).</i></p>
  <div style='margin-bottom:30px;'>
    <label style='font-weight:bold;color:{C_DEEP_NAVY};font-size:16px;'>Seleccione un filtro:</label><br>
    <select id='schoolSelect' style='width:50%;padding:10px;border:2px solid {C_SKY_BLUE};border-radius:4px;font-size:14px;'></select>
  </div>
  <div id='dynamicPlotL' style='width:100%;height:2800px;margin-bottom:60px;'></div>
  <div id='dynamicPlotM' style='width:100%;height:1000px;'></div>
</div>
<script src='https://code.jquery.com/jquery-3.7.0.min.js'></script>
<script src='https://cdn.datatables.net/1.13.6/js/jquery.dataTables.js'></script>
<script>
function openTab(evt,tabName){{
  document.querySelectorAll('.tabcontent').forEach(t=>t.style.display='none');
  document.querySelectorAll('.tablinks').forEach(b=>b.classList.remove('active'));
  document.getElementById(tabName).style.display='block';
  evt.currentTarget.classList.add('active');
  $(window).trigger('resize');
  if(tabName==='Explorador') updateDynamicPlots();
}}

var db={json_data};
var itemLabels={json_labels};
var itemsL={json_items_l};
var itemsM={json_items_m};
var gradeOrder=['Primer Grado','Segundo Grado','Tercer Grado','Cuarto Grado','Quinto Grado',
  'Sexto Grado','Séptimo Grado','Octavo Grado','Noveno Grado','1er Año','2do Año'];

var schoolSelect=document.getElementById('schoolSelect');
var keys=Object.keys(db).sort();
var idxAll=keys.indexOf("-- Todas las escuelas --");
if(idxAll>-1){{keys.splice(idxAll,1);keys.unshift("-- Todas las escuelas --");}}
keys.forEach(function(s){{
  var opt=document.createElement('option');opt.value=s;opt.innerHTML=s;schoolSelect.appendChild(opt);
}});
schoolSelect.addEventListener('change',updateDynamicPlots);

function updateDynamicPlots(){{
  var school=schoolSelect.value;
  if(!school||!db[school]) return;
  var grades=Object.keys(db[school]).sort((a,b)=>gradeOrder.indexOf(a)-gradeOrder.indexOf(b));
  var yL=itemsL.map(i=>itemLabels[i]||i).reverse();
  var yM=itemsM.map(i=>itemLabels[i]||i).reverse();
  var tracesL=[],tracesM=[];
  var colors=['#094b93','#3077b9','#a7d2f2','#666766','#033b6d'];
  grades.forEach(function(grade,idx){{
    var row=db[school][grade];var N=row['N']||'?';
    var xL=itemsL.map(i=>row[i]||0).reverse();
    var xM=itemsM.map(i=>row[i]||0).reverse();
    var c=colors[idx%colors.length];
    tracesL.push({{x:xL,y:yL,name:grade,type:'bar',orientation:'h',marker:{{color:c}},
      hovertemplate:'<b>%{{y}}</b><br>Grado: '+grade+'<br>%{{x:.1f}}%<br>N='+N+'<extra></extra>'}});
    tracesM.push({{x:xM,y:yM,name:grade,type:'bar',orientation:'h',marker:{{color:c}},
      hovertemplate:'<b>%{{y}}</b><br>Grado: '+grade+'<br>%{{x:.1f}}%<br>N='+N+'<extra></extra>'}});
  }});
  var base={{plot_bgcolor:'white',paper_bgcolor:'white',font:{{color:'{C_DEEP_NAVY}'}},
    xaxis:{{title:'% de Aciertos',range:[0,115],gridcolor:'{C_SKY_BLUE}'}},
    yaxis:{{tickfont:{{size:11}},tickmode:'linear',dtick:1}},
    margin:{{l:280,r:40,t:50,b:40}},barmode:'group',legend:{{orientation:'h',y:1.02,x:0}}}};
  Plotly.newPlot('dynamicPlotL',tracesL,
    Object.assign({{title:'Rendimiento por Ítem — <b>LENGUA</b>',height:2800}},base),{{responsive:true}});
  Plotly.newPlot('dynamicPlotM',tracesM,
    Object.assign({{title:'Rendimiento por Ítem — <b>MATEMÁTICA</b>',height:1000}},base),{{responsive:true}});
}}

$(document).ready(function(){{
  $('#tablaPivotL,#tablaPivotM').DataTable({{
    language:{{url:'//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json'}},
    paging:false,info:false,searching:false,scrollX:false,autoWidth:false
  }});
  $('#tablaEstudiantes thead tr').clone(true).appendTo('#tablaEstudiantes thead');
  $('#tablaEstudiantes thead tr:eq(1) th').each(function(i){{
    $(this).removeClass('sorting sorting_asc sorting_desc');
    $(this).html('<input type="text" placeholder="Filtrar" style="width:100%;font-size:11px;padding:4px;"/>');
    $('input',this).on('keyup change',function(){{
      if(table.column(i).search()!==this.value) table.column(i).search(this.value).draw();
    }});
  }});
  var table=$('#tablaEstudiantes').DataTable({{
    orderCellsTop:true,
    language:{{url:'//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json'}},
    pageLength:15,scrollX:true,autoWidth:true,
    initComplete:function(){{setTimeout(function(){{table.columns.adjust().draw();}},50);}}
  }});
  $(window).on('resize',function(){{
    try{{Plotly.Plots.resize('dynamicPlotL');Plotly.Plots.resize('dynamicPlotM');}}catch(e){{}}
  }});
}});
</script></body></html>"""

    with open(ruta_salida, 'w', encoding='utf-8') as f:
        f.write(html)


# ==========================================
# 8. NÚCLEO ETL
# ==========================================
def _ejecutar_etl(current_month_path):
    PATH_FORMULARIOS = os.path.join(current_month_path, "01_Formularios_Crudos")
    PATH_DATASETS    = os.path.join(current_month_path, "02_Datasets_Procesados")
    os.makedirs(PATH_FORMULARIOS, exist_ok=True)
    os.makedirs(PATH_DATASETS,    exist_ok=True)

    archivos_form = glob.glob(os.path.join(PATH_FORMULARIOS, "*.csv"))
    if not archivos_form:
        return None, None, None, None, None, None, None

    archivo_reciente = max(archivos_form, key=os.path.getmtime)
    with open(archivo_reciente, 'r', encoding='utf-8', errors='ignore') as f:
        sep = ';' if ';' in f.readline() else ','

    df_form = pd.read_csv(archivo_reciente, dtype=str, encoding_errors='ignore', sep=sep)
    df_form.columns = df_form.columns.str.strip()

    col_email = next((c for c in df_form.columns
                      if 'correo' in c.lower() or 'email' in c.lower()), None)
    df_form = df_form[
        df_form[col_email].astype(str).str.endswith('@clases.edu.sv', na=False)
    ].copy()
    df_form['NIE'] = df_form[col_email].astype(str).apply(lambda x: x.split('@')[0].strip())

    columnas_items = df_form.columns[3:89]
    df_form = df_form.rename(columns={
        col: (f'L_item_{i+1}' if i < 68 else f'M_item_{i - 68 + 1}')
        for i, col in enumerate(columnas_items)
    })

    archivos_mat = glob.glob(os.path.join(PATH_METADATA, "MatriculaProgresoMes*.csv"))
    def ext_num(r):
        m = re.search(r'\d+', os.path.basename(r))
        return int(m.group()) if m else 0
    archivo_mat = max(archivos_mat, key=ext_num)
    df_mat = pd.read_csv(archivo_mat, sep=';', dtype=str, encoding_errors='ignore')

    col_nie   = next(c for c in df_mat.columns if 'nie'    in c.lower())
    col_cod   = next((c for c in df_mat.columns if 'codigo' in c.lower() or 'código' in c.lower()), None)
    col_nom   = next(c for c in df_mat.columns if 'nombre' in c.lower() and 'secc' not in c.lower())
    col_grado = next(c for c in df_mat.columns if 'grado'  in c.lower())
    col_grupo = next(c for c in df_mat.columns if 'nombre_secc' in c.lower())

    df_meta = df_mat[[col_nie, col_cod, col_nom, col_grado, col_grupo]].copy()
    df_meta.columns = ['NIE','Código de infraestructura','Nombre del centro','Grado','Grupo']
    df_meta['NIE']  = df_meta['NIE'].astype(str).str.replace('.0','',regex=False).str.strip()

    df_cruzado = pd.merge(df_meta, df_form, on='NIE', how='right')
    df_cruzado.fillna({
        'Código de infraestructura': 'SIN REGISTRO',
        'Nombre del centro':         'SIN REGISTRO',
        'Grado':                     'SIN REGISTRO',
        'Grupo':                     'SIN REGISTRO',
    }, inplace=True)

    col_puntaje = next((c for c in df_cruzado.columns
                        if 'puntuaci' in c.lower() or 'score' in c.lower()), None)
    df_cruzado['Puntaje_Global_Formulario'] = (
        df_cruzado[col_puntaje].astype(str).str.split('/').str[0].str.strip()
        if col_puntaje else "0"
    )

    todas_preguntas = (
        [f'L_item_{i}' for i in range(1, 69)] +
        [f'M_item_{i}' for i in range(1, 19)]
    )
    cols_orden = (
        ['Código de infraestructura','Nombre del centro','Grado','Grupo',
         'NIE', col_email, 'Puntaje_Global_Formulario']
        + todas_preguntas
    )
    cols_orden = [c for c in cols_orden if c in df_cruzado.columns]
    df_raw    = df_cruzado[cols_orden].copy()
    df_scored = df_cruzado.copy()

    # ── Análisis de distractores (sobre respuestas crudas) ────────────────
    distractores_dict = {}
    df_val = df_raw[df_raw['Nombre del centro'] != 'SIN REGISTRO'].copy()

    def generar_df_distractores(df_subset):
        rows = []
        for item in todas_preguntas:
            if item not in df_subset.columns:
                continue
            asig = 'Lengua' if item.startswith('L_') else 'Matemática'
            s    = df_subset[item].fillna('SIN RESPONDER').astype(str).str.strip()
            if len(s) == 0:
                continue
            vc = s.value_counts(normalize=True) * 100
            row_dict = {'Asignatura': asig, 'Ítem': item}
            for i, (opt, pct) in enumerate(vc.items(), 1):
                row_dict[f'Opción {i}'] = f"{opt} ({pct:.1f}%)"
            rows.append(row_dict)
        return pd.DataFrame(rows)

    distractores_dict['Todas las escuelas'] = generar_df_distractores(df_val)
    for grado, sheet_name in [('Segundo Grado','2do'),('Tercer Grado','3er'),('Cuarto Grado','4to')]:
        df_g = df_val[df_val['Grado'] == grado]
        if not df_g.empty:
            distractores_dict[sheet_name] = generar_df_distractores(df_g)

    # ── Calificación dicotómica (0 / 1) ──────────────────────────────────
    for item, correcta in CLAVE_RESPUESTAS.items():
        if item in df_scored.columns:
            df_scored[item] = np.where(
                df_scored[item].astype(str).str.strip().str.lower() == correcta.lower(),
                1, 0,
            )

    # ── Puntajes por bloque ───────────────────────────────────────────────
    for seccion, items in TODAS_SECCIONES.items():
        items_presentes = [it for it in items if it in df_scored.columns]
        df_scored[f'{seccion}_puntaje_0_100'] = calcular_puntaje_porcentual(
            df_scored[items_presentes].sum(axis=1), len(items)
        )

    items_L = [c for c in df_scored.columns if c.startswith('L_item_')]
    items_M = [c for c in df_scored.columns if c.startswith('M_item_')]

    df_scored['L_total_suma_correcta'] = df_scored[items_L].sum(axis=1) if items_L else 0
    df_scored['M_total_suma_correcta'] = df_scored[items_M].sum(axis=1) if items_M else 0
    df_scored['L_total_puntaje_0_100'] = calcular_puntaje_porcentual(
        df_scored['L_total_suma_correcta'], len(items_L) if items_L else 1)
    df_scored['M_total_puntaje_0_100'] = calcular_puntaje_porcentual(
        df_scored['M_total_suma_correcta'], len(items_M) if items_M else 1)

    # ── QC / auditoría ────────────────────────────────────────────────────
    df_qc = df_scored[['NIE','Nombre del centro','Grado','Grupo',
                        'Puntaje_Global_Formulario']].copy()
    df_qc['Total_Aciertos_Calculados'] = (df_scored['L_total_suma_correcta']
                                          + df_scored['M_total_suma_correcta'])
    df_qc['Total_Items_Prueba'] = len(items_L) + len(items_M)
    df_qc['Puntaje_Formulario_Num'] = pd.to_numeric(
        df_qc['Puntaje_Global_Formulario'], errors='coerce').fillna(0)
    df_qc['¿Hay Discrepancia?'] = np.where(
        df_qc['Puntaje_Formulario_Num'] != df_qc['Total_Aciertos_Calculados'], 'SÍ', 'NO')
    df_qc = df_qc.drop(columns=['Puntaje_Formulario_Num'])

    # ── Resumen por escuela ───────────────────────────────────────────────
    df_validos  = df_scored[df_scored['Nombre del centro'] != 'SIN REGISTRO'].copy()
    evaluados   = df_validos.groupby('Nombre del centro').size().reset_index(name='Estudiantes Evaluados')
    cols_punt   = [c for c in df_scored.columns if 'puntaje_0_100' in c]
    promedios   = df_validos.groupby('Nombre del centro')[cols_punt].mean().reset_index()
    grados_eval = df_validos['Grado'].dropna().unique()
    df_meta_u   = df_meta.drop_duplicates(subset=['NIE'])
    esperados   = (df_meta_u[df_meta_u['Grado'].isin(grados_eval)]
                   .groupby('Nombre del centro').size()
                   .reset_index(name='Estudiantes Esperados'))

    df_escuelas = pd.merge(evaluados, promedios, on='Nombre del centro', how='left')
    df_escuelas = pd.merge(df_escuelas, esperados, on='Nombre del centro', how='left')
    codigos     = (df_validos[['Nombre del centro','Código de infraestructura']]
                   .drop_duplicates(subset=['Nombre del centro']))
    df_escuelas = pd.merge(df_escuelas, codigos, on='Nombre del centro', how='left')
    df_escuelas['Estudiantes Esperados'] = df_escuelas['Estudiantes Esperados'].fillna(
        df_escuelas['Estudiantes Evaluados'])
    df_escuelas['Cobertura (%)'] = np.where(
        df_escuelas['Estudiantes Esperados'] > 0,
        (df_escuelas['Estudiantes Evaluados'] / df_escuelas['Estudiantes Esperados'] * 100),
        0).round(1).clip(0, 100.0)

    for col in cols_punt:
        df_escuelas[col] = df_escuelas[col].round(1)
    cols_order  = (['Código de infraestructura','Nombre del centro',
                    'Estudiantes Esperados','Estudiantes Evaluados','Cobertura (%)']
                   + cols_punt)
    df_escuelas = df_escuelas[[c for c in cols_order if c in df_escuelas.columns]]

    # ── Base de datos de estudiantes ──────────────────────────────────────
    df_estudiantes = df_scored.copy()
    cols_show = (['NIE','Nombre del centro','Grado','Grupo',
                  'L_total_puntaje_0_100','M_total_puntaje_0_100']
                 + [c for c in df_estudiantes.columns
                    if 'puntaje_0_100' in c and 'total' not in c])
    df_estudiantes = df_estudiantes[[c for c in cols_show if c in df_estudiantes.columns]]
    for c in df_estudiantes.select_dtypes(include=['float64']).columns:
        df_estudiantes[c] = df_estudiantes[c].round(1)
    df_estudiantes = df_estudiantes.rename(columns=RENAME_MAP_BLOQUES)

    df_mapping = pd.DataFrame([{'Item': k, 'Proceso': v}
                                for k, v in TODO_PROCESS_MAP.items()])
    paths = {
        'raw_df':         df_raw,
        'scored':         df_scored,
        'mapping':        df_mapping,
        'escuelas':       df_escuelas,
        'qc':             df_qc,
        'html':           os.path.join(PATH_DATASETS, "4_Reporte_Grafico_Fundamentos.html"),
        'xl_escuelas':    os.path.join(PATH_DATASETS, "5_Resultados_Escuelas_Fundamentos.xlsx"),
        'xl_raw':         os.path.join(PATH_DATASETS, "1_Resultados_Fundamento_Respuestas_Crudas.xlsx"),
        'xl_scored':      os.path.join(PATH_DATASETS, "2_Resultados_Fundamento_Dicotomico_0_1.xlsx"),
        'xl_mapping':     os.path.join(PATH_DATASETS, "3_Item_Process_Mapping.xlsx"),
        'xl_qc':          os.path.join(PATH_DATASETS, "6_QC_Auditoria_Puntajes.xlsx"),
        'xl_estudiantes': os.path.join(PATH_DATASETS, "7_Base_Datos_Estudiantes.xlsx"),
        'xl_distractores':os.path.join(PATH_DATASETS, "8_Analisis_Distractores_Por_Item.xlsx"),
    }
    return df_scored, df_mapping, paths, df_escuelas, df_qc, df_estudiantes, distractores_dict


def procesar_y_generar_excel(current_month_path):
    print("\n=======================================================")
    print(f"  PROCESANDO: {os.path.basename(current_month_path)}")
    print("=======================================================\n")
    df_scored, df_mapping, paths, df_escuelas, df_qc, df_estudiantes, distractores_dict = \
        _ejecutar_etl(current_month_path)
    if df_scored is None:
        return False

    paths['raw_df'].to_excel(paths['xl_raw'],         index=False)
    df_scored.to_excel(paths['xl_scored'],             index=False)
    df_mapping.to_excel(paths['xl_mapping'],           index=False)
    df_escuelas.to_excel(paths['xl_escuelas'],         index=False)
    df_qc.to_excel(paths['xl_qc'],                    index=False)
    df_estudiantes.to_excel(paths['xl_estudiantes'],   index=False)

    exportar_distractores(distractores_dict, paths['xl_distractores'], CLAVE_RESPUESTAS)

    print("[*] Generando Dashboard HTML Interactivo...")
    generar_html_estatico(df_scored, df_mapping, paths['html'])
    print(f"[OK] Todos los reportes generados exitosamente en:\n     {os.path.dirname(paths['html'])}")
    return True