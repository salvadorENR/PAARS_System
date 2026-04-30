import os
import glob
import sys
import pandas as pd
import numpy as np
import warnings
from pandas.errors import PerformanceWarning

# Silenciar warnings no críticos de Pandas
warnings.filterwarnings('ignore', category=PerformanceWarning)

# ==========================================
# 1. CONFIGURACIÓN COMÚN
# ==========================================
ROOT_DIR = r"G:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse"
YEAR_DIR = os.path.join(ROOT_DIR, "2026")
PATH_METADATA = os.path.join(ROOT_DIR, "00_Metadata")

# ==========================================
# 2. CLAVE DE RESPUESTAS OFICIAL
# ==========================================
CLAVE_RESPUESTAS = {
    # ------------------ ÍTEMS DE LENGUA (68) ------------------
    # BLOQUE 0 — Reconocimiento del Alfabeto (Grafemas)
    # Grupo 1: Vocales (ítems 1–5)
    'L_item_1': 'b',  # A
    'L_item_2': 'c',  # E
    'L_item_3': 'b',  # I
    'L_item_4': 'c',  # O
    'L_item_5': 'b',  # U
    # Grupo 2: Consonantes de Alta Frecuencia (ítems 6–13)
    'L_item_6':  'c', # M
    'L_item_7':  'b', # P
    'L_item_8':  'c', # S
    'L_item_9':  'c', # L
    'L_item_10': 'b', # T
    'L_item_11': 'c', # N
    'L_item_12': 'c', # D
    'L_item_13': 'b', # R
    # Grupo 3: Consonantes de Confusión Visual y Baja Frecuencia (ítems 14–27)
    'L_item_14': 'b', # B
    'L_item_15': 'b', # C
    'L_item_16': 'b', # F
    'L_item_17': 'c', # G
    'L_item_18': 'a', # H
    'L_item_19': 'c', # J
    'L_item_20': 'b', # K
    'L_item_21': 'b', # Ñ
    'L_item_22': 'c', # Q
    'L_item_23': 'b', # V
    'L_item_24': 'c', # W
    'L_item_25': 'b', # X
    'L_item_26': 'c', # Y
    'L_item_27': 'c', # Z
    # BLOQUE 2 — Asociación Sonido–Letra (18 Fonemas, ítems 28–45)
    # Subgrupo 1: Vocales (ítems 28–32)
    'L_item_28': 'c', # /a/
    'L_item_29': 'c', # /e/
    'L_item_30': 'b', # /i/
    'L_item_31': 'c', # /o/
    'L_item_32': 'a', # /u/
    # Subgrupo 2: Consonantes de Alta Frecuencia (ítems 33–39)
    'L_item_33': 'c', # /m/
    'L_item_34': 'b', # /p/
    'L_item_35': 'a', # /s/
    'L_item_36': 'b', # /l/
    'L_item_37': 'a', # /t/
    'L_item_38': 'c', # /n/
    'L_item_39': 'b', # /d/
    # Subgrupo 3: Consonantes de Menor Frecuencia o Mayor Dificultad (ítems 40–45)
    'L_item_40': 'b', # /rr/
    'L_item_41': 'a', # /r/
    'L_item_42': 'b', # /g/
    'L_item_43': 'b', # /f/
    'L_item_44': 'c', # /b/
    'L_item_45': 'b', # /k/
    # BLOQUE 3 — Conciencia Fonológica y Principio Alfabético (ítems 46–50)
    'L_item_46': 'b', # Sonido inicial: sopa
    'L_item_47': 'a', # Rima: taza
    'L_item_48': 'a', # Segmentación: /s/–/o/–/l/
    'L_item_49': 'c', # Fusión de sonidos: pan
    'L_item_50': 'b', # Sonido final: l
    # BLOQUE 2 — Decodificación y Fluidez Básica (ítems 51–58)
    'L_item_51': 'a', # Decodificación sílaba CV: pa
    'L_item_52': 'c', # Decodificación palabra CVC: sol
    'L_item_53': 'a', # Pseudopalabra: mufe
    'L_item_54': 'b', # Identificación de palabra: palo
    'L_item_55': 'c', # Fluidez: imagen de un sol
    'L_item_56': 'a', # Lectura funcional: camino
    'L_item_57': 'b', # Comprensión literal personaje: un perro
    'L_item_58': 'a', # Comprensión literal acción: salta
    # BLOQUE 3 — Comprensión Lectora Literal e Inferencial (ítems 59–62)
    'L_item_59': 'a', # Literal personaje: un perro
    'L_item_60': 'a', # Literal acción: salta
    'L_item_61': 'a', # Inferencial causa: va a llover
    'L_item_62': 'b', # Relación con experiencia: niño leyendo
    # BLOQUE 4 — Comprensión Auditiva (ítems 63–66)
    'L_item_63': 'c', # Secuencia básica: camina al parque (Ana primero → opción c)
    'L_item_64': 'b', # Información explícita: el perro hizo el ruido
    'L_item_65': 'a', # Inferencia emocional: triste
    'L_item_66': 'c', # Causa o motivo: porque empieza a llover
    # Comprensión Lectora Escrita — 2° y 3° grado (ítems 67–68 → del bloque escrito final)
    'L_item_67': '2)', # Nivel LITERAL: Nube (¿Cómo se llama el gato?)
    'L_item_68': '3)', # Nivel CRÍTICO: Actuó de forma honesta

    # ------------------ ÍTEMS DE MATEMÁTICA (18) ------------------
    'M_item_1': '3) 9', 'M_item_2': '2) 15', 'M_item_3': '2) 18 es mayor que 15',
    'M_item_4': '2) 25', 'M_item_5': '1) 4 decenas y 8 unidades', 'M_item_6': '1) 34',
    'M_item_7': '1) 67 < 76', 'M_item_8': '2) 25, 40 y 52', 'M_item_9': '2) Decenas',
    'M_item_10': '3) 205', 'M_item_11': '2) 7', 'M_item_12': '1) 4', 'M_item_13': '3) 57',
    'M_item_14': '2) 33', 'M_item_15': '3) 23', 'M_item_16': '2) 19', 'M_item_17': '3) 16',
    'M_item_18': '2) 12'
}

# ==========================================
# 3. MAPEOS PEDAGÓGICOS
# ==========================================
L_SECCIONES = {
    # BLOQUE 0 — Reconocimiento del Alfabeto (Grafemas) → ítems 1–27
    'L_Bloque0_Reconocimiento_del_Alfabeto': [f'L_item_{i}' for i in range(1, 28)],
    # BLOQUE 2 — Asociación Sonido–Letra (18 Fonemas) → ítems 28–45
    'L_Bloque2_Asociacion_Sonido_Letra': [f'L_item_{i}' for i in range(28, 46)],
    # BLOQUE 3 — Conciencia Fonológica y Principio Alfabético → ítems 46–50
    'L_Bloque3_Conciencia_Fonologica_y_Principio_Alfabetico': [f'L_item_{i}' for i in range(46, 51)],
    # BLOQUE 2 — Decodificación y Fluidez Básica → ítems 51–58
    'L_Bloque2b_Decodificacion_y_Fluidez_Basica': [f'L_item_{i}' for i in range(51, 59)],
    # BLOQUE 3 — Comprensión Lectora Literal e Inferencial → ítems 59–62
    'L_Bloque3b_Comprension_Lectora_Literal_e_Inferencial': [f'L_item_{i}' for i in range(59, 63)],
    # BLOQUE 4 — Comprensión Auditiva → ítems 63–66
    'L_Bloque4_Comprension_Auditiva': [f'L_item_{i}' for i in range(63, 67)],
    # Ítems de Comprensión Lectora Escrita (2° y 3° grado) → ítems 67–68
    'L_Bloque5_Comprension_Lectora_Escrita': [f'L_item_{i}' for i in range(67, 69)]
}
M_SECCIONES = {
    # BLOQUE 1 — Conteo y Sentido Numérico (4 Ítems) → M_item_1–4
    'M_Bloque1_Conteo_y_Sentido_Numerico': [f'M_item_{i}' for i in range(1, 5)],
    # BLOQUE 2 — Valor Posicional y Sistema Decimal (6 Ítems) → M_item_5–10
    'M_Bloque2_Valor_Posicional_y_Sistema_Decimal': [f'M_item_{i}' for i in range(5, 11)],
    # BLOQUE 3 — Operaciones Aditivas y Sustractivas (8 Ítems) → M_item_11–18
    'M_Bloque3_Operaciones_Aditivas_y_Sustractivas': [f'M_item_{i}' for i in range(11, 19)]
}
TODAS_SECCIONES = {**L_SECCIONES, **M_SECCIONES}

L_PROCESS_MAP = {
    # BLOQUE 0 — Reconocimiento del Alfabeto
    'L_item_1':  'Reconocimiento de Vocales',
    'L_item_6':  'Consonantes de Alta Frecuencia',
    'L_item_14': 'Consonantes de Confusión Visual y Baja Frecuencia',
    # BLOQUE 2 — Asociación Sonido–Letra
    'L_item_28': 'Fonemas Vocálicos',
    'L_item_33': 'Fonemas de Alta Frecuencia',
    'L_item_40': 'Fonemas de Menor Frecuencia o Mayor Dificultad',
    # BLOQUE 3 — Conciencia Fonológica y Principio Alfabético
    'L_item_46': 'Sonido Inicial',
    'L_item_47': 'Rima',
    'L_item_48': 'Segmentación de Sonidos',
    'L_item_49': 'Fusión de Sonidos',
    'L_item_50': 'Identificación de Sonido Final',
    # BLOQUE 2 — Decodificación y Fluidez Básica
    'L_item_51': 'Decodificación de Sílaba Directa (CV)',
    'L_item_52': 'Decodificación de Palabra (CVC)',
    'L_item_53': 'Pseudopalabra Simple',
    'L_item_54': 'Identificación de Palabra entre Opciones',
    'L_item_55': 'Fluidez: Reconocer Palabras por Sonido',
    'L_item_56': 'Fluidez: Lectura Funcional de Palabras',
    'L_item_57': 'Comprensión Literal (Personaje)',
    'L_item_58': 'Comprensión Literal (Acción)',
    # BLOQUE 3 — Comprensión Lectora Literal e Inferencial
    'L_item_59': 'Comprensión Literal (Personaje)',
    'L_item_60': 'Comprensión Literal (Acción)',
    'L_item_61': 'Comprensión Inferencial: Causa Simple',
    'L_item_62': 'Comprensión Inferencial: Relación con Experiencia',
    # BLOQUE 4 — Comprensión Auditiva
    'L_item_63': 'Comprensión Auditiva: Secuencia Básica',
    'L_item_64': 'Comprensión Auditiva: Información Explícita',
    'L_item_65': 'Comprensión Auditiva: Inferencia Emocional',
    'L_item_66': 'Comprensión Auditiva: Causa o Motivo',
    # Comprensión Lectora Escrita
    'L_item_67': 'Comprensión Lectora Escrita: Nivel Literal',
    'L_item_68': 'Comprensión Lectora Escrita: Nivel Crítico/Valorativo',
}

M_PROCESS_MAP = {
    'M_item_1': 'Conteo visual', 'M_item_2': 'Secuencia numérica',
    'M_item_3': 'Comparación numérica', 'M_item_4': 'Patrón numérico',
    'M_item_5': 'Descomposición DU', 'M_item_6': 'Convertir DU a número',
    'M_item_7': 'Comparación con signos', 'M_item_8': 'Ordenar números',
    'M_item_9': 'Valor posicional', 'M_item_10': 'Construcción C-D-U',
    'M_item_11': 'Suma básica', 'M_item_12': 'Resta básica',
    'M_item_13': 'Suma de dos cifras', 'M_item_14': 'Resta sin préstamo',
    'M_item_15': 'Resta con préstamo', 'M_item_16': 'Problema aditivo',
    'M_item_17': 'Problema sustractivo', 'M_item_18': 'Multiplicación inicial'
}
TODO_PROCESS_MAP = {**L_PROCESS_MAP, **M_PROCESS_MAP}


# ==========================================
# 4. FUNCIÓN DE PUNTUACIÓN PORCENTUAL
#    Fórmula: Puntaje = (Respuestas correctas x 100) / Total de ítems
# ==========================================
def calcular_puntaje_porcentual(suma_correcta_serie, total_items):
    """
    Convierte la suma de respuestas correctas a una escala de 0 a 100.
    Fórmula: Puntaje = (Respuestas correctas * 100) / Total de ítems de la prueba
    """
    if total_items == 0:
        return pd.Series([0.0] * len(suma_correcta_serie), index=suma_correcta_serie.index)
    return (suma_correcta_serie / total_items * 100).clip(0, 100)


# ==========================================
# 5. GENERADOR DE HTML INTERACTIVO
# ==========================================

# ── Jerarquía completa de análisis — nombres exactos del documento ────────────
#    Cada nivel produce sus propias gráficas independientemente del instrumento.

LENGUA_JERARQUIA = [
    {
        'bloque_col':  'L_Bloque0_Reconocimiento_del_Alfabeto',
        'titulo':      'Bloque 0 — Reconocimiento del Alfabeto (Grafemas)',
        'grupos': [
            {'label': 'Grupo 1: Vocales',
             'items': [f'L_item_{i}' for i in range(1, 6)]},
            {'label': 'Grupo 2: Consonantes de Alta Frecuencia',
             'items': [f'L_item_{i}' for i in range(6, 14)]},
            {'label': 'Grupo 3: Consonantes de Confusión Visual y Baja Frecuencia',
             'items': [f'L_item_{i}' for i in range(14, 28)]},
        ]
    },
    {
        'bloque_col':  'L_Bloque2_Asociacion_Sonido_Letra',
        'titulo':      'Bloque 2 — Asociación Sonido–Letra (18 Fonemas)',
        'grupos': [
            {'label': 'Subgrupo 1: Vocales (Inicio/Final de palabra)',
             'items': [f'L_item_{i}' for i in range(28, 33)]},
            {'label': 'Subgrupo 2: Consonantes de Alta Frecuencia',
             'items': [f'L_item_{i}' for i in range(33, 40)]},
            {'label': 'Subgrupo 3: Consonantes de Menor Frecuencia o Mayor Dificultad',
             'items': [f'L_item_{i}' for i in range(40, 46)]},
        ]
    },
    {
        'bloque_col':  'L_Bloque3_Conciencia_Fonologica_y_Principio_Alfabetico',
        'titulo':      'Bloque 3 — Conciencia Fonológica y Principio Alfabético',
        'grupos': [
            {'label': 'Ítems 1–5 (audio-only)',
             'items': [f'L_item_{i}' for i in range(46, 51)]},
        ]
    },
    {
        'bloque_col':  'L_Bloque2b_Decodificacion_y_Fluidez_Basica',
        'titulo':      'Bloque 2 — Decodificación y Fluidez Básica',
        'grupos': [
            {'label': 'Decodificación (ítems 6–9)',
             'items': [f'L_item_{i}' for i in range(51, 55)]},
            {'label': 'Fluidez (ítems 10–11)',
             'items': [f'L_item_{i}' for i in range(55, 57)]},
            {'label': 'Comprensión básica (ítems 12–13 de este bloque)',
             'items': [f'L_item_{i}' for i in range(57, 59)]},
        ]
    },
    {
        'bloque_col':  'L_Bloque3b_Comprension_Lectora_Literal_e_Inferencial',
        'titulo':      'Bloque 3 — Comprensión Lectora Literal e Inferencial',
        'grupos': [
            {'label': 'C1 — Recuperar información explícita',
             'items': ['L_item_59', 'L_item_60']},
            {'label': 'C2 — Hacer inferencias directas',
             'items': ['L_item_61']},
            {'label': 'C3 — Relacionar contenido del texto con experiencias',
             'items': ['L_item_62']},
        ]
    },
    {
        'bloque_col':  'L_Bloque4_Comprension_Auditiva',
        'titulo':      'Bloque 4 — Comprensión Auditiva',
        'grupos': [
            {'label': 'C1 — Escuchar textos orales breves y recordar datos o secuencias',
             'items': ['L_item_63']},
            {'label': 'C2 — Identificar personajes, lugares o hechos mencionados',
             'items': ['L_item_64']},
            {'label': 'C3 — Inferir emociones, causas o consecuencias',
             'items': ['L_item_65', 'L_item_66']},
        ]
    },
    {
        'bloque_col':  'L_Bloque5_Comprension_Lectora_Escrita',
        'titulo':      'Ítems de Comprensión Lectora — Prueba de Fundamentos 2° y 3° grado',
        'grupos': [
            {'label': 'Nivel LITERAL',
             'items': ['L_item_67']},
            {'label': 'Nivel CRÍTICO / VALORATIVO',
             'items': ['L_item_68']},
        ]
    },
]

MATEMATICA_JERARQUIA = [
    {
        'bloque_col':  'M_Bloque1_Conteo_y_Sentido_Numerico',
        'titulo':      'Bloque 1 — Conteo y Sentido Numérico (4 Ítems)',
        'grupos': [
            {'label': 'Ítems 1–4',
             'items': [f'M_item_{i}' for i in range(1, 5)]},
        ]
    },
    {
        'bloque_col':  'M_Bloque2_Valor_Posicional_y_Sistema_Decimal',
        'titulo':      'Bloque 2 — Valor Posicional y Sistema Decimal (6 Ítems)',
        'grupos': [
            {'label': 'Decenas y Unidades (D-U)',
             'items': [f'M_item_{i}' for i in range(5, 8)]},
            {'label': 'Centenas, Decenas y Unidades (C-D-U)',
             'items': [f'M_item_{i}' for i in range(8, 11)]},
        ]
    },
    {
        'bloque_col':  'M_Bloque3_Operaciones_Aditivas_y_Sustractivas',
        'titulo':      'Bloque 3 — Operaciones Aditivas y Sustractivas (8 Ítems)',
        'grupos': [
            {'label': 'Operaciones básicas (una cifra)',
             'items': ['M_item_11', 'M_item_12']},
            {'label': 'Operaciones de dos cifras',
             'items': ['M_item_13', 'M_item_14', 'M_item_15']},
            {'label': 'Problemas verbales',
             'items': ['M_item_16', 'M_item_17', 'M_item_18']},
        ]
    },
]

# ── Etiquetas de ítem legibles para el eje Y ──────────────────────────────────
ITEM_LABELS = {
    # Bloque 0
    'L_item_1': 'Letra A', 'L_item_2': 'Letra E', 'L_item_3': 'Letra I',
    'L_item_4': 'Letra O', 'L_item_5': 'Letra U',
    'L_item_6': 'Letra M', 'L_item_7': 'Letra P', 'L_item_8': 'Letra S',
    'L_item_9': 'Letra L', 'L_item_10': 'Letra T', 'L_item_11': 'Letra N',
    'L_item_12': 'Letra D', 'L_item_13': 'Letra R',
    'L_item_14': 'Letra B', 'L_item_15': 'Letra C', 'L_item_16': 'Letra F',
    'L_item_17': 'Letra G', 'L_item_18': 'Letra H', 'L_item_19': 'Letra J',
    'L_item_20': 'Letra K', 'L_item_21': 'Letra Ñ', 'L_item_22': 'Letra Q',
    'L_item_23': 'Letra V', 'L_item_24': 'Letra W', 'L_item_25': 'Letra X',
    'L_item_26': 'Letra Y', 'L_item_27': 'Letra Z',
    # Bloque 2 — Fonemas
    'L_item_28': 'Fonema /a/', 'L_item_29': 'Fonema /e/', 'L_item_30': 'Fonema /i/',
    'L_item_31': 'Fonema /o/', 'L_item_32': 'Fonema /u/',
    'L_item_33': 'Fonema /m/', 'L_item_34': 'Fonema /p/', 'L_item_35': 'Fonema /s/',
    'L_item_36': 'Fonema /l/', 'L_item_37': 'Fonema /t/', 'L_item_38': 'Fonema /n/',
    'L_item_39': 'Fonema /d/',
    'L_item_40': 'Fonema /rr/ (vibrante múltiple)', 'L_item_41': 'Fonema /r/ (vibrante simple)',
    'L_item_42': 'Fonema /g/ (sonido suave)', 'L_item_43': 'Fonema /f/',
    'L_item_44': 'Fonema /b/', 'L_item_45': 'Fonema /k/',
    # Bloque 3 — Conciencia Fonológica
    'L_item_46': 'Ítem 1 — Sonido inicial', 'L_item_47': 'Ítem 2 — Rima',
    'L_item_48': 'Ítem 3 — Segmentación de sonidos',
    'L_item_49': 'Ítem 4 — Combinar sonidos (fusión)',
    'L_item_50': 'Ítem 5 — Identificación de sonido final',
    # Bloque 2b — Decodificación y Fluidez
    'L_item_51': 'Ítem 6 — Decodificación de sílaba directa (CV)',
    'L_item_52': 'Ítem 7 — Decodificación de palabra CVC',
    'L_item_53': 'Ítem 8 — Pseudopalabra simple',
    'L_item_54': 'Ítem 9 — Identificación de palabra entre opciones',
    'L_item_55': 'Ítem 10 — Fluidez: reconocer palabras por sonido',
    'L_item_56': 'Ítem 11 — Lectura funcional de palabras',
    'L_item_57': 'Ítem 12 — Comprensión literal (personaje)',
    'L_item_58': 'Ítem 13 — Comprensión literal (acción)',
    # Bloque 3b — Comprensión Lectora
    'L_item_59': 'Ítem 12 — Comprensión literal (personaje)',
    'L_item_60': 'Ítem 13 — Comprensión literal (acción)',
    'L_item_61': 'Ítem 14 — Comprensión inferencial (causa simple)',
    'L_item_62': 'Ítem 15 — Comprensión inferencial (relación con experiencia)',
    # Bloque 4 — Comprensión Auditiva
    'L_item_63': 'Ítem 16 — Secuencia básica',
    'L_item_64': 'Ítem 17 — Información explícita',
    'L_item_65': 'Ítem 18 — Inferencia emocional',
    'L_item_66': 'Ítem 19 — Causa o motivo',
    # Comprensión Lectora Escrita
    'L_item_67': 'Ítem 20 — Nivel Literal (¿Cómo se llama el gato?)',
    'L_item_68': 'Ítem 23 — Nivel Crítico/Valorativo (¿Cómo actuó Juan?)',
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
    'M_item_11': 'Ítem 11 — Suma básica (4+3)',
    'M_item_12': 'Ítem 12 — Resta básica (9−5)',
    'M_item_13': 'Ítem 13 — Suma de dos cifras (21+36)',
    'M_item_14': 'Ítem 14 — Resta sin préstamo (54−21)',
    'M_item_15': 'Ítem 15 — Resta con préstamo (42−19)',
    'M_item_16': 'Ítem 16 — Problema aditivo',
    'M_item_17': 'Ítem 17 — Problema sustractivo',
    'M_item_18': 'Ítem 18 — Multiplicación inicial',
}


def generar_html_estatico(df_scored, df_mapping, ruta_salida):
    try:
        import plotly.express as px
        import plotly.graph_objects as go
    except ImportError:
        print("    [!] Para generar el HTML necesitas Plotly (pip install plotly).")
        return

    # ── Paleta corporativa ────────────────────────────────────────────────────
    C_DEEP_NAVY  = '#033b6d'
    C_ROYAL_BLUE = '#094b93'
    C_OCEAN_BLUE = '#3077b9'
    C_SKY_BLUE   = '#a7d2f2'
    C_STEEL_GREY = '#666766'

    # Paleta por asignatura: Lengua = azules, Matemática = teals/verdes
    COLORS_L = [C_OCEAN_BLUE, C_ROYAL_BLUE, C_DEEP_NAVY]
    COLORS_M = ['#1D9E75', '#0F6E56', '#085041']

    def _lm(labels, px_per_char=8, minimum=180):
        return max(max((len(str(l)) for l in labels), default=0) * px_per_char, minimum)

    def _bar_h(df_items, title, color, include_plotlyjs=False):
        """Bar chart horizontal % acierto por ítem."""
        labels = [ITEM_LABELS.get(r['Item'], r['Item']) for _, r in df_items.iterrows()]
        values = [df_scored[r['Item']].mean() * 100
                  if r['Item'] in df_scored.columns else 0
                  for _, r in df_items.iterrows()]
        fig = px.bar(x=values, y=labels, orientation='h',
                     title=title, text=values,
                     color_discrete_sequence=[color])
        fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside',
                          marker_color=color)
        fig.update_layout(
            xaxis_range=[0, 115],
            margin=dict(l=_lm(labels), r=70, t=50, b=30),
            plot_bgcolor='white', paper_bgcolor='white',
            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
            yaxis=dict(tickfont=dict(size=11, color=C_DEEP_NAVY), autorange='reversed'),
            xaxis=dict(title='% de Aciertos', gridcolor=C_SKY_BLUE)
        )
        return fig.to_html(full_html=False, include_plotlyjs='cdn' if include_plotlyjs else False)

    def _bloque_box(bloque_col, btitle, color, include_plotlyjs=False):
        """Box plot de distribución del puntaje de un bloque — ancho completo."""
        col = f'{bloque_col}_puntaje_0_100'
        if col not in df_scored.columns:
            return ''
        import plotly.graph_objects as go

        # Convert hex color to rgba for fill (Plotly rejects 8-digit hex)
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
        fill_rgba = f'rgba({r},{g},{b},0.2)'

        data = df_scored[col].dropna()
        fig = go.Figure()
        fig.add_trace(go.Violin(
            y=data, box_visible=True, meanline_visible=True,
            points='all', pointpos=0,
            fillcolor=fill_rgba, line_color=color,
            marker=dict(color=color, size=5, opacity=0.5),
            name='Distribución'
        ))
        fig.update_layout(
            title=f'Distribución del puntaje — {btitle}',
            yaxis=dict(title='Puntaje (0–100)', range=[0, 100],
                       gridcolor=C_SKY_BLUE, tickfont=dict(color=C_DEEP_NAVY)),
            xaxis=dict(showticklabels=False),
            plot_bgcolor='white', paper_bgcolor='white',
            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
            showlegend=False,
            height=320,
            margin=dict(l=60, r=60, t=55, b=20)
        )
        mean_val = data.mean()
        fig.add_hline(y=mean_val, line_dash='dot', line_color=C_DEEP_NAVY, opacity=0.6,
                      annotation_text=f'Media: {mean_val:.1f}',
                      annotation_position='top right',
                      annotation_font_color=C_DEEP_NAVY)
        return fig.to_html(full_html=False,
                           include_plotlyjs='cdn' if include_plotlyjs else False)

    def _bar_items(items, labels, vals, title, color, include_plotlyjs=False):
        """Bar chart horizontal con altura dinámica según número de ítems."""
        n = len(items)
        # 28px per bar + margins, minimum 200px
        height = max(200, n * 32 + 80)
        fig = px.bar(x=vals, y=labels, orientation='h', title=title,
                     text=vals, color_discrete_sequence=[color])
        fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside',
                          marker_color=color)
        fig.update_layout(
            xaxis_range=[0, 118],
            height=height,
            margin=dict(l=_lm(labels), r=80, t=50, b=30),
            plot_bgcolor='white', paper_bgcolor='white',
            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
            yaxis=dict(tickfont=dict(size=11, color=C_DEEP_NAVY),
                       autorange='reversed', fixedrange=True),
            xaxis=dict(title='% de Aciertos', gridcolor=C_SKY_BLUE,
                       ticksuffix='%')
        )
        return fig.to_html(full_html=False,
                           include_plotlyjs='cdn' if include_plotlyjs else False)

    # ── Preparar tabla interactiva ────────────────────────────────────────────
    df_tabla = df_scored.copy()
    cols_show = ['NIE', 'Nombre del centro', 'Grado', 'Grupo',
                 'L_total_puntaje_0_100', 'M_total_puntaje_0_100']
    cols_show += [c for c in df_tabla.columns if 'puntaje_0_100' in c and 'total' not in c]
    cols_show  = [c for c in cols_show if c in df_tabla.columns]
    df_tabla   = df_tabla[cols_show]
    for c in df_tabla.select_dtypes(include=['float64']).columns:
        df_tabla[c] = df_tabla[c].round(1)
    tabla_html = df_tabla.to_html(classes='display wrap', table_id='tablaEstudiantes', index=False)

    # ── HTML header ───────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Reporte Integral de Fundamentos</title>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css">
<style>
  body{{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;margin:0;padding:30px 40px;background:#f0f4f8;}}
  h1{{color:{C_DEEP_NAVY};text-align:center;border-bottom:3px solid {C_ROYAL_BLUE};padding-bottom:10px;margin-bottom:40px;font-size:22px;}}
  h2{{color:{C_ROYAL_BLUE};margin:50px 0 16px;background:{C_SKY_BLUE}30;border-left:5px solid {C_OCEAN_BLUE};padding:10px 15px;border-radius:4px;font-size:17px;}}
  h3{{color:{C_DEEP_NAVY};margin:30px 0 10px;font-size:14px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid {C_SKY_BLUE};padding-bottom:6px;}}
  h4{{color:{C_ROYAL_BLUE};margin:20px 0 8px;font-size:13px;font-weight:600;}}
  .container{{background:#fff;padding:25px;border-radius:8px;box-shadow:0 4px 12px rgba(3,59,109,.10);margin-bottom:24px;}}
  .metrics{{display:flex;justify-content:space-around;text-align:center;margin-bottom:10px;}}
  .metric-box{{background:{C_SKY_BLUE}25;padding:20px;border-radius:8px;width:30%;border-top:4px solid {C_ROYAL_BLUE};}}
  .metric-box.l{{border-top-color:{C_OCEAN_BLUE};}}
  .metric-box.m{{border-top-color:{C_DEEP_NAVY};}}
  .metric-title{{font-size:12px;color:{C_STEEL_GREY};text-transform:uppercase;font-weight:700;letter-spacing:.05em;}}
  .metric-value{{font-size:30px;font-weight:700;color:{C_DEEP_NAVY};margin-top:8px;}}
  .grid2{{display:flex;gap:16px;flex-wrap:wrap;}}
  .grid2>div{{flex:1;min-width:300px;}}
  .subject-header{{padding:10px 16px;border-radius:6px;color:#fff;font-size:15px;font-weight:700;margin-bottom:16px;}}
  .sh-l{{background:{C_ROYAL_BLUE};}}
  .sh-m{{background:{C_DEEP_NAVY};}}
  .bloque-card{{border:1px solid {C_SKY_BLUE};border-radius:8px;margin-bottom:24px;overflow:hidden;}}
  .bloque-title{{padding:10px 16px;font-size:13px;font-weight:700;color:{C_DEEP_NAVY};background:{C_SKY_BLUE}40;border-bottom:1px solid {C_SKY_BLUE};}}
  .bloque-body{{padding:16px;}}
  .grupo-label{{font-size:12px;font-weight:600;color:{C_STEEL_GREY};text-transform:uppercase;letter-spacing:.04em;margin:14px 0 6px;}}
  .summary-row{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:10px;}}
  .stat-pill{{background:{C_SKY_BLUE}30;border:1px solid {C_SKY_BLUE};border-radius:20px;padding:3px 12px;font-size:12px;color:{C_DEEP_NAVY};}}
  .stat-pill span{{font-weight:700;}}
  #tablaEstudiantes th{{background:{C_DEEP_NAVY};color:#fff;text-align:center;}}
  #tablaEstudiantes td{{text-align:center;}}
  #tablaEstudiantes tr:hover{{background:{C_SKY_BLUE}40!important;}}
  .toc{{background:#fff;border-radius:8px;padding:20px 28px;margin-bottom:32px;box-shadow:0 2px 8px rgba(3,59,109,.08);}}
  .toc ul{{list-style:none;padding:0;margin:0;columns:2;gap:30px;}}
  .toc li{{margin:4px 0;}}
  .toc a{{color:{C_OCEAN_BLUE};text-decoration:none;font-size:13px;}}
  .toc a:hover{{text-decoration:underline;}}
</style>
</head>
<body>
<h1>📊 Reporte de resultados de la aplicación de la Prueba de Fundamentos</h1>

<div class='container'>
  <div class='metrics'>
    <div class='metric-box'>
      <div class='metric-title'>Total de Estudiantes</div>
      <div class='metric-value'>{len(df_scored)}</div>
    </div>
    <div class='metric-box l'>
      <div class='metric-title'>Promedio General — Lengua</div>
      <div class='metric-value'>{df_scored['L_total_puntaje_0_100'].mean():.1f} / 100</div>
    </div>
    <div class='metric-box m'>
      <div class='metric-title'>Promedio General — Matemática</div>
      <div class='metric-value'>{df_scored['M_total_puntaje_0_100'].mean():.1f} / 100</div>
    </div>
  </div>
</div>
"""

    # ── Tabla de contenidos ────────────────────────────────────────────────────
    html += "<div class='toc'><strong style='font-size:13px;color:" + C_DEEP_NAVY + ";'>Contenido</strong><ul>"
    sec = 1
    html += f"<li><a href='#sec{sec}'>Sección {sec}. Distribución global de puntajes</a></li>"; sec += 1
    html += f"<li><a href='#sec{sec}'>Sección {sec}. Análisis exhaustivo — Lengua</a></li>"; sec += 1
    html += f"<li><a href='#sec{sec}'>Sección {sec}. Análisis exhaustivo — Matemática</a></li>"; sec += 1
    html += f"<li><a href='#sec{sec}'>Sección {sec}. Correlación entre bloques</a></li>"; sec += 1
    html += f"<li><a href='#sec{sec}'>Sección {sec}. Base de datos de estudiantes</a></li>"
    html += "</ul></div>"

    sec = 1

    # ═══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 1 — DISTRIBUCIÓN GLOBAL
    # ═══════════════════════════════════════════════════════════════════════════
    html += f"<h2 id='sec{sec}'>Sección {sec}. Distribución Global de Puntajes por Asignatura</h2>"; sec += 1

    cols_melt = [c for c in ['L_total_puntaje_0_100', 'M_total_puntaje_0_100'] if c in df_scored.columns]
    if cols_melt:
        df_m = df_scored.melt(value_vars=cols_melt, var_name='Asignatura', value_name='Puntaje')
        df_m['Asignatura'] = df_m['Asignatura'].map({
            'L_total_puntaje_0_100': 'Lengua',
            'M_total_puntaje_0_100': 'Matemática'
        })
        fig_g = px.box(df_m, x='Asignatura', y='Puntaje', color='Asignatura', points='all',
                       title='Comparativa de Puntajes: Lengua vs Matemática',
                       labels={'Puntaje': 'Puntaje (0–100)'},
                       color_discrete_map={'Lengua': C_OCEAN_BLUE, 'Matemática': C_DEEP_NAVY})
        fig_g.update_layout(yaxis_range=[0, 100], plot_bgcolor='white', paper_bgcolor='white',
                            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
                            height=450, margin=dict(l=70, r=70, t=60, b=40))
        html += f"<div class='container'>{fig_g.to_html(full_html=False, include_plotlyjs='cdn')}</div>"

    # ═══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 2 — ANÁLISIS EXHAUSTIVO LENGUA
    # ═══════════════════════════════════════════════════════════════════════════
    html += f"<h2 id='sec{sec}'>Sección {sec}. Análisis Exhaustivo — Lengua</h2>"; sec += 1
    html += "<div class='subject-header sh-l'>Prueba Diagnóstica Lenguaje · Nivel 1: Ciclos 1 y 2</div>"

    first_l = True
    for bloque in LENGUA_JERARQUIA:
        bcol   = bloque['bloque_col']
        btitle = bloque['titulo']
        grupos = bloque['grupos']

        # ── puntaje promedio del bloque ──
        pcol = f'{bcol}_puntaje_0_100'
        prom_bloque = df_scored[pcol].mean() if pcol in df_scored.columns else 0

        html += f"<div class='bloque-card'>"
        html += f"<div class='bloque-title'>{btitle}</div>"
        html += f"<div class='bloque-body'>"

        # ── pill de resumen ──
        n_items_bloque = sum(len(g['items']) for g in grupos)
        html += (f"<div class='summary-row'>"
                 f"<span class='stat-pill'>Promedio del bloque: <span>{prom_bloque:.1f} / 100</span></span>"
                 f"<span class='stat-pill'>Ítems: <span>{n_items_bloque}</span></span>"
                 f"</div>")

        # ── box plot distribución del bloque (ancho completo) ────────────────
        if pcol in df_scored.columns:
            bp = _bloque_box(bcol, btitle, C_OCEAN_BLUE, include_plotlyjs=first_l)
            first_l = False
            html += f"<div class='container'>{bp}</div>"

        # ── % acierto por ítem — bloque completo (ancho completo) ────────────
        all_items_bloque = [it for g in grupos for it in g['items'] if it in df_scored.columns]
        if all_items_bloque:
            vals_b = [df_scored[it].mean() * 100 for it in all_items_bloque]
            labs_b = [ITEM_LABELS.get(it, it) for it in all_items_bloque]
            html += (f"<div class='container'>"
                     f"{_bar_items(all_items_bloque, labs_b, vals_b, '% de Acierto por Ítem — Bloque completo', C_OCEAN_BLUE)}"
                     f"</div>")

        # ── análisis por grupo/subgrupo ──
        for grupo in grupos:
            glabel = grupo['label']
            gitems = [it for it in grupo['items'] if it in df_scored.columns]
            if not gitems:
                continue

            prom_g = sum(df_scored[it].mean() for it in gitems) / len(gitems) * 100
            html += f"<div class='grupo-label'>{glabel}</div>"
            html += (f"<div class='summary-row'>"
                     f"<span class='stat-pill'>Promedio del grupo: <span>{prom_g:.1f}%</span></span>"
                     f"<span class='stat-pill'>Ítems: <span>{len(gitems)}</span></span>"
                     f"</div>")

            g_vals = [df_scored[it].mean() * 100 for it in gitems]
            g_labs = [ITEM_LABELS.get(it, it) for it in gitems]
            html += (f"<div class='container'>"
                     f"{_bar_items(gitems, g_labs, g_vals, f'% de Acierto — {glabel}', C_ROYAL_BLUE)}"
                     f"</div>")

        html += "</div></div>"  # bloque-body + bloque-card

    # ═══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 3 — ANÁLISIS EXHAUSTIVO MATEMÁTICA
    # ═══════════════════════════════════════════════════════════════════════════
    html += f"<h2 id='sec{sec}'>Sección {sec}. Análisis Exhaustivo — Matemática</h2>"; sec += 1
    html += "<div class='subject-header sh-m'>Prueba de Fundamentos Matemáticas · Ciclo 1-2. Prueba Pre y Post</div>"

    first_m = True
    for bloque in MATEMATICA_JERARQUIA:
        bcol   = bloque['bloque_col']
        btitle = bloque['titulo']
        grupos = bloque['grupos']

        pcol        = f'{bcol}_puntaje_0_100'
        prom_bloque = df_scored[pcol].mean() if pcol in df_scored.columns else 0
        n_items_bloque = sum(len(g['items']) for g in grupos)

        html += f"<div class='bloque-card'>"
        html += f"<div class='bloque-title'>{btitle}</div>"
        html += f"<div class='bloque-body'>"
        html += (f"<div class='summary-row'>"
                 f"<span class='stat-pill'>Promedio del bloque: <span>{prom_bloque:.1f} / 100</span></span>"
                 f"<span class='stat-pill'>Ítems: <span>{n_items_bloque}</span></span>"
                 f"</div>")

        # ── box plot distribución del bloque (ancho completo) ────────────────
        if pcol in df_scored.columns:
            bpm = _bloque_box(bcol, btitle, COLORS_M[0], include_plotlyjs=first_m)
            first_m = False
            html += f"<div class='container'>{bpm}</div>"

        # ── % acierto por ítem — bloque completo (ancho completo) ────────────
        all_items_bloque = [it for g in grupos for it in g['items'] if it in df_scored.columns]
        if all_items_bloque:
            vals_b = [df_scored[it].mean() * 100 for it in all_items_bloque]
            labs_b = [ITEM_LABELS.get(it, it) for it in all_items_bloque]
            html += (f"<div class='container'>"
                     f"{_bar_items(all_items_bloque, labs_b, vals_b, '% de Acierto por Ítem — Bloque completo', COLORS_M[0])}"
                     f"</div>")

        for grupo in grupos:
            glabel = grupo['label']
            gitems = [it for it in grupo['items'] if it in df_scored.columns]
            if not gitems:
                continue

            prom_g = sum(df_scored[it].mean() for it in gitems) / len(gitems) * 100
            html += f"<div class='grupo-label'>{glabel}</div>"
            html += (f"<div class='summary-row'>"
                     f"<span class='stat-pill'>Promedio del grupo: <span>{prom_g:.1f}%</span></span>"
                     f"<span class='stat-pill'>Ítems: <span>{len(gitems)}</span></span>"
                     f"</div>")

            g_vals = [df_scored[it].mean() * 100 for it in gitems]
            g_labs = [ITEM_LABELS.get(it, it) for it in gitems]
            html += (f"<div class='container'>"
                     f"{_bar_items(gitems, g_labs, g_vals, f'% de Acierto — {glabel}', COLORS_M[1])}"
                     f"</div>")

        html += "</div></div>"  # bloque-body + bloque-card

    # ═══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 4 — CORRELACIÓN ENTRE BLOQUES
    # ═══════════════════════════════════════════════════════════════════════════
    html += f"<h2 id='sec{sec}'>Sección {sec}. Correlación entre Bloques</h2>"; sec += 1

    cols_l_b = [c for c in df_scored.columns if c.startswith('L_Bloque') and c.endswith('_puntaje_0_100')]
    if len(cols_l_b) > 1:
        corr_l = df_scored[cols_l_b].corr()
        etq_l  = [c.replace('L_Bloque', '').replace('_puntaje_0_100', '')
                   .lstrip('0123456789b_').replace('_', ' ').strip()
                  for c in corr_l.columns]
        corr_l.columns, corr_l.index = etq_l, etq_l
        fig_cl = px.imshow(corr_l, text_auto='.2f', aspect='auto',
                           color_continuous_scale=[C_SKY_BLUE, C_OCEAN_BLUE, C_DEEP_NAVY],
                           title='Correlación entre Bloques — Lengua')
        bottom_m = max(max(len(e) for e in etq_l) * 6, 120)
        fig_cl.update_layout(
            height=520,
            margin=dict(l=_lm(etq_l, minimum=200), r=40, t=60, b=bottom_m),
            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
            xaxis=dict(tickangle=-40, tickfont=dict(size=11)),
            yaxis=dict(tickfont=dict(size=11))
        )
        html += f"<div class='container'>{fig_cl.to_html(full_html=False, include_plotlyjs=False)}</div>"

    cols_m_b = [c for c in df_scored.columns if c.startswith('M_Bloque') and c.endswith('_puntaje_0_100')]
    if len(cols_m_b) > 1:
        corr_m = df_scored[cols_m_b].corr()
        etq_m  = [c.replace('M_Bloque', '').replace('_puntaje_0_100', '')
                   .lstrip('0123456789_').replace('_', ' ').strip()
                  for c in corr_m.columns]
        corr_m.columns, corr_m.index = etq_m, etq_m
        fig_cm = px.imshow(corr_m, text_auto='.2f', aspect='auto',
                           color_continuous_scale=[C_SKY_BLUE, COLORS_M[0], COLORS_M[2]],
                           title='Correlación entre Bloques — Matemática')
        bottom_m2 = max(max(len(e) for e in etq_m) * 6, 100)
        fig_cm.update_layout(
            height=420,
            margin=dict(l=_lm(etq_m, minimum=200), r=40, t=60, b=bottom_m2),
            font_color=C_DEEP_NAVY, title_font_color=C_DEEP_NAVY,
            xaxis=dict(tickangle=-40, tickfont=dict(size=11)),
            yaxis=dict(tickfont=dict(size=11))
        )
        html += f"<div class='container'>{fig_cm.to_html(full_html=False, include_plotlyjs=False)}</div>"

    # ═══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 5 — TABLA DE ESTUDIANTES
    # ═══════════════════════════════════════════════════════════════════════════
    html += f"""
<h2 id='sec{sec}'>Sección {sec}. Base de Datos de Estudiantes</h2>
<div class='container' style='overflow-x:auto;'>
  <p style='color:{C_STEEL_GREY};font-size:13px;margin-bottom:12px;'>
    <i>Usa la caja de búsqueda para filtrar por NIE o Nombre de Centro. Haz clic en los encabezados para ordenar.</i>
  </p>
  {tabla_html}
</div>
<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.js"></script>
<script>
$(document).ready(function(){{
  $('#tablaEstudiantes').DataTable({{
    language:{{url:'//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json'}},
    pageLength:10, scrollX:true
  }});
}});
</script>
</body></html>"""

    with open(ruta_salida, 'w', encoding='utf-8') as f:
        f.write(html)


# ==========================================
# 6. NÚCLEO ETL: CARGA, CRUCE Y CALIFICACIÓN
#    Función interna compartida por los tres flujos públicos.
# ==========================================
def _ejecutar_etl(current_month_path):
    """
    Carga el CSV de Google Forms, cruza con la matrícula oficial,
    califica dicotómicamente y calcula puntajes porcentuales.
    Devuelve (df_scored, df_mapping, paths) o (None, None, None) si falla.
    """
    PATH_FORMULARIOS = os.path.join(current_month_path, "01_Formularios_Crudos")
    PATH_DATASETS    = os.path.join(current_month_path, "02_Datasets_Procesados")
    os.makedirs(PATH_FORMULARIOS, exist_ok=True)
    os.makedirs(PATH_DATASETS, exist_ok=True)

    archivos_form = glob.glob(os.path.join(PATH_FORMULARIOS, "*.csv"))
    if not archivos_form:
        print(f"[!] ERROR: No hay archivos CSV en:\n    {PATH_FORMULARIOS}")
        return None, None, None

    archivo_reciente_form = max(archivos_form, key=os.path.getmtime)
    print(f"Leyendo formulario: {os.path.basename(archivo_reciente_form)}")
    df_form = pd.read_csv(archivo_reciente_form, dtype=str, encoding_errors='ignore')

    df_form = df_form[
        df_form['Dirección de correo electrónico'].str.endswith('@clases.edu.sv', na=False)
    ].copy()
    df_form['NIE'] = df_form['Dirección de correo electrónico'].apply(
        lambda x: x.split('@')[0].strip()
    )
    df_form = df_form.dropna(subset=['NIE'])

    col_puntuacion = next(
        (c for c in df_form.columns if 'puntuaci' in c.lower() or 'score' in c.lower()), None
    )
    df_form['Puntaje_Global_Formulario'] = (
        df_form[col_puntuacion].astype(str).str.split('/').str[0].str.strip()
        if col_puntuacion else "0"
    )

    print(f"Registros válidos: {len(df_form)}")

    columnas_items = df_form.columns[3:89]
    mapa = {
        col: (f'L_item_{i+1}' if i < 68 else f'M_item_{i - 68 + 1}')
        for i, col in enumerate(columnas_items)
    }
    df_form = df_form.rename(columns=mapa)

    archivos_mat = glob.glob(os.path.join(PATH_METADATA, "MatriculaProgresoMes*.csv"))
    if not archivos_mat:
        print("[!] ERROR: No se encontró ningún archivo de matrícula.")
        return None, None, None

    archivo_mat_reciente = max(archivos_mat, key=os.path.getmtime)
    df_mat = pd.read_csv(archivo_mat_reciente, sep=';', dtype=str, encoding_errors='ignore')

    col_nie   = next((c for c in df_mat.columns if 'nie' in c.lower()), None)
    col_cod   = next((c for c in df_mat.columns if 'codigo' in c.lower() or 'código' in c.lower()), None)
    col_nom   = next((c for c in df_mat.columns if 'nombre' in c.lower() and 'secc' not in c.lower()), None)
    col_grado = next((c for c in df_mat.columns if 'grado' in c.lower()), None)
    col_grupo = next((c for c in df_mat.columns if 'nombre_secc' in c.lower()), None)

    df_meta = df_mat[[col_nie, col_cod, col_nom, col_grado, col_grupo]].copy()
    df_meta.columns = ['NIE', 'Código de infraestructura', 'Nombre del centro', 'Grado', 'Grupo']
    df_meta['NIE'] = df_meta['NIE'].astype(str).str.replace('.0', '', regex=False).str.strip()

    df_cruzado = pd.merge(df_meta, df_form, on='NIE', how='right')
    for col in ['Código de infraestructura', 'Nombre del centro', 'Grado', 'Grupo']:
        df_cruzado[col] = df_cruzado[col].fillna("SIN REGISTRO")

    todas_preguntas = (
        [f'L_item_{i}' for i in range(1, 69)] +
        [f'M_item_{i}' for i in range(1, 19)]
    )
    columnas_orden = (
        ['Código de infraestructura', 'Nombre del centro', 'Grado', 'Grupo',
         'NIE', 'Dirección de correo electrónico', 'Puntaje_Global_Formulario']
        + todas_preguntas
    )

    df_raw    = df_cruzado[columnas_orden].copy()
    df_scored = df_raw.copy()

    for item in todas_preguntas:
        if item in CLAVE_RESPUESTAS:
            correcta = str(CLAVE_RESPUESTAS[item]).strip().lower()
            df_scored[item] = np.where(
                df_scored[item].astype(str).str.strip().str.lower() == correcta, 1, 0
            )
        else:
            df_scored[item] = 0

    df_scored = df_scored.copy()

    for seccion, items in TODAS_SECCIONES.items():
        items_exist = [it for it in items if it in df_scored.columns]
        if items_exist:
            n_items = len(items_exist)
            df_scored[f'{seccion}_suma_correcta'] = df_scored[items_exist].sum(axis=1)
            df_scored[f'{seccion}_puntaje_0_100'] = calcular_puntaje_porcentual(
                df_scored[f'{seccion}_suma_correcta'], n_items
            )
        else:
            df_scored[f'{seccion}_suma_correcta'] = 0
            df_scored[f'{seccion}_puntaje_0_100'] = 0.0

    df_scored = df_scored.copy()

    items_L = [c for c in df_scored.columns if c.startswith('L_item_')]
    items_M = [c for c in df_scored.columns if c.startswith('M_item_')]

    df_scored['L_total_suma_correcta'] = df_scored[items_L].sum(axis=1) if items_L else 0
    df_scored['M_total_suma_correcta'] = df_scored[items_M].sum(axis=1) if items_M else 0
    df_scored['L_total_puntaje_0_100'] = calcular_puntaje_porcentual(
        df_scored['L_total_suma_correcta'], len(items_L) if items_L else 1
    )
    df_scored['M_total_puntaje_0_100'] = calcular_puntaje_porcentual(
        df_scored['M_total_suma_correcta'], len(items_M) if items_M else 1
    )

    df_mapping = pd.DataFrame(
        [{'Item': k, 'Proceso': v} for k, v in TODO_PROCESS_MAP.items()]
    )

    paths = {
        'raw':     os.path.join(PATH_DATASETS, "1_Resultados_Fundamento_Respuestas_Crudas.xlsx"),
        'scored':  os.path.join(PATH_DATASETS, "2_Resultados_Fundamento_Dicotomico_0_1.xlsx"),
        'mapping': os.path.join(PATH_DATASETS, "3_Item_Process_Mapping.xlsx"),
        'html':    os.path.join(PATH_DATASETS, "4_Reporte_Grafico_Fundamentos.html"),
        'raw_df':  df_raw
    }

    return df_scored, df_mapping, paths


# ==========================================
# 7. FUNCIONES PÚBLICAS
# ==========================================

def procesar_datasets(current_month_path):
    """
    Opción (1): Solo genera los tres archivos Excel.
    """
    print("\n=======================================================")
    print(f"  PROCESANDO DATASETS: {os.path.basename(current_month_path)}")
    print("=======================================================\n")

    df_scored, df_mapping, paths = _ejecutar_etl(current_month_path)
    if df_scored is None:
        return False

    paths['raw_df'].to_excel(paths['raw'], index=False)
    df_scored.to_excel(paths['scored'], index=False)
    df_mapping.to_excel(paths['mapping'], index=False)

    print(f"\n[OK] Datasets generados en: {os.path.dirname(paths['raw'])}")
    print("  - 1_Resultados_Fundamento_Respuestas_Crudas.xlsx")
    print("  - 2_Resultados_Fundamento_Dicotomico_0_1.xlsx")
    print("  - 3_Item_Process_Mapping.xlsx")
    return True


def generar_reporte_html(current_month_path):
    """
    Opción (2): Solo genera el Dashboard HTML.
    Requiere que los Datasets ya existan en 02_Datasets_Procesados.
    Si no existen, ejecuta el ETL igualmente para poder generar el HTML.
    """
    print("\n=======================================================")
    print(f"  GENERANDO REPORTE HTML: {os.path.basename(current_month_path)}")
    print("=======================================================\n")

    PATH_DATASETS = os.path.join(current_month_path, "02_Datasets_Procesados")
    scored_path   = os.path.join(PATH_DATASETS, "2_Resultados_Fundamento_Dicotomico_0_1.xlsx")
    mapping_path  = os.path.join(PATH_DATASETS, "3_Item_Process_Mapping.xlsx")
    html_path     = os.path.join(PATH_DATASETS, "4_Reporte_Grafico_Fundamentos.html")

    # Si los datasets ya existen los reutiliza; si no, corre el ETL
    if os.path.exists(scored_path) and os.path.exists(mapping_path):
        print("[*] Datasets encontrados. Cargando desde disco...")
        df_scored  = pd.read_excel(scored_path, dtype=str)
        df_mapping = pd.read_excel(mapping_path)
        # Reconvertir columnas numéricas
        for col in df_scored.columns:
            if 'puntaje' in col or 'suma' in col or col.startswith(('L_item_', 'M_item_')):
                df_scored[col] = pd.to_numeric(df_scored[col], errors='coerce').fillna(0)
    else:
        print("[!] Datasets no encontrados. Ejecutando ETL primero...")
        df_scored, df_mapping, paths = _ejecutar_etl(current_month_path)
        if df_scored is None:
            return False
        html_path = paths['html']

    print("\n[*] Generando Dashboard HTML Interactivo...")
    generar_html_estatico(df_scored, df_mapping, html_path)

    print(f"\n[OK] Reporte HTML generado en:\n  {html_path}")
    return True


def procesar_y_generar_excel(current_month_path):
    """
    Opción (3): Genera los tres Datasets Excel y el Dashboard HTML.
    """
    print("\n=======================================================")
    print(f"  PROCESANDO: {os.path.basename(current_month_path)}")
    print("=======================================================\n")

    df_scored, df_mapping, paths = _ejecutar_etl(current_month_path)
    if df_scored is None:
        return False

    paths['raw_df'].to_excel(paths['raw'], index=False)
    df_scored.to_excel(paths['scored'], index=False)
    df_mapping.to_excel(paths['mapping'], index=False)

    print("\n[*] Generando Dashboard HTML Interactivo...")
    generar_html_estatico(df_scored, df_mapping, paths['html'])

    print(f"\n[OK] Archivos generados en: {os.path.dirname(paths['raw'])}")
    print("  - 1_Resultados_Fundamento_Respuestas_Crudas.xlsx")
    print("  - 2_Resultados_Fundamento_Dicotomico_0_1.xlsx")
    print("  - 3_Item_Process_Mapping.xlsx")
    print("  - 4_Reporte_Grafico_Fundamentos.html")
    return True