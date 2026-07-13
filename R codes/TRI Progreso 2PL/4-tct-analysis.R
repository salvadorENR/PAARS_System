# =============================================================================
# 4-tct-analysis.R
# Análisis de Teoría Clásica de los Tests (TCT)
#   1. RPbis corregido por opción de cada ítem
#   2. Análisis de distractores (frecuencia por grupo de habilidad: bajo/medio/alto)
# Prueba de Progreso - Mes 03 - Junio 2026
#
# Author: Geiser C Challco <geiser.challco@goes.gob.sv>
#
# Inputs:
#   - data/processed/IRT-{subject}-nominal.xlsx  (respuestas nominales por grado)
#   - data/raw/pruebas.xlsx                      (metadata: opción correcta por ítem)
#
# Outputs (en results/TCT/):
#   - TCT-LEC-Mes03-Junio.xlsx      (RPbis corregido por opción)
#   - TCT-MAT-Mes03-Junio.xlsx
#   - Distractores-LEC-Mes03-Junio.xlsx  (análisis de distractores por tercil)
#   - Distractores-MAT-Mes03-Junio.xlsx
# =============================================================================

required_packages <- c("readxl", "dplyr", "writexl")
missing_packages  <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readxl)
library(dplyr)
library(writexl)

# ---- Crear carpeta de salida ----
dir.create("./results/TCT", showWarnings = FALSE, recursive = TRUE)

# ---- Funciones auxiliares ----

# Normaliza letras de opción a mayúscula sin puntos.
# LEC usa "a.", "b.", "c.", "d." -> "A", "B", "C", "D"
# MAT usa "A ", "B ", etc.       -> "A", "B", "C", "D"
normalise_option <- function(x) {
  toupper(trimws(gsub("\\.", "", x)))
}

# =============================================================================
# FUNCIÓN 1: TCT con RPbis corregido (item-rest correlation)
# =============================================================================
compute_tct <- function(nominal_df, pruebas_meta) {

  items   <- colnames(nominal_df)
  options <- c("A", "B", "C", "D")

  # Respuesta correcta por ítem (normalizada)
  correct_answers <- setNames(
    normalise_option(pruebas_meta$OpcionCorrecta[match(items, pruebas_meta$ItemCodigo)]),
    items
  )

  # Matriz de aciertos (0/1) para todos los ítems — base del puntaje total
  scored_matrix <- as.matrix(mapply(function(col, correct) {
    as.integer(!is.na(col) & normalise_option(col) == correct)
  }, nominal_df, correct_answers))

  total_score <- rowSums(scored_matrix, na.rm = TRUE)

  rows <- lapply(items, function(item) {

    responses <- normalise_option(nominal_df[[item]])
    n_total   <- sum(!is.na(responses))
    correct   <- correct_answers[[item]]

    # Puntaje rest: total SIN el ítem actual (corrección part-whole)
    rest_score <- total_score - scored_matrix[, item]

    option_rows <- lapply(options, function(opt) {

      selected <- !is.na(responses) & responses == opt
      freq_abs <- sum(selected)

      x <- as.integer(selected)
      if (length(unique(x)) < 2 || sd(rest_score, na.rm = TRUE) == 0) {
        rpbis <- NA_real_
      } else {
        rpbis <- cor(x, rest_score, use = "complete.obs", method = "pearson")
      }

      data.frame(
        Item           = item,
        Opcion         = opt,
        Es_correcta    = ifelse(!is.na(correct) & opt == correct, "*", ""),
        Freq_Absoluta  = freq_abs,
        Freq_Relativa  = ifelse(n_total > 0, round(freq_abs / n_total, 4), NA_real_),
        rpbis          = round(rpbis, 4),
        stringsAsFactors = FALSE
      )
    })

    do.call(rbind, option_rows)
  })

  do.call(rbind, rows)
}

# =============================================================================
# FUNCIÓN 2: Análisis de distractores por grupo de habilidad (terciles)
#
# Para cada ítem y cada opción calcula la frecuencia relativa de elección
# en cada grupo de habilidad (bajo = tercil 1, medio = tercil 2, alto = tercil 3).
# Un buen ítem muestra que la opción correcta aumenta monotónicamente del
# grupo bajo al alto, y las opciones incorrectas (distractores) disminuyen.
# =============================================================================
compute_distractors <- function(nominal_df, pruebas_meta) {

  items   <- colnames(nominal_df)
  options <- c("A", "B", "C", "D")

  correct_answers <- setNames(
    normalise_option(pruebas_meta$OpcionCorrecta[match(items, pruebas_meta$ItemCodigo)]),
    items
  )

  # Puntaje total para clasificar estudiantes en terciles
  scored_matrix <- as.matrix(mapply(function(col, correct) {
    as.integer(!is.na(col) & normalise_option(col) == correct)
  }, nominal_df, correct_answers))

  total_score <- rowSums(scored_matrix, na.rm = TRUE)

  # Asignar grupo por tercil del puntaje total
  terciles <- quantile(total_score, probs = c(1/3, 2/3), na.rm = TRUE)
  grupo <- ifelse(total_score <= terciles[1], "Bajo",
           ifelse(total_score <= terciles[2], "Medio", "Alto"))
  grupo <- factor(grupo, levels = c("Bajo", "Medio", "Alto"))

  rows <- lapply(items, function(item) {

    responses <- normalise_option(nominal_df[[item]])
    correct   <- correct_answers[[item]]

    option_rows <- lapply(options, function(opt) {

      selected <- !is.na(responses) & responses == opt

      # Frecuencia general (todos los estudiantes sin distinción de grupo)
      n_total_resp <- sum(!is.na(responses))
      n_total      <- sum(selected, na.rm = TRUE)
      freq_total   <- ifelse(n_total_resp > 0, n_total / n_total_resp, NA_real_)

      # Frecuencia relativa de esta opción dentro de cada grupo de habilidad
      freq_bajo  <- sum(selected & grupo == "Bajo",  na.rm = TRUE) /
                    max(sum(!is.na(responses) & grupo == "Bajo"),  1)
      freq_medio <- sum(selected & grupo == "Medio", na.rm = TRUE) /
                    max(sum(!is.na(responses) & grupo == "Medio"), 1)
      freq_alto  <- sum(selected & grupo == "Alto",  na.rm = TRUE) /
                    max(sum(!is.na(responses) & grupo == "Alto"),  1)

      # N total por grupo (para contexto)
      n_bajo  <- sum(!is.na(responses) & grupo == "Bajo",  na.rm = TRUE)
      n_medio <- sum(!is.na(responses) & grupo == "Medio", na.rm = TRUE)
      n_alto  <- sum(!is.na(responses) & grupo == "Alto",  na.rm = TRUE)

      data.frame(
        Item        = item,
        Opcion      = opt,
        Es_correcta = ifelse(!is.na(correct) & opt == correct, "*", ""),
        N_Total     = n_total,
        Frec_Total  = round(freq_total, 4),
        N_Bajo      = n_bajo,
        Frec_Bajo   = round(freq_bajo,  4),
        N_Medio     = n_medio,
        Frec_Medio  = round(freq_medio, 4),
        N_Alto      = n_alto,
        Frec_Alto   = round(freq_alto,  4),
        Discrimina  = ifelse(!is.na(correct) & opt == correct,
                       ifelse(freq_alto - freq_bajo >= 0.40, "Excelente",
                       ifelse(freq_alto - freq_bajo >= 0.30, "Bueno",
                       ifelse(freq_alto - freq_bajo >= 0.20, "Aceptable",
                       ifelse(freq_alto - freq_bajo >= 0.10, "Debil",
                                                             "Deficiente")))), ""),
        stringsAsFactors = FALSE
      )
    })

    do.call(rbind, option_rows)
  })

  do.call(rbind, rows)
}

# ---- Lectura de metadatos (opción correcta por ítem) ----

pruebas_raw <- read_excel("./data/raw/pruebas.xlsx")
colnames(pruebas_raw) <- trimws(colnames(pruebas_raw))

pruebas_meta <- data.frame(
  ItemCodigo     = trimws(pruebas_raw$ItemCodigo),
  OpcionCorrecta = trimws(pruebas_raw$`Opcion Correcta`),
  stringsAsFactors = FALSE
)
pruebas_meta <- unique(pruebas_meta)

# ---- Procesamiento por subject y grado ----

grades <- c("2do", "3er", "4to", "5to", "6to", "7mo", "8vo", "9no", "Bach-1er", "Bach-2do")

lapply(c("LEC", "MAT"), function(subject) {

  message("\n=== Procesando TCT y Distractores: ", subject, " ===")

  tct_by_grade  <- list()
  dist_by_grade <- list()

  for (grade in grades) {

    message("  Grado: ", grade)

    nominal_df <- tryCatch(
      read_excel(paste0("./data/processed/IRT-", subject, "-nominal.xlsx"), sheet = grade),
      error = function(e) {
        warning("No se encontro hoja ", grade, " en nominal de ", subject, ": ", e$message)
        return(NULL)
      }
    )

    if (is.null(nominal_df)) next

    tct_by_grade[[grade]]  <- compute_tct(nominal_df, pruebas_meta)
    dist_by_grade[[grade]] <- compute_distractors(nominal_df, pruebas_meta)
  }

  # --- Guardar TCT (RPbis) ---
  tct_path <- paste0("./results/TCT/TCT-", subject, "-Mes03-Junio.xlsx")
  writexl::write_xlsx(tct_by_grade, tct_path)
  message("  Guardado: ", tct_path)

  # --- Guardar Distractores ---
  dist_path <- paste0("./results/TCT/Distractores-", subject, "-Mes03-Junio.xlsx")
  writexl::write_xlsx(dist_by_grade, dist_path)
  message("  Guardado: ", dist_path)
})

message("\n=== TCT y Distractores completados para LEC y MAT ===")
