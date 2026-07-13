# ============================================================
# PAARS — Script 2: Equiparación entre Grados · Prueba de Fundamentos
#
# PRE-REQUISITO: haber ejecutado irt_fundamentos.R (Script 1),
# que genera:
#   results/models/{LEC|MAT}-{2do|3er|4to}.rds   ← modelos IRT
#   results/{LEC|MAT}-resultados.xlsx             ← puntajes por estudiante
#
# QUÉ HACE ESTE SCRIPT:
#   La misma prueba fue aplicada a 2do, 3er y 4to grado, pero cada
#   grado tiene su propio modelo 2PL (escalas con diferente origen
#   y unidad). Para comparar el desempeño entre grados se necesita
#   transformar todas las escalas a una escala común.
#
#   Estrategia: equiparación directa Stocking-Lord hacia 2do Grado.
#     - 3er → 2do (independiente)
#     - 4to → 2do (independiente)
#     - 2do → 2do (identidad; es la escala de referencia)
#
#   Equiparación directa (no encadenada) para no acumular el error
#   de dos transformaciones en el grado de 4to.
#
#   Ítems ancla: intersección de ítems válidos de cada par de modelos.
#   Como la prueba es la misma, el ancla es casi siempre el conjunto
#   completo de ítems no removidos por baja calidad en alguno de los
#   dos grados del par.
#
#   POST-EQUIPARACIÓN: Re-escalado a escala 0-100 sobre la
#   distribución POBLACIONAL COMBINADA de los tres grados
#   (media=50, SD=17, truncado a [0, 100]).
#   Esto garantiza que el puntaje final sea comparable entre grados
#   y entre escuelas que tienen estudiantes de distintos grados.
#
#   DIFERENCIA CLAVE respecto a irt_fundamentos.R (Script 1):
#   Script 1 aplica el re-escalado 0-100 usando la media y SD de
#   CADA GRADO por separado, lo que hace que un puntaje de 60 en
#   2do y un puntaje de 60 en 3er no sean comparables entre sí.
#   Este script usa la media y SD de la distribución COMBINADA de
#   los tres grados, produciendo una escala 0-100 uniforme y
#   comparable para todos los estudiantes independientemente del
#   grado al que pertenecen.
#
# SALIDA:
#   results/{LEC|MAT}-equalized_scores.xlsx
#     · Una hoja por grado con los puntajes originales MÁS las
#       columnas:
#         equalized.theta.global                ← theta equiparado (escala IRT)
#         equalized.theta.global (escala 0-100) ← legado; NO usar para comparación
#         score.final (escala 0-100)            ← COLUMNA RECOMENDADA
#     · Hoja "linking_diagnostics" con slope, intercept y n_anchor
#       de cada equiparación (trazabilidad/auditoría).
#     · Hoja "pooled_scale_params" con los parámetros del re-escalado
#       poblacional (auditoría/reproducibilidad).
#
# LÓGICA REPLICADA DE: 0-equateIRT.R
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(openxlsx)
  library(mirt)
  library(equateIRT)
  library(dplyr)
  library(plyr)
})

# ============================================================
# CONFIGURACIÓN
# ============================================================

# 1. Heredar la ruta maestra directamente desde la memoria de Python
base_drive <- Sys.getenv("BASE_DRIVE")
if (base_drive == "") {
  stop("[!] ERROR: R no recibió la variable BASE_DRIVE desde Python.")
}

# 2. Leer los argumentos enviados por Python (La carpeta del mes seleccionado)
args <- commandArgs(trailingOnly = TRUE)
if (length(args) == 0) {
  stop("[!] ERROR: R no recibió la ruta de la carpeta del mes desde Python.")
}
carpeta_mes <- args[1]

# 3. Construir rutas dinámicas combinando el argumento del mes
CARPETA_RESULTADOS <- file.path(carpeta_mes, "results")
CARPETA_MODELOS    <- file.path(CARPETA_RESULTADOS, "models")

SUBJECT_CODES <- c("LEC", "MAT")

SHEET_NAMES <- c('Segundo Grado' = '2do',
                 'Tercer Grado'  = '3er',
                 'Cuarto Grado'  = '4to')

# Grado de referencia: todos los demás se transforman a esta escala.
GRADE_REFERENCE <- '2do'

# Columnas de theta producidas por irt_fundamentos.R que serán equiparadas.
SCORE_COLUMNS <- c(
  "theta.global",
  "theta.global (escala 0-100)"
)

# Parámetros de la escala 0-100 final (mismos que usa Script 1,
# pero aplicados sobre la distribución POBLACIONAL COMBINADA).
SCALE_MEAN <- 50
SCALE_SD   <- 17
SCALE_MIN  <- 0
SCALE_MAX  <- 100

# Métodos calculados; el primero es el aplicado.
APPLIED_LINKING_METHOD    <- "Stocking-Lord"
LINKING_METHODS_TO_REPORT <- c("Stocking-Lord", "Haebara")

# Color de cabecera Excel por asignatura
BG_COLOR <- c(LEC = "#033b6d", MAT = "#085041")

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

build_model_path <- function(subject_code, grade_code) {
  file.path(CARPETA_MODELOS,
            paste0(subject_code, "-", grade_code, ".rds"))
}

build_results_path <- function(subject_code) {
  file.path(CARPETA_RESULTADOS,
            paste0(subject_code, "-resultados.xlsx"))
}

validate_required_columns <- function(data_table, required_columns, table_name) {
  missing <- setdiff(required_columns, names(data_table))
  if (length(missing) > 0) {
    stop(
      paste0("Faltan columnas en ", table_name, ": ",
             paste(missing, collapse = ", ")),
      call. = FALSE
    )
  }
}

read_irt_model <- function(subject_code, grade_code) {
  path <- build_model_path(subject_code, grade_code)
  if (!file.exists(path)) {
    stop(sprintf(
      "Modelo no encontrado: %s\n  Ejecute primero irt_fundamentos.R.",
      path
    ), call. = FALSE)
  }
  readRDS(path)
}

read_grade_results <- function(subject_code, grade_code) {
  path <- build_results_path(subject_code)
  if (!file.exists(path)) {
    stop(sprintf(
      "Archivo de resultados no encontrado: %s\n  Ejecute primero irt_fundamentos.R.",
      path
    ), call. = FALSE)
  }
  as.data.frame(read_excel(path, sheet = grade_code))
}

get_anchor_items <- function(model_source, model_ref,
                             label_source = "source",
                             label_ref    = "referencia") {
  anchor <- intersect(
    colnames(model_source@Data$data),
    colnames(model_ref@Data$data)
  )
  if (length(anchor) == 0) {
    stop(sprintf(
      "No hay ítems ancla entre %s y %s. Verifique los modelos.",
      label_source, label_ref
    ), call. = FALSE)
  }
  message(sprintf("    Ítems ancla  %s → %s: %d",
                  label_source, label_ref, length(anchor)))
  anchor
}

estimate_linking_coefficients <- function(
    model_source,
    model_ref,
    anchor_items,
    applied_method    = APPLIED_LINKING_METHOD,
    methods_to_report = LINKING_METHODS_TO_REPORT
) {
  linked_models <- equateIRT::modIRT(
    est.mods = list(model_source, model_ref),
    names    = c("source", "ref"),
    display  = FALSE
  )
  
  invisible(equateIRT::linkp(linked_models))
  
  linking_results <- lapply(methods_to_report, function(method_name) {
    eq <- equateIRT::direc(
      mods         = linked_models,
      which        = c("source", "ref"),
      method       = method_name,
      items.select = anchor_items
    )
    message(sprintf("\n    Resumen equiparación — %s:", method_name))
    print(summary(eq))
    eq
  })
  names(linking_results) <- methods_to_report
  
  if (!applied_method %in% names(linking_results)) {
    stop(sprintf(
      "Método '%s' no calculado. Disponibles: %s",
      applied_method, paste(names(linking_results), collapse = ", ")
    ), call. = FALSE)
  }
  
  coef_tbl  = as.data.frame(summary(linking_results[[applied_method]])$coefficients)
  slope     = as.numeric(coef_tbl$Estimate[1])
  intercept = as.numeric(coef_tbl$Estimate[2])
  
  list(
    slope        = slope,
    intercept    = intercept,
    all_results  = linking_results,
    anchor_items = anchor_items
  )
}

add_equalized_scores <- function(results_df, score_columns, slope, intercept) {
  validate_required_columns(results_df, score_columns, "tabla de resultados")
  for (sc in score_columns) {
    results_df[[paste0("equalized.", sc)]] <-
      slope * results_df[[sc]] + intercept
  }
  results_df
}

build_linking_diagnostics <- function(linking_log) {
  plyr::rbind.fill(lapply(linking_log, function(e) {
    data.frame(
      subject        = e$subject,
      grade_source   = e$grade_source,
      grade_ref      = e$grade_ref,
      method         = e$method,
      slope          = e$slope,
      intercept      = e$intercept,
      n_anchor_items = length(e$anchor_items),
      anchor_items   = paste(e$anchor_items, collapse = "; "),
      stringsAsFactors = FALSE
    )
  }))
}

# ============================================================
# FUNCIÓN PRINCIPAL: equate_subject
# ============================================================
equate_subject <- function(subject_code,
                           grade_codes   = c("2do", "3er", "4to"),
                           grade_ref     = GRADE_REFERENCE,
                           score_columns = SCORE_COLUMNS) {
  
  cat(sprintf("\n%s\n  Equiparando: %s\n%s\n",
              paste(rep("-", 50), collapse = ""),
              subject_code,
              paste(rep("-", 50), collapse = "")))
  
  # ── Leer modelos y resultados ──────────────────────────────────────
  models  <- list()
  results <- list()
  
  for (gc in grade_codes) {
    model_exists   <- file.exists(build_model_path(subject_code, gc))
    results_exists <- file.exists(build_results_path(subject_code))
    
    if (!model_exists || !results_exists) {
      message(sprintf("  [!] Saltando %s — %s: archivos no encontrados.",
                      subject_code, gc))
      next
    }
    
    cat(sprintf("  Leyendo modelo y resultados: %s — %s\n",
                subject_code, gc))
    models[[gc]]  <- read_irt_model(subject_code, gc)
    results[[gc]] <- read_grade_results(subject_code, gc)
  }
  
  if (!grade_ref %in% names(models)) {
    stop(sprintf(
      "Grado de referencia '%s' no disponible para %s.",
      grade_ref, subject_code
    ), call. = FALSE)
  }
  
  # ── Equiparación por grado ─────────────────────────────────────────
  results_equated <- list()
  linking_log     <- list()
  
  for (gc in names(models)) {
    
    if (gc == grade_ref) {
      cat(sprintf("\n  %s [referencia] → slope=1, intercept=0\n", gc))
      results_equated[[gc]] <- add_equalized_scores(
        results_df    = results[[gc]],
        score_columns = score_columns,
        slope         = 1,
        intercept     = 0
      )
      next
    }
    
    cat(sprintf("\n  Equiparando %s → %s\n", gc, grade_ref))
    
    anchor <- get_anchor_items(
      model_source = models[[gc]],
      model_ref    = models[[grade_ref]],
      label_source = gc,
      label_ref    = grade_ref
    )
    
    linking <- estimate_linking_coefficients(
      model_source      = models[[gc]],
      model_ref         = models[[grade_ref]],
      anchor_items      = anchor,
      applied_method    = APPLIED_LINKING_METHOD,
      methods_to_report = LINKING_METHODS_TO_REPORT
    )
    
    cat(sprintf(
      "  Coeficientes (%s): slope = %.4f | intercept = %.4f\n",
      APPLIED_LINKING_METHOD, linking$slope, linking$intercept
    ))
    
    results_equated[[gc]] <- add_equalized_scores(
      results_df    = results[[gc]],
      score_columns = score_columns,
      slope         = linking$slope,
      intercept     = linking$intercept
    )
    
    linking_log[[paste(subject_code, gc, sep = "_")]] <- list(
      subject      = subject_code,
      grade_source = gc,
      grade_ref    = grade_ref,
      method       = APPLIED_LINKING_METHOD,
      slope        = linking$slope,
      intercept    = linking$intercept,
      anchor_items = anchor
    )
  }
  
  # ── Exportar {subject}-equalized_scores.xlsx ──────────────────────
  ruta_eq <- file.path(CARPETA_RESULTADOS,
                       paste0(subject_code, "-equalized_scores.xlsx"))
  wb <- openxlsx::createWorkbook()
  hs <- openxlsx::createStyle(
    fontColour = "#FFFFFF", fgFill = BG_COLOR[[subject_code]],
    halign = "center", textDecoration = "bold", wrapText = TRUE)
  
  for (gc in names(results_equated)) {
    openxlsx::addWorksheet(wb, gc)
    openxlsx::writeData(wb, gc, results_equated[[gc]], headerStyle = hs)
    openxlsx::freezePane(wb, gc, firstRow = TRUE)
  }
  
  if (length(linking_log) > 0) {
    diag_tbl <- build_linking_diagnostics(linking_log)
    openxlsx::addWorksheet(wb, "linking_diagnostics")
    openxlsx::writeData(wb, "linking_diagnostics", diag_tbl, headerStyle = hs)
  }
  
  openxlsx::saveWorkbook(wb, ruta_eq, overwrite = TRUE)
  cat(sprintf("\n    ✓ %s\n", ruta_eq))
  
  invisible(list(results_equated = results_equated,
                 linking_log     = linking_log))
}

# ============================================================
# MAIN — Paso 1: Equiparación Stocking-Lord
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  PAARS — Script 2: Equiparación entre Grados\n")
cat("          Prueba de Fundamentos\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")
cat("  Referencia : Segundo Grado (2do)\n")
cat("  Método     :", APPLIED_LINKING_METHOD, "\n")
cat("  Diagnóstico:", paste(LINKING_METHODS_TO_REPORT, collapse = " + "), "\n\n")

for (subj in SUBJECT_CODES) {
  equate_subject(
    subject_code  = subj,
    grade_codes   = as.character(SHEET_NAMES),
    grade_ref     = GRADE_REFERENCE,
    score_columns = SCORE_COLUMNS
  )
}

# ============================================================
# MAIN — Paso 2: Re-escalado 0-100 sobre distribución pooled
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  Paso 2: Re-escalado 0-100 sobre distribución pooled\n")
cat("          (media=50, SD=17, truncado a [0, 100])\n")
cat(paste(rep("=", 60), collapse = ""), "\n")

for (subj in SUBJECT_CODES) {
  
  ruta_eq <- file.path(CARPETA_RESULTADOS,
                       paste0(subj, "-equalized_scores.xlsx"))
  
  # ── 2a. Leer equalized.theta.global de todos los grados ───────────
  grade_codes_present <- as.character(SHEET_NAMES)
  grade_data <- list()
  
  for (gc in grade_codes_present) {
    df <- tryCatch(
      as.data.frame(readxl::read_excel(ruta_eq, sheet = gc)),
      error = function(e) NULL
    )
    if (!is.null(df) && "equalized.theta.global" %in% names(df)) {
      grade_data[[gc]] <- df
    }
  }
  
  if (length(grade_data) == 0) {
    warning(sprintf(
      "  [!] No se encontraron hojas con equalized.theta.global para %s", subj))
    next
  }
  
  # ── 2b. Calcular media y SD de la distribución POOLED ─────────────
  all_theta <- unlist(lapply(grade_data, function(d) d[["equalized.theta.global"]]))
  all_theta <- all_theta[!is.na(all_theta)]
  
  pool_mean <- mean(all_theta)
  pool_sd   <- sd(all_theta)
  pool_n    <- length(all_theta)
  
  cat(sprintf("\n  %s — Distribución pooled (n=%d):\n", subj, pool_n))
  cat(sprintf("    media  = %.6f\n", pool_mean))
  cat(sprintf("    SD     = %.6f\n", pool_sd))
  cat(sprintf("    rango  = [%.4f, %.4f]\n",
              min(all_theta), max(all_theta)))
  
  # ── 2c. Agregar score.final a cada hoja y volver a guardar ────────
  wb <- openxlsx::loadWorkbook(ruta_eq)
  hs <- openxlsx::createStyle(
    fontColour = "#FFFFFF", fgFill = BG_COLOR[[subj]],
    halign = "center", textDecoration = "bold", wrapText = TRUE)
  
  for (gc in names(grade_data)) {
    df <- grade_data[[gc]]
    
    # Transformación: theta equiparado → escala 0-100 poblacional
    df[["score.final (escala 0-100)"]] <-
      round(
        pmin(SCALE_MAX, pmax(SCALE_MIN,
                             (df[["equalized.theta.global"]] - pool_mean) / pool_sd * SCALE_SD + SCALE_MEAN
        )),
        1
      )
    
    # Reescribir la hoja en el workbook existente
    openxlsx::removeWorksheet(wb, gc)
    openxlsx::addWorksheet(wb, gc)
    openxlsx::writeData(wb, gc, df, headerStyle = hs)
    openxlsx::freezePane(wb, gc, firstRow = TRUE)
    
    cat(sprintf(
      "    %s %s — score.final: media=%.1f  DE=%.1f  rango=[%.1f, %.1f]\n",
      subj, gc,
      mean(df[["score.final (escala 0-100)"]], na.rm = TRUE),
      sd(df[["score.final (escala 0-100)"]], na.rm = TRUE),
      min(df[["score.final (escala 0-100)"]], na.rm = TRUE),
      max(df[["score.final (escala 0-100)"]], na.rm = TRUE)
    ))
  }
  
  # ── 2d. Hoja de auditoría: parámetros del re-escalado ─────────────
  scale_params <- data.frame(
    subject               = subj,
    grade_reference       = GRADE_REFERENCE,
    linking_method        = APPLIED_LINKING_METHOD,
    scale_mean_target     = SCALE_MEAN,
    scale_sd_target       = SCALE_SD,
    scale_min             = SCALE_MIN,
    scale_max             = SCALE_MAX,
    pool_n                = pool_n,
    pool_mean_theta_eq    = round(pool_mean, 6),
    pool_sd_theta_eq      = round(pool_sd,   6),
    formula = paste0(
      "score.final = pmin(100, pmax(0, ",
      "(equalized.theta.global - ", round(pool_mean, 6), ") / ",
      round(pool_sd, 6), " * ", SCALE_SD, " + ", SCALE_MEAN, "))"
    ),
    interpretation = paste0(
      "50 = promedio poblacional combinado (2do+3er+4to). ",
      "Cada 17 puntos = 1 SD de la distribucion combinada. ",
      "Escala anclada a 2do Grado via Stocking-Lord. ",
      "Truncado a [0, 100]."
    ),
    stringsAsFactors = FALSE
  )
  
  existing_sheets <- openxlsx::getSheetNames(ruta_eq)
  if ("pooled_scale_params" %in% existing_sheets) {
    openxlsx::removeWorksheet(wb, "pooled_scale_params")
  }
  openxlsx::addWorksheet(wb, "pooled_scale_params")
  openxlsx::writeData(wb, "pooled_scale_params", scale_params, headerStyle = hs)
  
  openxlsx::saveWorkbook(wb, ruta_eq, overwrite = TRUE)
  cat(sprintf("  ✓ %s actualizado con score.final (escala 0-100)\n", ruta_eq))
}

# ============================================================
# RESUMEN FINAL
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  [✓] Script 2 terminado.\n\n")
cat("  Archivos generados en:", CARPETA_RESULTADOS, "\n\n")
cat("  Columna recomendada para reportes y comparacion de escuelas:\n")
cat("    score.final (escala 0-100)\n\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")