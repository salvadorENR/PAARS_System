# ============================================================
# PAARS — Script 3: Reporte de Estudiantes con score.final
#
# PRE-REQUISITO: haber ejecutado equate_fundamentos.R (Script 2)
# actualizado, que genera la columna score.final (escala 0-100)
# en los archivos:
#    results/LEC-equalized_scores.xlsx
#    results/MAT-equalized_scores.xlsx
#
# QUÉ HACE ESTE SCRIPT:
#    Lee los archivos de puntajes equiparados de LEC y MAT,
#    extrae las columnas de identificación del estudiante y
#    score.final (escala 0-100), y produce un Excel con DOS hojas:
#
#    HOJA 1 — Estudiantes:
#      Código de infraestructura | Nombre del centro | Grado | NIE |
#      score.final LEC | Categoría LEC | score.final MAT | Categoría MAT
#      Las celdas de categoría tienen relleno de color según nivel:
#        Crítico   ≤ 35  → rojo     (#FF4B4B), texto blanco
#        Bajo    36-45   → naranja  (#FF9966), texto blanco
#        Medio   46-55   → amarillo (#FFD966), texto negro
#        Bueno   56-65   → verde claro (#A9D18E), texto negro
#        Excelente > 65  → verde oscuro (#1E7145), texto blanco
#
#    HOJA 2 — Escuelas:
#      Código | Nombre | N estudiantes | % Crítico LEC | % Crítico MAT |
#      Gana LEC | Gana MAT | Estatus Escuela
#      Clasificación del estatus:
#        Excelente (verde)  → gana LEC Y gana MAT   (1,1)
#        Regular   (naranja)→ gana solo uno          (1,0) o (0,1)
#        Alerta    (rojo)   → no gana ninguno        (0,0)
#      Una escuela "gana" una asignatura si su % de estudiantes
#      Crítico es <= al % Crítico del universo (todas las escuelas).
#
# SALIDA:
#    results/reporte_estudiantes.xlsx
#
# NOTAS:
#    - Si un NIE aparece duplicado dentro del mismo grado, se conserva
#      la última ocurrencia.
#    - score.final está anclada a la distribución poblacional combinada
#      de 2do, 3er y 4to Grado (media=50, SD=17). Comparable entre grados.
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(openxlsx)
  library(dplyr)
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

# 3. Construir la ruta de resultados dentro de la carpeta del mes
CARPETA_RESULTADOS <- file.path(carpeta_mes, "results")

ARCHIVOS <- list(
  LEC = file.path(CARPETA_RESULTADOS, "LEC-equalized_scores.xlsx"),
  MAT = file.path(CARPETA_RESULTADOS, "MAT-equalized_scores.xlsx")
)

HOJAS_GRADOS <- c("2do", "3er", "4to")

COLS_ID <- c(
  "NIE",
  "Código de infraestructura",
  "Nombre del centro",
  "Grado"
)

COL_SCORE <- "score.final (escala 0-100)"

ARCHIVO_SALIDA <- file.path(CARPETA_RESULTADOS, "reporte_estudiantes.xlsx")

ORDEN_GRADOS <- c(
  "Segundo Grado" = 1,
  "Tercer Grado"  = 2,
  "Cuarto Grado"  = 3
)

# ── Colores de categoría estudiante ───────────────────────────────────
NIVEL_COLORS <- list(
  "Crítico"   = list(fgFill = "FF4B4B", fontColour = "FFFFFF"),
  "Bajo"      = list(fgFill = "FF9966", fontColour = "FFFFFF"),
  "Medio"     = list(fgFill = "FFD966", fontColour = "000000"),
  "Bueno"     = list(fgFill = "A9D18E", fontColour = "000000"),
  "Excelente" = list(fgFill = "1E7145", fontColour = "FFFFFF")
)

# ── Colores de estatus escuela ────────────────────────────────────────
ESTATUS_COLORS <- list(
  "Excelente" = list(fgFill = "1E7145", fontColour = "FFFFFF"),
  "Regular"   = list(fgFill = "FF8C2E", fontColour = "FFFFFF"),
  "Alerta"    = list(fgFill = "FF4B4B", fontColour = "FFFFFF")
)

# ── Colores de cabecera ───────────────────────────────────────────────
BG_LEC   <- "#033B6D"
BG_MAT   <- "#085041"
BG_GEN   <- "#2E4057"
BG_2DO   <- "#EBF5FB"
BG_3ER   <- "#E8F8F5"
BG_4TO   <- "#FEF9E7"

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

#' Clasifica un puntaje 0-100 en categoría
clasificar_puntaje <- function(val) {
  if (is.na(val)) return("Crítico")
  val <- as.numeric(val)
  if      (val <= 35) return("Crítico")
  else if (val <= 45) return("Bajo")
  else if (val <= 55) return("Medio")
  else if (val <= 65) return("Bueno")
  else                return("Excelente")
}

#' Verifica archivos de entrada
verificar_archivos <- function() {
  errores <- character(0)
  for (subj in names(ARCHIVOS)) {
    ruta <- ARCHIVOS[[subj]]
    if (!file.exists(ruta)) {
      errores <- c(errores, sprintf(
        "  [!] Archivo no encontrado: %s\n      Ejecute primero equate_fundamentos.R.",
        ruta))
      next
    }
    for (hoja in HOJAS_GRADOS) {
      cols <- tryCatch(names(read_excel(ruta, sheet = hoja, n_max = 1)),
                       error = function(e) NULL)
      if (is.null(cols)) {
        errores <- c(errores, sprintf("  [!] No se pudo leer %s - %s.", subj, hoja))
        next
      }
      if (!COL_SCORE %in% trimws(gsub("\n", " ", cols))) {
        errores <- c(errores, sprintf(
          "  [!] Columna '%s' no encontrada en %s - %s.\n      Ejecute primero equate_fundamentos.R.",
          COL_SCORE, subj, hoja))
      }
    }
  }
  return(errores)
}

#' Lee puntajes de todos los grados de un archivo equiparado
leer_puntajes <- function(subj, ruta) {
  frames <- list()
  for (hoja in HOJAS_GRADOS) {
    df <- tryCatch(
      as.data.frame(read_excel(ruta, sheet = hoja, col_types = "guess")),
      error = function(e) { message(sprintf("  [!] No se pudo leer %s - %s", subj, hoja)); NULL }
    )
    if (is.null(df)) next
    
    # Limpiar nombres de columnas
    names(df) <- trimws(gsub("\n", " ", names(df)))
    
    cols_necesarias <- c(COLS_ID, COL_SCORE)
    faltantes <- setdiff(cols_necesarias, names(df))
    if (length(faltantes) > 0) {
      message(sprintf("  [!] Columnas faltantes en %s - %s: %s",
                      subj, hoja, paste(faltantes, collapse = ", ")))
      next
    }
    
    df <- df[, cols_necesarias, drop = FALSE]
    
    n_dupes <- sum(duplicated(df$NIE))
    if (n_dupes > 0) {
      message(sprintf("  [i] %s %s: %d NIE(s) duplicados — conservando última ocurrencia.", subj, hoja, n_dupes))
      df <- df[!duplicated(df$NIE, fromLast = TRUE), ]
    }
    frames[[hoja]] <- df
  }
  if (length(frames) == 0) return(data.frame())
  combined <- do.call(rbind, frames)
  rownames(combined) <- NULL
  names(combined)[names(combined) == COL_SCORE] <- paste0("score.final.", subj)
  return(combined)
}

#' Crea un estilo openxlsx para una categoría de nivel
estilo_nivel <- function(nivel) {
  col <- NIVEL_COLORS[[nivel]]
  openxlsx::createStyle(
    fontName     = "Arial", fontSize = 10,
    fontColour   = paste0("#", col$fontColour),
    fgFill       = paste0("#", col$fgFill),
    halign       = "center", valign = "center",
    textDecoration = "bold",
    border       = "TopBottomLeftRight",
    borderColour = "#CCCCCC"
  )
}

#' Crea un estilo openxlsx para un estatus de escuela
estilo_estatus <- function(estatus) {
  col <- ESTATUS_COLORS[[estatus]]
  openxlsx::createStyle(
    fontName     = "Arial", fontSize = 10,
    fontColour   = paste0("#", col$fontColour),
    fgFill       = paste0("#", col$fgFill),
    halign       = "center", valign = "center",
    textDecoration = "bold",
    border       = "TopBottomLeftRight",
    borderColour = "#CCCCCC"
  )
}

# ============================================================
# MAIN
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  PAARS — Script 3: Reporte de Estudiantes\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")

# ── Verificar archivos ────────────────────────────────────────────────
errores <- verificar_archivos()
if (length(errores) > 0) {
  cat("  ERROR: No se puede continuar.\n\n")
  cat(paste(errores, collapse = "\n"), "\n")
  stop("Verifique los mensajes anteriores.", call. = FALSE)
}

# ── Leer puntajes ─────────────────────────────────────────────────────
cat("  Leyendo puntajes LEC...\n")
df_lec <- leer_puntajes("LEC", ARCHIVOS$LEC)
cat(sprintf("    → %d estudiantes\n\n", nrow(df_lec)))

cat("  Leyendo puntajes MAT...\n")
df_mat <- leer_puntajes("MAT", ARCHIVOS$MAT)
cat(sprintf("    → %d estudiantes\n\n", nrow(df_mat)))

# ── Combinar por NIE ──────────────────────────────────────────────────
cat("  Combinando LEC y MAT por NIE...\n")
estudiantes <- merge(
  df_lec,
  df_mat[, c("NIE", "score.final.MAT"), drop = FALSE],
  by = "NIE", all = FALSE
)

estudiantes <- estudiantes[!is.na(estudiantes$Grado) & estudiantes$Grado != "", ]

estudiantes$orden_grado <- ORDEN_GRADOS[estudiantes$Grado]
estudiantes <- estudiantes[order(
  estudiantes$`Código de infraestructura`,
  estudiantes$orden_grado,
  estudiantes$NIE), ]
estudiantes$orden_grado <- NULL
rownames(estudiantes) <- NULL

estudiantes$score.final.LEC <- round(as.numeric(estudiantes$score.final.LEC), 1)
estudiantes$score.final.MAT <- round(as.numeric(estudiantes$score.final.MAT), 1)

estudiantes$cat.LEC <- sapply(estudiantes$score.final.LEC, clasificar_puntaje)
estudiantes$cat.MAT <- sapply(estudiantes$score.final.MAT, clasificar_puntaje)

cat(sprintf("    → %d estudiantes en total\n", nrow(estudiantes)))
cat(sprintf("    → %d escuelas\n",
            length(unique(estudiantes$`Código de infraestructura`))))

# ============================================================
# CLASIFICACIÓN DE ESCUELAS
# ============================================================
cat("  Clasificando escuelas...\n")

universo_pct_lec <- mean(estudiantes$cat.LEC == "Crítico", na.rm = TRUE)
universo_pct_mat <- mean(estudiantes$cat.MAT == "Crítico", na.rm = TRUE)

cat(sprintf("    %% Crítico universo LEC: %.1f%%\n", universo_pct_lec * 100))
cat(sprintf("    %% Crítico universo MAT: %.1f%%\n", universo_pct_mat * 100))

escuelas_list <- list()
codigos <- unique(estudiantes$`Código de infraestructura`)

for (cod in codigos) {
  sub <- estudiantes[estudiantes$`Código de infraestructura` == cod, ]
  nombre <- sub$`Nombre del centro`[1]
  n_stu  <- nrow(sub)
  
  pct_crit_lec <- mean(sub$cat.LEC == "Crítico", na.rm = TRUE)
  pct_crit_mat <- mean(sub$cat.MAT == "Crítico", na.rm = TRUE)
  
  gana_lec <- (!is.nan(pct_crit_lec)) && (pct_crit_lec <= universo_pct_lec)
  gana_mat <- (!is.nan(pct_crit_mat)) && (pct_crit_mat <= universo_pct_mat)
  
  estatus <- if (gana_lec && gana_mat) {
    "Excelente"
  } else if (gana_lec || gana_mat) {
    "Regular"
  } else {
    "Alerta"
  }
  
  escuelas_list[[length(escuelas_list) + 1]] <- data.frame(
    `Código de infraestructura` = cod,
    `Nombre del centro`         = nombre,
    `N Estudiantes`             = n_stu,
    `% Crítico LEC`             = round(pct_crit_lec * 100, 1),
    `% Crítico MAT`             = round(pct_crit_mat * 100, 1),
    `% Crítico universo LEC`    = round(universo_pct_lec * 100, 1),
    `% Crítico universo MAT`    = round(universo_pct_mat * 100, 1),
    `Gana LEC`                  = ifelse(gana_lec, "✔ Sí", "✘ No"),
    `Gana MAT`                  = ifelse(gana_mat, "✔ Sí", "✘ No"),
    `Estatus`                   = estatus,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}

escuelas <- do.call(rbind, escuelas_list)

orden_estatus <- c("Alerta" = 1, "Regular" = 2, "Excelente" = 3)
escuelas$ord_est <- orden_estatus[escuelas$Estatus]
escuelas <- escuelas[order(escuelas$ord_est, escuelas$`Nombre del centro`), ]
escuelas$ord_est <- NULL
rownames(escuelas) <- NULL

cat("  Distribución de estatus de escuelas:\n")
for (est in c("Alerta", "Regular", "Excelente")) {
  n_est <- sum(escuelas$Estatus == est)
  cat(sprintf("    %s: %d escuela(s)\n", est, n_est))
}

# ============================================================
# CONSTRUIR EXCEL
# ============================================================
cat(sprintf("\n  Generando %s...\n", ARCHIVO_SALIDA))

wb <- openxlsx::createWorkbook()

hs_gen <- openxlsx::createStyle(
  fontColour = "#FFFFFF", fgFill = BG_GEN,
  halign = "center", valign = "center",
  textDecoration = "bold", wrapText = TRUE,
  fontSize = 10, fontName = "Arial")

hs_lec <- openxlsx::createStyle(
  fontColour = "#FFFFFF", fgFill = BG_LEC,
  halign = "center", valign = "center",
  textDecoration = "bold", wrapText = TRUE,
  fontSize = 10, fontName = "Arial")

hs_mat <- openxlsx::createStyle(
  fontColour = "#FFFFFF", fgFill = BG_MAT,
  halign = "center", valign = "center",
  textDecoration = "bold", wrapText = TRUE,
  fontSize = 10, fontName = "Arial")

estilo_dato <- openxlsx::createStyle(
  fontName = "Arial", fontSize = 10,
  halign = "center", valign = "center",
  border = "TopBottomLeftRight", borderColour = "#CCCCCC")

estilo_dato_izq <- openxlsx::createStyle(
  fontName = "Arial", fontSize = 10,
  halign = "left", valign = "center", indent = 1,
  border = "TopBottomLeftRight", borderColour = "#CCCCCC")

estilo_num <- openxlsx::createStyle(
  fontName = "Arial", fontSize = 10,
  halign = "center", valign = "center",
  numFmt = "0.0",
  border = "TopBottomLeftRight", borderColour = "#CCCCCC")

grado_fill <- list(
  "Segundo Grado" = BG_2DO,
  "Tercer Grado"  = BG_3ER,
  "Cuarto Grado"  = BG_4TO
)

# ============================================================
# HOJA 1: ESTUDIANTES
# ============================================================
openxlsx::addWorksheet(wb, "Estudiantes")

openxlsx::mergeCells(wb, "Estudiantes", cols = 1:4, rows = 1)
openxlsx::writeData(wb, "Estudiantes", x = "Identificación del Estudiante", startCol = 1, startRow = 1)
openxlsx::addStyle(wb, "Estudiantes", hs_gen, rows = 1, cols = 1:4, gridExpand = TRUE)

openxlsx::mergeCells(wb, "Estudiantes", cols = 5:6, rows = 1)
openxlsx::writeData(wb, "Estudiantes", x = "Lectura (LEC)", startCol = 5, startRow = 1)
openxlsx::addStyle(wb, "Estudiantes", hs_lec, rows = 1, cols = 5:6, gridExpand = TRUE)

openxlsx::mergeCells(wb, "Estudiantes", cols = 7:8, rows = 1)
openxlsx::writeData(wb, "Estudiantes", x = "Matemática (MAT)", startCol = 7, startRow = 1)
openxlsx::addStyle(wb, "Estudiantes", hs_mat, rows = 1, cols = 7:8, gridExpand = TRUE)

openxlsx::setRowHeights(wb, "Estudiantes", rows = 1, heights = 22)

cab_stu <- c(
  "Código de\nInfraestructura",
  "Nombre del Centro",
  "Grado",
  "NIE",
  "score.final\n(escala 0-100)",
  "Categoría",
  "score.final\n(escala 0-100)",
  "Categoría"
)
openxlsx::writeData(wb, "Estudiantes", x = as.data.frame(t(cab_stu)), startCol = 1, startRow = 2, colNames = FALSE)
openxlsx::addStyle(wb, "Estudiantes", hs_gen, rows = 2, cols = 1:4, gridExpand = TRUE)
openxlsx::addStyle(wb, "Estudiantes", hs_lec, rows = 2, cols = 5:6, gridExpand = TRUE)
openxlsx::addStyle(wb, "Estudiantes", hs_mat, rows = 2, cols = 7:8, gridExpand = TRUE)
openxlsx::setRowHeights(wb, "Estudiantes", rows = 2, heights = 36)

df_stu_write <- data.frame(
  `Código de infraestructura` = estudiantes$`Código de infraestructura`,
  `Nombre del centro`         = estudiantes$`Nombre del centro`,
  Grado                       = estudiantes$Grado,
  NIE                         = estudiantes$NIE,
  score_LEC                   = estudiantes$score.final.LEC,
  cat_LEC                     = estudiantes$cat.LEC,
  score_MAT                   = estudiantes$score.final.MAT,
  cat_MAT                     = estudiantes$cat.MAT,
  check.names = FALSE,
  stringsAsFactors = FALSE
)

openxlsx::writeData(wb, "Estudiantes", x = df_stu_write, startRow = 3, startCol = 1, colNames = FALSE)

n_stu <- nrow(df_stu_write)
filas_stu <- seq(3, 2 + n_stu)

openxlsx::addStyle(wb, "Estudiantes", estilo_dato, rows = filas_stu, cols = c(1,3,4,5,7), gridExpand = TRUE)
openxlsx::addStyle(wb, "Estudiantes", estilo_dato_izq, rows = filas_stu, cols = 2, gridExpand = TRUE)
openxlsx::addStyle(wb, "Estudiantes", estilo_num, rows = filas_stu, cols = c(5,7), gridExpand = TRUE)

for (grado in names(grado_fill)) {
  filas_g <- which(df_stu_write$Grado == grado) + 2
  if (length(filas_g) == 0) next
  fill_hex <- grado_fill[[grado]]
  
  openxlsx::addStyle(wb, "Estudiantes",
                     openxlsx::createStyle(fontName="Arial", fontSize=10, halign="center", valign="center",
                                           fgFill=fill_hex, border="TopBottomLeftRight", borderColour="#CCCCCC"),
                     rows=filas_g, cols=c(1,3,4), gridExpand=TRUE)
  
  openxlsx::addStyle(wb, "Estudiantes",
                     openxlsx::createStyle(fontName="Arial", fontSize=10, halign="left", valign="center", indent=1,
                                           fgFill=fill_hex, border="TopBottomLeftRight", borderColour="#CCCCCC"),
                     rows=filas_g, cols=2, gridExpand=TRUE)
  
  openxlsx::addStyle(wb, "Estudiantes",
                     openxlsx::createStyle(fontName="Arial", fontSize=10, halign="center", valign="center", numFmt="0.0",
                                           fgFill=fill_hex, border="TopBottomLeftRight", borderColour="#CCCCCC"),
                     rows=filas_g, cols=c(5,7), gridExpand=TRUE)
}

for (nivel in names(NIVEL_COLORS)) {
  st <- estilo_nivel(nivel)
  filas_lec_nivel <- which(df_stu_write$cat_LEC == nivel) + 2
  if (length(filas_lec_nivel) > 0)
    openxlsx::addStyle(wb, "Estudiantes", st, rows = filas_lec_nivel, cols = 6, gridExpand = TRUE)
  
  filas_mat_nivel <- which(df_stu_write$cat_MAT == nivel) + 2
  if (length(filas_mat_nivel) > 0)
    openxlsx::addStyle(wb, "Estudiantes", st, rows = filas_mat_nivel, cols = 8, gridExpand = TRUE)
}

openxlsx::setRowHeights(wb, "Estudiantes", rows = filas_stu, heights = rep(16, n_stu))
openxlsx::setColWidths(wb, "Estudiantes", cols = 1:8, widths = c(20, 42, 18, 14, 18, 14, 18, 14))
openxlsx::freezePane(wb, "Estudiantes", firstActiveRow = 3)

fila_nota_stu <- n_stu + 4
openxlsx::mergeCells(wb, "Estudiantes", cols = 1:8, rows = fila_nota_stu)
openxlsx::writeData(wb, "Estudiantes", x = paste0(
  "Nota: score.final (escala 0-100) es el puntaje IRT equiparado (método Stocking-Lord, referencia 2do Grado), re-escalado a la ",
  "distribución poblacional combinada (media=50, SD=17). Categorías: Crítico ≤35 | Bajo 36-45 | Medio 46-55 | Bueno 56-65 | Excelente >65."
), startRow = fila_nota_stu, startCol = 1, colNames = FALSE)
openxlsx::addStyle(wb, "Estudiantes",
                   openxlsx::createStyle(fontName="Arial", fontSize=9, fontColour="#666666", textDecoration="italic",
                                         halign="left", valign="center", wrapText=TRUE), rows = fila_nota_stu, cols = 1)
openxlsx::setRowHeights(wb, "Estudiantes", rows = fila_nota_stu, heights = 30)

# ============================================================
# HOJA 2: ESCUELAS
# ============================================================
openxlsx::addWorksheet(wb, "Escuelas")

openxlsx::mergeCells(wb, "Escuelas", cols = 1:3, rows = 1)
openxlsx::writeData(wb, "Escuelas", x = "Identificación del Centro", startCol = 1, startRow = 1)
openxlsx::addStyle(wb, "Escuelas", hs_gen, rows = 1, cols = 1:3, gridExpand = TRUE)

openxlsx::mergeCells(wb, "Escuelas", cols = 4:6, rows = 1)
openxlsx::writeData(wb, "Escuelas", x = "Lectura (LEC)", startCol = 4, startRow = 1)
openxlsx::addStyle(wb, "Escuelas", hs_lec, rows = 1, cols = 4:6, gridExpand = TRUE)

openxlsx::mergeCells(wb, "Escuelas", cols = 7:9, rows = 1)
openxlsx::writeData(wb, "Escuelas", x = "Matemática (MAT)", startCol = 7, startRow = 1)
openxlsx::addStyle(wb, "Escuelas", hs_mat, rows = 1, cols = 7:9, gridExpand = TRUE)

openxlsx::writeData(wb, "Escuelas", x = "Estatus\nEscuela", startCol = 10, startRow = 1)
openxlsx::addStyle(wb, "Escuelas", hs_gen, rows = 1, cols = 10, gridExpand = TRUE)
openxlsx::mergeCells(wb, "Escuelas", cols = 10, rows = 1:2)

openxlsx::setRowHeights(wb, "Escuelas", rows = 1, heights = 22)

cab_esc <- c(
  "Código de\nInfraestructura",
  "Nombre del Centro",
  "N\nEstudiantes",
  "% Crítico\nEscuela",
  "% Crítico\nUniverso",
  "Gana\nLEC",
  "% Crítico\nEscuela",
  "% Crítico\nUniverso",
  "Gana\nMAT",
  ""
)
openxlsx::writeData(wb, "Escuelas", x = as.data.frame(t(cab_esc)), startCol = 1, startRow = 2, colNames = FALSE)
openxlsx::addStyle(wb, "Escuelas", hs_gen, rows = 2, cols = c(1,2,3,10), gridExpand = TRUE)
openxlsx::addStyle(wb, "Escuelas", hs_lec, rows = 2, cols = 4:6, gridExpand = TRUE)
openxlsx::addStyle(wb, "Escuelas", hs_mat, rows = 2, cols = 7:9, gridExpand = TRUE)
openxlsx::setRowHeights(wb, "Escuelas", rows = 2, heights = 36)

df_esc_write <- data.frame(
  cod      = escuelas$`Código de infraestructura`,
  nombre   = escuelas$`Nombre del centro`,
  n_stu    = escuelas$`N Estudiantes`,
  pct_lec  = escuelas$`% Crítico LEC`,
  univ_lec = escuelas$`% Crítico universo LEC`,
  gana_lec = escuelas$`Gana LEC`,
  pct_mat  = escuelas$`% Crítico MAT`,
  univ_mat = escuelas$`% Crítico universo MAT`,
  gana_mat = escuelas$`Gana MAT`,
  estatus  = escuelas$Estatus,
  stringsAsFactors = FALSE
)

openxlsx::writeData(wb, "Escuelas", x = df_esc_write, startRow = 3, startCol = 1, colNames = FALSE)

n_esc <- nrow(df_esc_write)
filas_esc <- seq(3, 2 + n_esc)

openxlsx::addStyle(wb, "Escuelas", estilo_dato, rows = filas_esc, cols = c(1,3,4,5,6,7,8,9), gridExpand = TRUE)
openxlsx::addStyle(wb, "Escuelas", estilo_dato_izq, rows = filas_esc, cols = 2, gridExpand = TRUE)
openxlsx::addStyle(wb, "Escuelas", openxlsx::createStyle(fontName="Arial", fontSize=10, halign="center", valign="center",
                                                         numFmt="0.0", border="TopBottomLeftRight", borderColour="#CCCCCC"),
                   rows = filas_esc, cols = c(4,5,7,8), gridExpand = TRUE)

estilo_gana_si <- openxlsx::createStyle(fontName="Arial", fontSize=10, textDecoration="bold", fontColour="#276221",
                                        fgFill="#C6EFCE", halign="center", valign="center", border="TopBottomLeftRight", borderColour="#CCCCCC")
estilo_gana_no <- openxlsx::createStyle(fontName="Arial", fontSize=10, textDecoration="bold", fontColour="#9C0006",
                                        fgFill="#FFC7CE", halign="center", valign="center", border="TopBottomLeftRight", borderColour="#CCCCCC")

for (col_idx in c(6, 9)) {
  col_name <- if (col_idx == 6) "gana_lec" else "gana_mat"
  filas_si <- which(df_esc_write[[col_name]] == "✔ Sí") + 2
  filas_no <- which(df_esc_write[[col_name]] == "✘ No") + 2
  if (length(filas_si) > 0) openxlsx::addStyle(wb, "Escuelas", estilo_gana_si, rows=filas_si, cols=col_idx, gridExpand=TRUE)
  if (length(filas_no) > 0) openxlsx::addStyle(wb, "Escuelas", estilo_gana_no, rows=filas_no, cols=col_idx, gridExpand=TRUE)
}

for (est in names(ESTATUS_COLORS)) {
  filas_est <- which(df_esc_write$estatus == est) + 2
  if (length(filas_est) > 0)
    openxlsx::addStyle(wb, "Escuelas", estilo_estatus(est), rows=filas_est, cols=10, gridExpand=TRUE)
}

openxlsx::setRowHeights(wb, "Escuelas", rows = filas_esc, heights = rep(18, n_esc))
openxlsx::setColWidths(wb, "Escuelas", cols = 1:10, widths = c(20, 44, 12, 14, 14, 10, 14, 14, 10, 16))
openxlsx::freezePane(wb, "Escuelas", firstActiveRow = 3)

# Nota al pie hoja escuelas
fila_nota_esc <- n_esc + 4
openxlsx::mergeCells(wb, "Escuelas", cols = 1:10, rows = fila_nota_esc)
openxlsx::writeData(
  wb, "Escuelas",
  x = paste0(
    "Clasificación: Excelente (verde) = escuela con % Crítico ≤ universo en AMBAS asignaturas. ",
    "Regular (naranja) = ≤ universo en solo UNA asignatura. ",
    "Alerta (rojo) = % Crítico > universo en AMBAS asignaturas. ",
    "Universo = todas las escuelas de la muestra. ",
    "% Crítico = proporción de estudiantes con score.final ≤ 35."
  ),
  startRow = fila_nota_esc, startCol = 1, colNames = FALSE
)
openxlsx::addStyle(wb, "Escuelas",
                   openxlsx::createStyle(fontName="Arial", fontSize=9, fontColour="#666666",
                                         textDecoration="italic", halign="left", valign="center", wrapText=TRUE),
                   rows = fila_nota_esc, cols = 1)
openxlsx::setRowHeights(wb, "Escuelas", rows = fila_nota_esc, heights = 40)

# ── Guardar ───────────────────────────────────────────────────────────
openxlsx::saveWorkbook(wb, ARCHIVO_SALIDA, overwrite = TRUE)

cat(sprintf("\n  ✓ Archivo guardado: %s\n", ARCHIVO_SALIDA))
cat("\n [✓] Script 3 terminado.\n")