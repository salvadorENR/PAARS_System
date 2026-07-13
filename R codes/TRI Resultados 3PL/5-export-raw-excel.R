# =============================================================================
# 5-export-raw-excel.R
# Exporta los archivos .txt originales de Letras y Puntajes a formato Excel
# replicando exactamente el formato de los archivos de referencia:
#   - LEC: hojas L2, L3, ..., L11  (Lectura grados 2do a Bach-2do)
#   - MAT: hojas M2, M3, ..., M11  (Matemática grados 2do a Bach-2do)
# Prueba de Progreso - Mes 04 - Junio 2026
#
# Author: Geiser C Challco <geiser.challco@goes.gob.sv>
#
# Inputs:
#   - data/raw/Reporte junio puntajes/*.txt   (respuestas 0/1)
#   - data/raw/Reporte junio letras/*.txt     (respuestas nominales A/B/C/D)
#   - data/raw/pruebas.xlsx                  (metadata: PruebaCodigo por grado/area)
#
# Outputs (en results/datos_originales/):
#   - JUNIO_LECTURA_puntajes.xlsx
#   - JUNIO_MATEMATICAS_puntajes.xlsx
#   - JUNIO_LECTURA_letras.xlsx
#   - JUNIO_MATEMATICAS_letras.xlsx
# =============================================================================

required_packages <- c("readxl", "readr", "dplyr", "stringr", "writexl")
missing_packages  <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readxl)
library(readr)
library(dplyr)
library(stringr)
library(writexl)

# ---- Crear carpeta de salida ----
dir.create("./results/datos_originales", showWarnings = FALSE, recursive = TRUE)

# ---- Lectura de metadatos ----
Pruebas_SV <- read_excel("./data/raw/pruebas.xlsx")
Pruebas_SV <- unique(Pruebas_SV[, c("PruebaCodigo", "PruebaTitulo")]) %>%
  mutate(
    Area = case_when(
      str_detect(PruebaTitulo, regex("Lectura",    ignore_case = TRUE)) ~ "LEC",
      str_detect(PruebaTitulo, regex("Matemática", ignore_case = TRUE)) ~ "MAT",
      TRUE ~ NA_character_
    ),
    Grado = str_extract(PruebaTitulo, "\\d+°")
  )
Pruebas_SV$PruebaCodigo <- as.character(as.integer(Pruebas_SV$PruebaCodigo))

mapping <- c("2°" = "2do", "3°" = "3er",  "4°" = "4to",  "5°" = "5to",
             "6°" = "6to", "7°" = "7mo",  "8°" = "8vo",  "9°" = "9no",
             "10°" = "Bach-1er", "11°" = "Bach-2do")
Pruebas_SV$Grado <- mapping[as.character(Pruebas_SV$Grado)]
Pruebas_SV <- Pruebas_SV[!is.na(Pruebas_SV$Grado), ]

# Mapeo grado -> nombre de hoja (L2..L11 / M2..M11)
sheet_name_map <- c(
  "2do" = "2", "3er" = "3", "4to" = "4", "5to" = "5", "6to"  = "6",
  "7mo" = "7", "8vo" = "8", "9no" = "9", "Bach-1er" = "10", "Bach-2do" = "11"
)

# ---- Función para leer un .txt de reporte ----
read_report_txt <- function(file_path) {
  df <- read_delim(
    file      = file_path,
    delim     = "|",
    quote     = "",
    trim_ws   = TRUE,
    locale    = locale(encoding = "ISO8859-1"),
    show_col_types = FALSE
  )
  names(df) <- trimws(gsub('"', '', names(df)))
  df <- df %>%
    mutate(across(where(is.character), ~ trimws(gsub('"', '', .))))
  return(df)
}

# ---- Función principal: exportar un tipo de reporte ----
export_report <- function(prefix, base_path, tipo) {
  
  message("\n=== Exportando ", tipo, " - ", prefix, " ===")
  
  exams    <- Pruebas_SV[Pruebas_SV$Area == prefix, ]
  prefix_letter <- ifelse(prefix == "LEC", "L", "M")
  subject_name  <- ifelse(prefix == "LEC", "LECTURA", "MATEMATICAS")
  
  sheets <- list()
  
  for (i in seq_len(nrow(exams))) {
    grade      <- exams$Grado[i]
    prueba_id  <- exams$PruebaCodigo[i]
    sheet_key  <- sheet_name_map[grade]
    sheet_name <- paste0(prefix_letter, sheet_key)
    file_path  <- paste0(base_path, prueba_id, ".txt")
    
    message("  ", grade, " -> ", sheet_name, " (", file_path, ")")
    
    df <- tryCatch(
      read_report_txt(file_path),
      error = function(e) {
        warning("No se pudo leer: ", file_path, " - ", e$message)
        return(NULL)
      }
    )
    
    if (is.null(df)) next
    
    sheets[[sheet_name]] <- df
  }
  
  # Ordenar hojas por número (L2 < L3 ... < L11)
  sheet_order <- paste0(prefix_letter, c("2","3","4","5","6","7","8","9","10","11"))
  sheets <- sheets[intersect(sheet_order, names(sheets))]
  
  tipo_label <- ifelse(tipo == "puntajes", "puntajes", "letras")
  out_path <- paste0("./results/datos_originales/JUNIO_", subject_name, "_", tipo_label, ".xlsx")
  writexl::write_xlsx(sheets, out_path)
  message("  Guardado: ", out_path)
}

# ---- Ejecutar exportaciones ----

# Puntajes (respuestas 0/1)
export_report("LEC", "./data/raw/Reporte junio puntajes/", "puntajes")
export_report("MAT", "./data/raw/Reporte junio puntajes/", "puntajes")

# Letras (respuestas nominales A/B/C/D)
export_report("LEC", "./data/raw/Reporte junio letras/", "letras")
export_report("MAT", "./data/raw/Reporte junio letras/", "letras")

message("\n=== Exportación completada ===")