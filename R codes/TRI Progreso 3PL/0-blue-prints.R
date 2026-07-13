required_packages <- c("rstatix", "stringr", "dplyr", "readr", "readxl", "writexl")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readr)
library(dplyr)
library(stringr)
library(readxl)

# Crear directorios de salida si no existen ----
dir.create("./data/processed", recursive = TRUE, showWarnings = FALSE)

# Lectura de los metadatos de las pruebas ----

Pruebas_SV_full <- read_excel("./data/raw/pruebas.xlsx") %>%
  mutate(
    Area = case_when(
      str_detect(`PruebaTitulo`, regex("Lectura", ignore_case = TRUE)) ~ "LEC",
      str_detect(`PruebaTitulo`, regex("Matemática", ignore_case = TRUE)) ~ "MAT",
      TRUE ~ NA_character_
    ),
    Grado = str_extract(`PruebaTitulo`, "\\d+°")
  )

mapping <- c("2°" = "2do", "3°" = "3er", "4°" = "4to", "5°" = "5to", "6°" = "6to",
             "7°" = "7mo", "8°" = "8vo", "9°" = "9no", "10°" = "Bach-1er", "11°" = "Bach-2do")
Pruebas_SV_full$Grado <- mapping[as.character(Pruebas_SV_full$Grado)]
Pruebas_SV_full$PruebaCodigo <- as.character(as.integer(Pruebas_SV_full$PruebaCodigo))
Pruebas_SV_full <- Pruebas_SV_full[!is.na(Pruebas_SV_full$Grado), ]

# Pruebas_SV reducido — solo lo que necesitan las funciones de proceso ----
Pruebas_SV <- unique(Pruebas_SV_full[, c("PruebaCodigo", "PruebaTitulo", "Area", "Grado")])

# Construcción del BluePrint ----

items_pruebas <- Pruebas_SV_full
items_pruebas$Item          <- items_pruebas$ItemCodig
items_pruebas$Bloque        <- paste0("B00", items_pruebas$Area, items_pruebas$BloqueCodigo)
items_pruebas$Dimension     <- paste0("D00", items_pruebas$Area, items_pruebas$Dimension)
items_pruebas$Subafirmacion <- paste0("S00", items_pruebas$Area, items_pruebas$Subafirmacion)

cnames <- c("Area", "Grado", "Item", "Bloque", "BloqueDescripcion",
            "Dimension", "Dimension Descripcion",
            "Subafirmacion", "Subafirmacion descripción")
items_pruebas <- unique(items_pruebas[, cnames])

get_legends <- function(items_pruebas, subject = "LEC") {
  items_pruebas <- items_pruebas[items_pruebas$Area == subject, ]
  
  legend_bloque <- unique(items_pruebas[, c("Bloque", "BloqueDescripcion")])
  legend_bloque <- legend_bloque[complete.cases(legend_bloque), ]
  
  legend_dimension <- unique(items_pruebas[, c("Dimension", "Dimension Descripcion")])
  legend_dimension <- legend_dimension[complete.cases(legend_dimension), ]
  
  legend_subafirmacion <- unique(items_pruebas[, c("Subafirmacion", "Subafirmacion descripción")])
  legend_subafirmacion <- legend_subafirmacion[complete.cases(legend_subafirmacion), ]
  
  return(list(
    legend_bloque        = legend_bloque,
    legend_dimension     = legend_dimension,
    legend_subafirmacion = legend_subafirmacion
  ))
}

for (subject in c("LEC", "MAT")) {
  toSave <- items_pruebas[items_pruebas$Area == subject, ]
  legs   <- get_legends(items_pruebas, subject = subject)
  writexl::write_xlsx(
    c(split(toSave, toSave$Grado), legs),
    paste0("./data/processed/BluePrint-", subject, ".xlsx")
  )
  print(paste0("BluePrint-", subject, ".xlsx creado correctamente."))
}

# Lectura de los datos de las pruebas (transformación de txt en data.frames) ----

process_exam <- function(
    Pruebas_SV, prefix = "LEC",
    base_path = "./data/raw/Reporte junio puntajes/", is_nominal = FALSE) {
  
  exams <- Pruebas_SV[Pruebas_SV$Area == prefix, ]
  
  list_exams <- as.list(paste0(base_path, exams[["PruebaCodigo"]], ".txt"))
  names(list_exams) <- exams$Grado
  
  lapply(list_exams, FUN = function(file_path) {
    print(file_path)
    df <- read_delim(
      file = file_path,
      delim = "|",
      quote = "",
      trim_ws = TRUE,
      locale = locale(encoding = "ISO8859-1"),
      show_col_types = FALSE
    )
    names(df) <- trimws(gsub('"', '', names(df)))
    df <- df %>%
      mutate(across(where(is.character), ~ trimws(gsub('"', '', .))))
    
    cols_char <- sapply(df, is.character)
    long_text_matrix <- sapply(df[, cols_char, drop = FALSE], function(col) {
      nchar(col) > 200
    })
    if (length(which(long_text_matrix, arr.ind = TRUE)) > 0) {
      stop("ERROR - existen celulas de tamaño inválido")
    }
    
    for (cname in colnames(df)[startsWith(colnames(df), prefix)]) {
      if (!is_nominal)
        df[[cname]] <- as.integer(df[[cname]])
      else
        df[[cname]] <- as.character(df[[cname]])
    }
    
    df$duration <- as.numeric(difftime(
      as.POSIXct(df$`Fecha-Hora de Fin`,    format = "%d/%m/%y %H:%M", tz = "UTC"),
      as.POSIXct(df$`Fecha-Hora de Inicio`, format = "%d/%m/%y %H:%M", tz = "UTC"),
      units = "mins"
    ))
    
    df <- select(df, -matches(paste0("^Fecha\\.+"))) %>%
      select(-c("Puntaje Total")) %>%
      mutate(across(where(is.character), ~ ifelse(nchar(.) > 32767, substr(., 1, 32767), .)))
    
    return(df)
  })
}

# Extraer las respuestas de las pruebas para realizar TRI ----

process_resp <- function(
    Pruebas_SV, prefix = "LEC",
    base_path = "./data/raw/Reporte junio puntajes/", min_duration = 5, is_nominal = FALSE) {
  
  list_df <- process_exam(Pruebas_SV, prefix, base_path, is_nominal = is_nominal)
  
  # ... solo se utilizan respuestas cuyo tiempo de ejecución es mayor a 5 minutos
  # ... y respuestas completas y cuyo tiempo de ejecución no sea un valor extremo
  lapply(list_df, FUN = function(df) {
    
    idx_remove <- rstatix::is_extreme(df$duration)
    
    resp <- df[(df$duration >= min_duration & !idx_remove), ]
    resp <- select(resp, matches(paste0("^", prefix, "\\d+")))
    resp <- resp[complete.cases(resp), ]
    
    if (nrow(df) - nrow(resp) > nrow(df) * 0.25) {
      stop(paste0("ERROR: nro de examenes anulados muito maior do possivel en ",
                  prefix, " - Grado: ", unique(df$Grado), " - file: ", base_path))
    } else {
      print(paste0(" - Processados ", nrow(df), " observaciones para TRI del ",
                   unique(df$Grado), " - file: ", base_path))
    }
    
    return(resp[rowSums(is.na(resp)) < ncol(resp), ])
  })
}

# Procesamiento y escritura de archivos de examen ----

exam_MAT <- process_exam(Pruebas_SV, "MAT", "./data/raw/Reporte junio puntajes/")
exam_LEC <- process_exam(Pruebas_SV, "LEC", "./data/raw/Reporte junio puntajes/")

writexl::write_xlsx(exam_MAT, "./data/processed/MAT-exam.xlsx")
writexl::write_xlsx(exam_LEC, "./data/processed/LEC-exam.xlsx")

writexl::write_xlsx(process_resp(Pruebas_SV, "MAT"), "./data/processed/IRT-MAT.xlsx")
writexl::write_xlsx(process_resp(Pruebas_SV, "LEC"), "./data/processed/IRT-LEC.xlsx")

writexl::write_xlsx(
  process_resp(Pruebas_SV, prefix = "MAT",
               base_path = "./data/raw/Reporte junio letras/", is_nominal = TRUE),
  "./data/processed/IRT-MAT-nominal.xlsx"
)

writexl::write_xlsx(
  process_resp(Pruebas_SV, prefix = "LEC",
               base_path = "./data/raw/Reporte junio letras/", is_nominal = TRUE),
  "./data/processed/IRT-LEC-nominal.xlsx"
)