required_packages <- c("rstatix", "stringr","dplyr","readr","readxl")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readr)
library(dplyr)
library(stringr)
library(readxl)

# Lectura de los metadatos de las pruebas ----

Pruebas_SV <- read_excel("./data/raw/pruebas.xlsx")
Pruebas_SV <- unique(Pruebas_SV[,c("PruebaCodigo","PruebaTitulo")]) %>%
  mutate(
    Area = case_when(
      str_detect(`PruebaTitulo`, regex("Lectura", ignore_case = TRUE)) ~ "LEC",
      str_detect(`PruebaTitulo`, regex("Matemática", ignore_case = TRUE)) ~ "MAT",
      TRUE ~ NA_character_
    ),
    Grado = str_extract(`PruebaTitulo`, "\\d+°")
  )
Pruebas_SV$PruebaCodigo <- as.character(as.integer(Pruebas_SV$PruebaCodigo))

mapping <- c("2°" = "2do", "3°" = "3er", "4°" = "4to", "5°" = "5to","6°" = "6to",
             "7°" = "7mo", "8°" = "8vo", "9°" = "9no", "10°" = "Bach-1er", "11°" = "Bach-2do")
Pruebas_SV$Grado <- mapping[as.character(Pruebas_SV$Grado)]

Pruebas_SV <- Pruebas_SV[!is.na(Pruebas_SV$Grado),]

# Lectura de los datos de las pruebas (transformación de txt en data.frames) ----

process_exam <- function(
    Pruebas_SV, prefix = "LEC",
    base_path = "./data/raw/Reporte CML 2 puntajes/", is_nominal = F) {
  
  exams <- Pruebas_SV[Pruebas_SV$Area == prefix,]
  
  list_exams <- as.list(paste0(base_path, exams[["PruebaCodigo"]],".txt"))
  names(list_exams) <- exams$Grado  
  
  lapply(list_exams, FUN = function(file_path) {
    print(file_path)
    df <- read_delim(
      file = file_path,
      delim = "|",
      quote = "",
      trim_ws = TRUE,
      locale = locale(encoding = "ISO8859-1"),  
      show_col_types = FALSE)
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
      as.POSIXct(df$`Fecha-Hora de Fin`, format = "%d/%m/%y %H:%M",tz = "UTC"),
      as.POSIXct(df$`Fecha-Hora de Inicio`, format = "%d/%m/%y %H:%M",tz = "UTC"), units = "mins"))
    
    df <- select(df, -matches(paste0("^Fecha\\.+"))) %>%
      select(-c("Puntaje Total")) %>%
      mutate(across(where(is.character), ~ifelse(nchar(.) > 32767, substr(., 1, 32767), .)))
    
    return(df)
  })
}

# Extraer las respuestas de las pruebas para realizar TRI  ----

process_resp <- function(
    Pruebas_SV, prefix = "LEC",
    base_path = "./data/raw/Reporte CML 2 puntajes/", min_duration = 5, is_nominal = F) {
  
  list_df <- process_exam(Pruebas_SV, prefix, base_path, is_nominal = is_nominal)
  
  # ... solo se utilizan respuestas cuyo tiempo de execución es mayor a 5 minutos
  # ... y respuestas completas y cuyo tiempo de ejecución no sea un valor extremo
  lapply(list_df, FUN = function(df) {
    
    idx_remove <- rstatix::is_extreme(df$duration)
    
    resp <- df[(df$duration >= min_duration & !idx_remove),]
    resp <- select(resp, matches(paste0("^",prefix,"\\d+")))
    resp <- resp[complete.cases(resp),]
    
    if (nrow(df)-nrow(resp) > nrow(df)*0.25) {
      stop(paste0("ERROR: nro de examenes anulados muito maior do possivel en ",
                  prefix, " - Grado: ", unique(df$Grado), " - file: ", base_path))
    } else {
      print(paste0(" - Processados ",nrow(df), " observaciones para TRI del ", unique(df$Grado), " - file: ", base_path))
    }
    
    return(resp[rowSums(is.na(resp)) < ncol(resp), ])
  })
}

exam_MAT <- process_exam(Pruebas_SV, "MAT", "./data/raw/Reporte CML 2 puntajes/")
exam_LEC <- process_exam(Pruebas_SV, "LEC", "./data/raw/Reporte CML 2 puntajes/")

writexl::write_xlsx(exam_MAT, "./data/processed/MAT-exam.xlsx")
writexl::write_xlsx(exam_LEC, "./data/processed/LEC-exam.xlsx")

writexl::write_xlsx(process_resp(Pruebas_SV, "MAT"), "data/processed/IRT-MAT.xlsx")
writexl::write_xlsx(process_resp(Pruebas_SV, "LEC"), "data/processed/IRT-LEC.xlsx")

writexl::write_xlsx(
  process_resp(Pruebas_SV, prefix = "MAT",
               base_path = "./data/raw/Reporte CML 2 letras/", is_nominal = T),
  "./data/processed/IRT-MAT-nominal.xlsx")

writexl::write_xlsx(
  process_resp(Pruebas_SV, prefix = "LEC",
               base_path = "./data/raw/Reporte CML 2 letras/", is_nominal = T),
  "./data/processed/IRT-LEC-nominal.xlsx")