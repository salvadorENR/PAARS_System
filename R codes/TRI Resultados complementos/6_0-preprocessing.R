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
    base_path = "./data/raw/Reporte junio puntajes/", is_nominal = F) {
  
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


exam_MAT <- process_exam(Pruebas_SV, "MAT", "./data/raw/Reporte junio puntajes/")
exam_LEC <- process_exam(Pruebas_SV, "LEC", "./data/raw/Reporte junio puntajes/")


# Complementar respostas previamente no processadas existentes  ----

new_exam_MAT <- process_exam(Pruebas_SV, "MAT", "./data/raw/v2/puntajes/")
dif_exam_MAT <- lapply(names(new_exam_MAT), FUN = function(grade) {
  old_exam <- exam_MAT[[grade]]
  new_exam <- new_exam_MAT[[grade]]
  return(new_exam[new_exam$Documento %in% setdiff(new_exam$Documento, old_exam$Documento),])
})
names(dif_exam_MAT) <- names(new_exam_MAT)


new_exam_LEC <- process_exam(Pruebas_SV, "LEC", "./data/raw/v2/puntajes/")
dif_exam_LEC <- lapply(names(new_exam_LEC), FUN = function(grade) {
  old_exam <- exam_LEC[[grade]]
  new_exam <- new_exam_LEC[[grade]]
  return(new_exam[new_exam$Documento %in% setdiff(new_exam$Documento, old_exam$Documento),])
})
names(dif_exam_LEC) <- names(new_exam_LEC)


writexl::write_xlsx(dif_exam_MAT, "./data/processed/MAT-exam_pendiente.xlsx")
writexl::write_xlsx(dif_exam_LEC, "./data/processed/LEC-exam_pendiente.xlsx")

