required_packages <- c("templates","dplyr","readr","readxl")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(dplyr)
library(readxl)
library(templates)

## Processando modelos de TRI a partir del template ----


lapply(c("MAT","LEC"), FUN = function(subject) {
  lapply(excel_sheets(paste0("./data/processed/IRT-",subject,".xlsx")), FUN = function(grade) {
    txt <- do.call(
      tmpl, c(list(".t" = paste(readLines("./templates/irt_analysis.Rmd"), collapse="\n")),
              list(subject = subject, grade = grade, other = " - Progreso Mes 04 - Aplicación Junio")))
    writeLines(txt, paste0("./code/IRT/",subject,"/",grade,".Rmd"), useBytes=T) 
  })
})


lapply(c("LEC","MAT"), FUN = function(subject) {
  lapply(excel_sheets(paste0("./data/processed/IRT-",subject,".xlsx")), FUN = function(grade) {
    for (output_format in c("word_document")) {
      file.input <- paste0("./code/IRT/",subject,"/",grade,".Rmd")
      rmarkdown::render(file.input, output_dir = paste0('./results/IRT/',subject,'/'),
                        clean = T, output_format = output_format)
    }
  })
})


## Armazendo parametros de IRT como archivos de excel ----

lapply(c("LEC","MAT"), FUN = function(subject) {
  
  grades <- excel_sheets(paste0("data/processed/IRT-",subject,".xlsx"))
  lgrades <- as.list(grades)
  names(lgrades) <- grades
  
  toSave <- lapply(lgrades, FUN = function(grade) {
    df <- read_excel(paste0("./data/tmp/",subject,"-",grade,".xlsx"), sheet = "reviewed")
  })
  writexl::write_xlsx(toSave, paste0("./data/processed/",subject,"-IRT-params.xlsx"))
})