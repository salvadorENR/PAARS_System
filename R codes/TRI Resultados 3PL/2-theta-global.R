required_packages <- c("mirt","dplyr","readr","readxl")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readxl)
library(dplyr)
library(mirt)

# Definition of functions and variables ----

get_scaled_theta <- function(theta, scale = list(mean = 500, sd = 100), trunc = NULL) {
  sd.theta <- sd(theta, na.rm = T)
  mean.theta <- mean(theta, na.rm = T)
  scaled.theta <- (theta - mean.theta)/sd.theta
  scaled.theta <- scaled.theta * scale$sd + scale$mean
  if (!is.null(trunc)) {
    scaled.theta <- pmin(trunc$max, pmax(trunc$min, scaled.theta))
  }
  return(scaled.theta)
}


grades <- c("2do","3er","4to","5to","6to","7mo","8vo","9no","Bach-1er","Bach-2do")
lgrades <- as.list(grades)
names(lgrades) <- grades


# Calculate global theta - Math and Lecture performance ----


lapply(c("MAT","LEC"), FUN = function(subject) {
  
  
  resultados <- lapply(lgrades, FUN = function(grade) {
    
    md <- readRDS(paste0("./data/md/IRT/",subject,"-",grade,".rds"))
    
    df <- read_excel(paste0("./data/processed/",subject,"-exam.xlsx"), sheet = grade)
    p_df <- select(df, matches(paste0("^",subject,"\\d+")))
    resp <- df[,colnames(md@Data$tabdata)]
    
    itens_removidos <- setdiff(colnames(p_df), colnames(resp))
    
    df[["anular_prueba"]] <- NA
    df[["anular_prueba"]][
      df$duration < 5 | rstatix::is_extreme(df$duration)] <- "duración <5 min o tiempo extremo de demora"
    
    df[["tasa de aciertos (sin remover itens)"]] <- rowMeans(p_df, na.rm = T)
    df[["num itens removidos"]] <- length(itens_removidos)
    
    df[["tasa de aciertos (apenas itens validos)"]] <- rowMeans(resp, na.rm = T)
    
    df[["theta.global"]] <- as.data.frame(fscores(md, response.pattern = resp))$F1
    df[["theta.global (escala PISA)"]] <- get_scaled_theta(
      df$theta.global, scale = list(mean = 500, sd = 100)) 
    df[["theta.global (escala 0-100)"]] <- get_scaled_theta(
      df$theta.global, scale = list(mean = 50, sd = 17), trunc = list(min = 0, max = 100)) 
    
    df[["Obs."]] <- paste0("Itens removidos: ", paste(itens_removidos, collapse = ","))
    
    return(df)
  })
  
  writexl::write_xlsx(resultados, paste0("./results/",subject,"-resultados-CML-Junio.xlsx")) 
  
})