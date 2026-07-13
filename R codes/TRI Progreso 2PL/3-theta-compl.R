required_packages <- c("mirt","tibble","readxl","dplyr","psych","lavaan")
missing_packages <- required_packages[!required_packages %in% rownames(installed.packages())]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

library(readxl)
library(tibble)
library(psych)
library(lavaan)
library(mirt)

# ---- Declaración de funciones y variables ----


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

best_omega_psych <- function(df, cnames = colnames(df), k_min = 3) {
  if (!requireNamespace("psych", quietly = TRUE)) {
    stop("O pacote 'psych' é necessário.")
  }
  
  stopifnot(k_min >= 3)
  
  X <- as.data.frame(df[, cnames, drop = FALSE])
  
  if (ncol(X) < k_min) {
    stop("O número de itens em 'cnames' é menor que k_min.")
  }
  
  # ... critérios fixos da heurística
  loading_cutoff <- 0.30   # carga mínima aceitável
  omega_target   <- 0.60   # omega mínimo aceitável (uso exploratório)
  
  # ... helpers ...
  
  # Matriz de correlação apropriada para itens 0/1
  get_cor_mat <- function(dat) {
    suppressWarnings(psych::tetrachoric(dat)$rho)
  }
  
  # Ajusta 1 fator via psych::fa para obter cargas
  fit_onefactor <- function(dat) {
    R <- get_cor_mat(dat)
    
    fa_fit <- psych::fa(
      r        = R,
      nfactors = 1,
      fm       = "minres",
      rotate   = "none"
    )
    
    list(fa = fa_fit, R = R)
  }
  
  # Extrai cargas padronizadas do 1 fator
  get_loadings <- function(fit_obj, item_names) {
    L <- as.matrix(fit_obj$fa$loadings)
    vals <- as.numeric(L[, 1])
    stats::setNames(vals, item_names)
  }
  
  # Calcula omega total com psych::omega
  get_omega <- function(dat) {
    R <- get_cor_mat(dat)
    
    om <- suppressWarnings(
      psych::omega(
        m       = R,
        nfactors = 1,
        plot    = FALSE,
        n.obs   = nrow(dat)
      )
    )
    
    return(list(om = om, omega.tot = as.numeric(om$omega.tot)))
  }
  
  # ---- objetos de saída ----
  history <- data.frame(
    step        = integer(0),
    k           = integer(0),
    removed     = character(0),
    min_loading = numeric(0),
    omega       = numeric(0),
    stringsAsFactors = FALSE
  )
  
  removed_items <- character(0)
  step <- 0L
  
  repeat {
    fit <- tryCatch(fit_onefactor(X), error = function(e) { e })
    
    if (inherits(fit, "error")) {
      stop(sprintf("A análise fatorial falhou na iteração %d: %s", step, fit$message))
    }
    
    loadings <- get_loadings(fit, colnames(X))
    omega    <- tryCatch(get_omega(X), error = function(e) NA_real_)
    if (step < 1) { omega_initial <- omega }
    
    min_item    <- names(which.min(loadings))
    min_loading <- min(loadings, na.rm = TRUE)
    
    history <- rbind(
      history,
      data.frame(
        step        = step,
        k           = ncol(X),
        removed     = if (step == 0L) NA_character_ else removed_items[length(removed_items)],
        min_loading = unname(min_loading),
        omega       = unname(omega$omega.tot),
        stringsAsFactors = FALSE
      )
    )
    
    # critério de parada
    if (ncol(X) <= k_min) break
    if (is.finite(min_loading) && is.finite(omega$omega.tot) &&
        (min_loading >= loading_cutoff || omega$omega.tot >= omega_target)) {
      break
    }
    
    # remove o item com menor carga
    removed_items <- c(removed_items, min_item)
    X <- X[, setdiff(colnames(X), min_item), drop = FALSE]
    step <- step + 1L
  }
  
  # refit final garantido
  fit_final <- fit_onefactor(X)
  loadings_final <- get_loadings(fit_final, colnames(X))
  omega_final <- tryCatch(get_omega(X), error = function(e) NA_real_)
  
  return(list(
    omega_initial  = omega_initial,
    items_final    = colnames(X),
    k_final        = ncol(X),
    removed_items  = removed_items,
    loadings_final = sort(loadings_final, decreasing = TRUE),
    omega_final    = omega_final,
    fit_final      = fit_final,
    history        = history,
    rules = list(
      loading_cutoff = loading_cutoff,
      omega_target   = omega_target
    )
  ))
}


reliability_omega <- function(md_df, bluePrint, fatores, item_col = "Item") {
  library(tibble)
  library(dplyr)
  library(psych)
  library(lavaan)
  
  lfatores <- as.list(fatores)
  names(lfatores) <- fatores
  
  # get tbl omega 
  get_tbl_omega <- function(best) {
    omega_tbl <- as_tibble(best$omega_initial$om$omega.group[1,])
    colnames(omega_tbl) <- paste0("omega.", colnames(omega_tbl))
    
    omega_tbl$alpha <- best$omega_initial$om$alpha
    omega_tbl <- mutate(omega_tbl, best = list(best))
    
    omega_tbl$Obs <- NA
    if (best$omega_initial$omega.tot < 0.6) {
      omega_tbl$Obs <- "Inadecuado"
    }
    return(omega_tbl)
  }
  
  df <- do.call(plyr::rbind.fill, lapply(lfatores, FUN = function(fator) {
    f_df <- do.call(plyr::rbind.fill, group_by_at(bluePrint, fator)  %>% group_map(~ {
      cnames <- c(intersect(.x[[item_col]], colnames(md_df)))
      if (length(cnames) >= 3) {
        best <- best_omega_psych(md_df[,cnames], cnames)
        return(cbind(.y, get_tbl_omega(best), n_items = length(cnames)))
      }
    }))
    return(f_df[!is.na(f_df[[fator]]), ])
  }))
  
  best <- best_omega_psych(md_df)
  df <- plyr::rbind.fill(cbind(get_tbl_omega(best), n_items = ncol(md_df)), df)
  return(df)
}





# ---- Calculating theta for competences and dimension/domain ----

grades <- c("2do","3er","4to","5to","6to","7mo","8vo","9no","Bach-1er","Bach-2do")
lgrades <- as.list(grades); names(lgrades) <- grades

fatores <- c("Bloque","Dimension","Subafirmacion")
lfatores <- as.list(fatores)
names(lfatores) <- fatores

lapply(c("MAT","LEC"), FUN = function(subject) {
  
  toSave <- lapply(lgrades, FUN = function(grade) {
    
    print(paste0("... processing ",subject," - grade: ",grade," ..."))
    
    exam_with_theta <- data.frame()
    rds_file <- paste0("./data/md/IRT/Compl-",subject,"-",grade,".rds")
    if (!file.exists(rds_file)) {
      md <- readRDS(paste0("./data/md/IRT/",subject,"-",grade,".rds"))
      md_df <- as.data.frame(md@Data$data)
      
      bluePrint <- read_excel(paste0("./data/processed/BluePrint-",subject,".xlsx"), sheet = grade)
      
      # calculation of omegas ----
      
      omega_tbl <- reliability_omega(md_df, bluePrint, fatores)
      
      new_omega_tbl <- do.call(
        plyr::rbind.fill, lapply(1:nrow(omega_tbl), FUN = function(i) {
          
          if (is.na(omega_tbl$Obs[i])) {
            to_return <- omega_tbl[i,]
            
            best <- omega_tbl$best[i][[1]]
            
            to_return$items_final <- list(sub("-$", "", rownames(best$omega_initial$om$R)))
            to_return$removed_items <- list(c())
            to_return$n_items <- nrow(best$omega_final$om$R)
            
            return(to_return)
          } else {
            best <- omega_tbl$best[i][[1]]
            if (best$omega_final$omega.tot >= 0.6) {
              to_return <- omega_tbl[i,]
              
              to_return$n_items <- best$k_final
              to_return$omega.total <- best$omega_final$om$omega.group$total[1]
              to_return$omega.general <- best$omega_final$om$omega.group$general[1]
              to_return$omega.group <- best$omega_final$om$omega.group$group[1]
              
              to_return$items_final <- list(best$items_final)
              to_return$removed_items <- list(best$removed_items)
              
              to_return$Obs <- NA
              
              return(to_return)
            }
          }
        }))
      
      
      exam_dat <- read_excel(paste0("./data/processed/",subject,"-exam.xlsx"), sheet = grade)
      
      lresult_mirt <- lapply(lfatores, FUN = function(fator) {
        
        # calculation of multidimensional mirt models ----
        
        tbl  <- new_omega_tbl[!is.na(new_omega_tbl[[fator]]), ]
        if (nrow(tbl) > 0) {
          mirt_model <- do.call(paste0, lapply(1:nrow(tbl), FUN = function(i) {
            paste0(tbl[[fator]][i] , " = ", paste0(tbl$items_final[i][[1]], collapse = ", "), " \n ")
          }))
          
          cnames <- unique(unlist(tbl$items_final))
          if (length(setdiff(cnames, colnames(md_df)) > 0)) {
            print(paste0("Warning: some items in the model are not in the data frame: ", 
                         paste0(setdiff(cnames, colnames(md_df)), collapse = ", ")))
          }
          md_sub <- mirt(md_df[, cnames], model = mirt_model, itemtype = "2PL", verbose = F)
          
          # estimate theta per themes and competences ----
          
          resp <- exam_dat[, cnames]
          theta <- as.data.frame(fscores(md_sub, response.pattern = resp))
          
          return(list(fator = fator, md = md_sub, items = cnames, theta = theta))  
        }
      })
      
      exam_with_theta <- exam_dat
      for (key in names(lresult_mirt)) {
        if (length(lresult_mirt[[key]]) > 0) {
          
          theta_df <- lresult_mirt[[key]]$theta
          cnames <- colnames(theta_df)[!startsWith(colnames(theta_df),"SE_")]
          for (cname in cnames) {
            theta_df[[paste0(cname," (escala PISA)")]] <- get_scaled_theta(
              theta_df[[cname]], scale = list(mean = 500, sd = 100))
            theta_df[[paste0(cname," (escala 0-100)")]] <- get_scaled_theta(
              theta_df[[cname]], scale = list(mean = 50, sd = 17), trunc = list(min = 0, max = 100))
          }
          exam_with_theta <- cbind(exam_with_theta, theta_df)
        }
      }
      
      # write the data in md folder ----
      
      saveRDS(list(
        exam_with_theta = exam_with_theta,
        omega_tbl = omega_tbl,
        new_omega_tbl = new_omega_tbl,
        lresult_mirt = lresult_mirt), paste0(rds_file))
    } else {
      result <- readRDS(rds_file)
      exam_with_theta <- result$exam_with_theta
    }
    
    exam_with_theta[["anular_prueba"]] <- NA
    exam_with_theta[["anular_prueba"]][exam_with_theta$duration < 5 |
                                       rstatix::is_extreme(exam_with_theta$duration)] <- "duración <5 min o tiempo extremo de demora"
    
    return(exam_with_theta)
  })

  legend_bloque <- read_excel(
    paste0("./data/processed/BluePrint-",subject,".xlsx"), sheet = "legend_bloque")
  legend_dimension <- read_excel(
    paste0("./data/processed/BluePrint-",subject,".xlsx"), sheet = "legend_dimension")
  legend_subafirmacion <- read_excel(
    paste0("./data/processed/BluePrint-",subject,".xlsx"), sheet = "legend_subafirmacion")
  
  writexl::write_xlsx(c(
    toSave, list(legend_bloque = legend_bloque, legend_dimension = legend_dimension, legend_subafirmacion = legend_subafirmacion)),
    paste0("./results/",subject,"-resultados_compl-Mes03-Junio.xlsx"))
    
})

