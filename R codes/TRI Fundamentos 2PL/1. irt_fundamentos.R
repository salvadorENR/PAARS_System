# ============================================================
# PAARS — Script 1: Pipeline IRT 2PL · Prueba de Fundamentos
#
# Ajusta modelos 2PL independientes por asignatura y grado,
# calcula theta global y por bloque/grupo, y exporta resultados.
#
# Único cambio respecto a la versión base: al final de cada grado
# se guarda el modelo final como .rds en results/models/ para que
# el Script 2 (equate_fundamentos.R) pueda leerlo sin re-ajustar.
#
# Lógica replicada de:
#   0-preprocessing.R   → carga y limpieza
#   1-irt-analysis.R    → ajuste 2PL, flagging de ítems
#   2-theta-global.R    → theta global + escalas
#   3-theta-compl.R     → omega iterativo + theta por bloque/grupo
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(writexl)
  library(openxlsx)
  library(mirt)
  library(psych)
  library(dplyr)
  library(plyr)
})

# ============================================================
# CONFIGURACIÓN
# ============================================================

# 1. Leer archivo de entorno para la unidad de Drive local
#readRenviron(".Renviron")
base_drive <- Sys.getenv("BASE_DRIVE")
if (base_drive == "") {
  stop("[!] ERROR: Variable BASE_DRIVE no encontrada en archivo .Renviron")
}

# 2. Leer los argumentos enviados por Python (La carpeta del mes seleccionado)
args <- commandArgs(trailingOnly = TRUE)
if (length(args) == 0) {
  stop("[!] ERROR: R no recibió la ruta de la carpeta del mes desde Python.")
}
carpeta_mes <- args[1]

# 3. Construir rutas dinámicas según requerimientos
# Entrada: 02_Datasets_Procesados / Salida: results en la misma carpeta
RUTA_DICOTOMICO  <- file.path(carpeta_mes, "02_Datasets_Procesados", "2_Resultados_Fundamento_Dicotomico_0_1.xlsx")
RUTA_BLUEPRINT   <- file.path(carpeta_mes, "01_Formularios_Crudos", "BluePrint-Fundamentos.xlsx")
CARPETA_RESULTADOS <- file.path(carpeta_mes, "results")
CARPETA_MODELOS    <- file.path(CARPETA_RESULTADOS, "models")   

if (!dir.exists(CARPETA_RESULTADOS)) dir.create(CARPETA_RESULTADOS, recursive = TRUE)
if (!dir.exists(CARPETA_MODELOS))    dir.create(CARPETA_MODELOS,    recursive = TRUE)

ID_COLS <- c('NIE', 'Código de infraestructura', 'Nombre del centro',
             'Grado', 'Grupo', 'Marca temporal',
             'Dirección de correo electrónico', 'Puntuación',
             'Puntaje_Global_Formulario')

ITEMS_LEC <- paste0("L_item_", 1:68)
ITEMS_MAT <- paste0("M_item_", 1:18)

ORDEN_GRADOS <- c('Segundo Grado', 'Tercer Grado', 'Cuarto Grado')
SHEET_NAMES  <- c('Segundo Grado' = '2do',
                  'Tercer Grado'  = '3er',
                  'Cuarto Grado'  = '4to')

# ============================================================
# FUNCIÓN: get_scaled_theta
# Réplica exacta de 2-theta-global.R + 3-theta-compl.R
# ============================================================
get_scaled_theta <- function(theta,
                             scale = list(mean = 500, sd = 100),
                             trunc = NULL) {
  sd.theta   <- sd(theta,   na.rm = TRUE)
  mean.theta <- mean(theta, na.rm = TRUE)
  if (is.na(sd.theta) || sd.theta == 0) return(theta)
  scaled <- (theta - mean.theta) / sd.theta * scale$sd + scale$mean
  if (!is.null(trunc)) {
    scaled <- pmin(trunc$max, pmax(trunc$min, scaled))
  }
  return(scaled)
}

# ============================================================
# FUNCIÓN: best_omega_psych
# Réplica exacta de 3-theta-compl.R
# ============================================================
best_omega_psych <- function(df, cnames = colnames(df), k_min = 3) {
  
  stopifnot(k_min >= 3)
  X <- as.data.frame(df[, cnames, drop = FALSE])
  if (ncol(X) < k_min)
    stop("Número de ítems menor que k_min.")
  
  loading_cutoff <- 0.30
  omega_target   <- 0.60
  
  get_cor_mat <- function(dat)
    suppressWarnings(psych::tetrachoric(dat)$rho)
  
  fit_onefactor <- function(dat) {
    R <- get_cor_mat(dat)
    fa_fit <- psych::fa(r = R, nfactors = 1, fm = "minres", rotate = "none")
    list(fa = fa_fit, R = R)
  }
  
  get_loadings <- function(fit_obj, item_names) {
    L <- as.matrix(fit_obj$fa$loadings)
    stats::setNames(as.numeric(L[, 1]), item_names)
  }
  
  get_omega <- function(dat) {
    R <- get_cor_mat(dat)
    om <- suppressWarnings(
      psych::omega(m = R, nfactors = 1, plot = FALSE, n.obs = nrow(dat))
    )
    list(om = om, omega.tot = as.numeric(om$omega.tot))
  }
  
  removed_items <- character(0)
  step          <- 0L
  omega_initial <- NULL
  
  repeat {
    fit <- tryCatch(fit_onefactor(X), error = function(e) e)
    if (inherits(fit, "error"))
      stop(sprintf("FA falló en iteración %d: %s", step, fit$message))
    
    loadings <- get_loadings(fit, colnames(X))
    omega    <- tryCatch(get_omega(X), error = function(e) list(omega.tot = NA_real_))
    
    if (step < 1) omega_initial <- omega
    
    min_item    <- names(which.min(loadings))
    min_loading <- min(loadings, na.rm = TRUE)
    
    if (ncol(X) <= k_min) break
    if (is.finite(min_loading) && is.finite(omega$omega.tot) &&
        (min_loading >= loading_cutoff || omega$omega.tot >= omega_target)) break
    
    removed_items <- c(removed_items, min_item)
    X    <- X[, setdiff(colnames(X), min_item), drop = FALSE]
    step <- step + 1L
  }
  
  fit_final      <- fit_onefactor(X)
  loadings_final <- get_loadings(fit_final, colnames(X))
  omega_final    <- tryCatch(get_omega(X), error = function(e) list(omega.tot = NA_real_))
  
  list(
    omega_initial  = omega_initial,
    items_final    = colnames(X),
    k_final        = ncol(X),
    removed_items  = removed_items,
    loadings_final = sort(loadings_final, decreasing = TRUE),
    omega_final    = omega_final,
    fit_final      = fit_final
  )
}

# ============================================================
# FUNCIÓN: reliability_omega
# Réplica exacta de 3-theta-compl.R
# ============================================================
reliability_omega <- function(md_df, bluePrint, fatores, item_col = "Ítem") {
  
  get_tbl_omega <- function(best) {
    omega_tbl <- as.data.frame(t(best$omega_initial$om$omega.group[1, ]))
    colnames(omega_tbl) <- paste0("omega.", colnames(omega_tbl))
    omega_tbl$omega.total   <- omega_tbl$omega.total   %||% NA
    omega_tbl$alpha         <- best$omega_initial$om$alpha
    omega_tbl               <- dplyr::mutate(omega_tbl, best = list(best))
    omega_tbl$Obs           <- NA_character_
    if (isTRUE(best$omega_initial$omega.tot < 0.6))
      omega_tbl$Obs <- "Inadecuado"
    return(omega_tbl)
  }
  
  df_all <- plyr::rbind.fill(lapply(fatores, FUN = function(fator) {
    f_df <- plyr::rbind.fill(
      dplyr::group_by_at(bluePrint, fator) %>%
        dplyr::group_map(~ {
          cnames <- intersect(.x[[item_col]], colnames(md_df))
          if (length(cnames) >= 3) {
            best <- best_omega_psych(md_df[, cnames], cnames)
            cbind(.y, get_tbl_omega(best), n_items = length(cnames))
          }
        })
    )
    f_df[!is.na(f_df[[fator]]), ]
  }))
  
  best_global <- best_omega_psych(md_df)
  df_global   <- cbind(get_tbl_omega(best_global), n_items = ncol(md_df))
  plyr::rbind.fill(df_global, df_all)
}

# ============================================================
# FUNCIÓN: item_flags
# Réplica exacta de la sección de flagging de los .Rmd
# ============================================================
item_flags <- function(md) {
  pars <- as.data.frame(coef(md, IRTpars = TRUE, simplify = TRUE)$items)
  pars$item <- rownames(pars)
  
  ifit <- tryCatch(itemfit(md, "infit"),      error = function(e) NULL)
  sx2  <- tryCatch(itemfit(md, fit_stats = "S_X2", na.rm = TRUE),
                   error = function(e) NULL)
  
  res <- pars
  if (!is.null(ifit)) res <- merge(res, ifit, by = "item", all.x = TRUE)
  if (!is.null(sx2))  res <- merge(res, sx2,  by = "item", all.x = TRUE)
  
  res$flag_a_nonpos       <- !is.na(res$a) & res$a <= 0
  res$flag_low_a_drop     <- !is.na(res$a) & res$a < 0.30
  res$flag_b_extreme      <- !is.na(res$b) & abs(res$b) > 5
  res$flag_infit_extreme  <- (!is.na(res$infit)  & (res$infit  < 0.5 | res$infit  > 1.5))
  res$flag_outfit_extreme <- (!is.na(res$outfit) & (res$outfit < 0.5 | res$outfit > 2.0))
  res$flag_misfit         <- (!is.na(res$RMSEA.S_X2) & res$RMSEA.S_X2 > 0.10)
  
  res$evidence_count <- (as.integer(res$flag_low_a_drop) +
                           as.integer(res$flag_b_extreme)  +
                           as.integer(res$flag_infit_extreme | res$flag_outfit_extreme) +
                           as.integer(res$flag_misfit))
  
  res$drop <- res$flag_a_nonpos | (res$evidence_count >= 2)
  return(res)
}

# ============================================================
# FUNCIÓN: procesar_grado
# Idéntica al original excepto que ahora retorna también
# 'model' e 'items_kept' para que el MAIN pueda guardarlos.
# ============================================================
procesar_grado <- function(df_g, items_list, bp, fatores) {
  
  items_p         <- intersect(items_list, colnames(df_g))
  id_cols_present <- intersect(ID_COLS, colnames(df_g))
  mat_data        <- as.data.frame(
    sapply(df_g[, items_p, drop = FALSE], as.numeric))
  
  # ── PASO 1: Modelo inicial 2PL ──────────────────────────────────────
  cat("    Ajustando modelo inicial 2PL ...\n")
  md1  <- mirt(mat_data, 1, itemtype = '2PL', verbose = FALSE)
  res  <- item_flags(md1)
  
  items_drop <- res$item[res$drop]
  items_kept <- res$item[!res$drop]
  cat(sprintf("    Ítems removidos: %d  → %s\n",
              length(items_drop),
              if (length(items_drop) > 0) paste(items_drop, collapse = ", ")
              else "ninguno"))
  
  # ── PASO 2: Modelo final ─────────────────────────────────────────────
  cat("    Ajustando modelo final 2PL ...\n")
  md_final <- if (length(items_drop) > 0)
    mirt(mat_data[, items_kept, drop = FALSE], 1, itemtype = '2PL',
         verbose = FALSE)
  else md1
  
  resp_final   <- df_g[, items_kept, drop = FALSE]
  theta_scores <- fscores(md_final, response.pattern = resp_final,
                          full.scores = TRUE, full.scores.SE = FALSE)
  theta_raw    <- as.numeric(theta_scores[, 1])
  
  # ── df_resultados (-resultados.xlsx) ─────────────────────────────────
  df_res <- df_g[, id_cols_present, drop = FALSE]
  for (it in items_p) df_res[[it]] <- mat_data[[it]]
  
  df_res[["anular_prueba"]] <- NA_character_
  
  df_res[["tasa de aciertos (sin remover itens)"]] <-
    rowMeans(mat_data[, items_p,    drop = FALSE], na.rm = TRUE)
  df_res[["num itens removidos"]]  <- length(items_drop)
  df_res[["tasa de aciertos (apenas itens validos)"]] <-
    rowMeans(mat_data[, items_kept, drop = FALSE], na.rm = TRUE)
  
  df_res[["theta.global"]] <- round(theta_raw, 4)
  df_res[["theta.global (escala PISA)"]] <-
    round(get_scaled_theta(theta_raw, list(mean = 500, sd = 100)), 1)
  df_res[["theta.global (escala 0-100)"]] <-
    round(get_scaled_theta(theta_raw, list(mean = 50, sd = 17),
                           trunc = list(min = 0, max = 100)), 1)
  
  df_res[["Obs."]] <- paste0("Itens removidos: ",
                             paste(items_drop, collapse = ","))
  
  # ── df_compl (-resultados_compl.xlsx) ────────────────────────────────
  cat("    Calculando omega iterativo + theta por bloque/grupo ...\n")
  
  omega_tbl <- reliability_omega(mat_data[, items_kept, drop = FALSE],
                                 bp, fatores)
  
  new_omega_tbl <- plyr::rbind.fill(lapply(seq_len(nrow(omega_tbl)), function(i) {
    if (is.na(omega_tbl$Obs[i])) {
      to_return <- omega_tbl[i, ]
      best <- omega_tbl$best[i][[1]]
      to_return$items_final   <- list(rownames(best$omega_initial$om$R))
      to_return$removed_items <- list(character(0))
      to_return$n_items       <- nrow(best$omega_final$om$R)
      return(to_return)
    } else {
      best <- omega_tbl$best[i][[1]]
      if (isTRUE(best$omega_final$omega.tot >= 0.6)) {
        to_return <- omega_tbl[i, ]
        to_return$n_items       <- best$k_final
        to_return$omega.total   <- best$omega_final$om$omega.group$total[1]
        to_return$items_final   <- list(best$items_final)
        to_return$removed_items <- list(best$removed_items)
        to_return$Obs           <- NA_character_
        return(to_return)
      }
    }
  }))
  
  desc_col <- c('Bloque_Cod' = 'Bloque_Descripción',
                'Grupo_Cod'  = 'Grupo_Descripción')
  
  df_compl <- df_g[, id_cols_present, drop = FALSE]
  for (it in items_p) df_compl[[it]] <- mat_data[[it]]
  
  for (fator in fatores) {
    tbl_f <- new_omega_tbl[!is.na(new_omega_tbl[[fator]]), ]
    if (nrow(tbl_f) == 0) next
    
    mirt_model_str <- paste0(
      sapply(seq_len(nrow(tbl_f)), function(i) {
        paste0(tbl_f[[fator]][i], " = ",
               paste(tbl_f$items_final[i][[1]], collapse = ", "))
      }), collapse = "\n")
    
    cnames_sub <- unique(unlist(tbl_f$items_final))
    md_sub <- tryCatch(
      mirt(mat_data[, cnames_sub, drop = FALSE],
           model    = mirt_model_str,
           itemtype = "2PL",
           verbose  = FALSE),
      error = function(e) {
        cat(sprintf("    [!] Error en modelo multidim %s: %s\n",
                    fator, e$message))
        NULL
      })
    
    if (is.null(md_sub)) next
    
    theta_df <- as.data.frame(
      fscores(md_sub, response.pattern = df_g[, cnames_sub, drop = FALSE],
              full.scores = TRUE, full.scores.SE = TRUE))
    
    cnames_theta <- colnames(theta_df)[!startsWith(colnames(theta_df), "SE_")]
    for (cn in cnames_theta) {
      desc_match <- bp[[desc_col[fator]]][bp[[fator]] == cn]
      label <- if (length(desc_match) > 0 && !is.na(desc_match[1]))
        desc_match[1] else cn
      
      theta_df[[paste0(label, " (escala PISA)")]] <-
        get_scaled_theta(theta_df[[cn]], list(mean = 500, sd = 100))
      theta_df[[paste0(label, " (escala 0-100)")]] <-
        get_scaled_theta(theta_df[[cn]], list(mean = 50, sd = 17),
                         trunc = list(min = 0, max = 100))
      
      se_cn <- paste0("SE_", cn)
      colnames(theta_df)[colnames(theta_df) == cn]    <- label
      colnames(theta_df)[colnames(theta_df) == se_cn] <- paste0("SE_", label)
    }
    
    df_compl <- cbind(df_compl, theta_df)
  }
  
  df_compl[["anular_prueba"]] <- NA_character_
  
  # model e items_kept se retornan para que el MAIN los guarde como .rds
  return(list(
    res        = df_res,
    compl      = df_compl,
    model      = md_final,
    items_kept = items_kept
  ))
}

# ============================================================
# MAIN
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  PAARS — Script 1: Pipeline IRT · Prueba de Fundamentos\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")

df <- read_excel(RUTA_DICOTOMICO) %>%
  dplyr::filter(Grado != 'SIN REGISTRO')

bp_lec <- read_excel(RUTA_BLUEPRINT, sheet = "BluePrint-LEC")
bp_mat <- read_excel(RUTA_BLUEPRINT, sheet = "BluePrint-MAT")

config <- list(
  list(asig = "LEC", items = ITEMS_LEC, bp = bp_lec,
       bg = "#033b6d", fatores = c("Bloque_Cod", "Grupo_Cod")),
  list(asig = "MAT", items = ITEMS_MAT, bp = bp_mat,
       bg = "#085041", fatores = c("Bloque_Cod", "Grupo_Cod"))
)

for (cfg in config) {
  res_l   <- list()
  compl_l <- list()
  
  for (grado in ORDEN_GRADOS) {
    if (!grado %in% df$Grado) next
    sname <- SHEET_NAMES[[grado]]
    df_g  <- df[df$Grado == grado, ]
    
    cat(sprintf("\n%s\n  %s — %s  (n=%d)\n%s\n",
                paste(rep("-", 50), collapse = ""),
                cfg$asig, grado, nrow(df_g),
                paste(rep("-", 50), collapse = "")))
    
    out <- procesar_grado(df_g, cfg$items, cfg$bp, cfg$fatores)
    
    if (!is.null(out)) {
      res_l[[sname]]   <- out$res
      compl_l[[sname]] <- out$compl
      
      # ── Guardar modelo .rds para el Script 2 ──────────────────────
      # Nombre: results/models/{LEC|MAT}-{2do|3er|4to}.rds
      # Replica la convención de 0-equateIRT.R (subject-grade.rds).
      ruta_rds <- file.path(CARPETA_MODELOS,
                            paste0(cfg$asig, "-", sname, ".rds"))
      saveRDS(out$model, ruta_rds)
      cat(sprintf("    ✓ Modelo guardado: %s\n", ruta_rds))
    }
  }
  
  # ── Exportar -resultados.xlsx ──────────────────────────────────────
  ruta_res <- file.path(CARPETA_RESULTADOS,
                        paste0(cfg$asig, "-resultados.xlsx"))
  wb_r <- openxlsx::createWorkbook()
  hs   <- openxlsx::createStyle(
    fontColour = "#FFFFFF", fgFill = cfg$bg,
    halign = "center", textDecoration = "bold", wrapText = TRUE)
  
  for (sn in names(res_l)) {
    openxlsx::addWorksheet(wb_r, sn)
    openxlsx::writeData(wb_r, sn, res_l[[sn]], headerStyle = hs)
    openxlsx::freezePane(wb_r, sn, firstRow = TRUE)
  }
  openxlsx::saveWorkbook(wb_r, ruta_res, overwrite = TRUE)
  cat(sprintf("\n    ✓ %s\n", ruta_res))
  
  # ── Exportar -resultados_compl.xlsx ───────────────────────────────
  ruta_compl <- file.path(CARPETA_RESULTADOS,
                          paste0(cfg$asig, "-resultados_compl.xlsx"))
  wb_c <- openxlsx::createWorkbook()
  
  for (sn in names(compl_l)) {
    openxlsx::addWorksheet(wb_c, sn)
    openxlsx::writeData(wb_c, sn, compl_l[[sn]], headerStyle = hs)
    openxlsx::freezePane(wb_c, sn, firstRow = TRUE)
  }
  
  leg_bloque <- cfg$bp %>%
    dplyr::select(Bloque_Cod, Bloque_Descripción) %>%
    dplyr::distinct() %>%
    dplyr::filter(!is.na(Bloque_Cod))
  
  leg_grupo <- cfg$bp %>%
    dplyr::select(Grupo_Cod, Grupo_Descripción, Bloque_Cod) %>%
    dplyr::distinct() %>%
    dplyr::filter(!is.na(Grupo_Cod))
  
  openxlsx::addWorksheet(wb_c, "legend_bloque")
  openxlsx::writeData(wb_c, "legend_bloque", leg_bloque, headerStyle = hs)
  openxlsx::addWorksheet(wb_c, "legend_grupo")
  openxlsx::writeData(wb_c, "legend_grupo",  leg_grupo,  headerStyle = hs)
  
  openxlsx::saveWorkbook(wb_c, ruta_compl, overwrite = TRUE)
  cat(sprintf("    ✓ %s\n", ruta_compl))
}

cat(paste(rep("=", 60), collapse = ""), "\n")
cat("  [✓] Script 1 terminado. Modelos guardados en:", CARPETA_MODELOS, "\n")
cat("      Ejecute ahora: equate_fundamentos.R\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")