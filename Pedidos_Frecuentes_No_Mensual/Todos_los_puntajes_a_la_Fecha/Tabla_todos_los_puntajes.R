# ─────────────────────────────────────────────
# LIBRARIES
# ─────────────────────────────────────────────
library(dplyr)
library(readr)
library(stringr)
library(tidyr)
library(purrr)
library(openxlsx)

# ─────────────────────────────────────────────
# 1. LOAD & FILTER DATA
# ─────────────────────────────────────────────
INPUT_FILE  <- "C:/Users/USUARIO.DESKTOP-KJ9DUHE.000/Documents/PAARS_System/Pedidos_Frecuentes_No_Mensual/Criterio_Inclusión_Exclusión_Escuelas/Excel_Maestro_HojaUnica.csv"
OUTPUT_B1   <- "C:/Users/USUARIO.DESKTOP-KJ9DUHE.000/Documents/PAARS_System/Pedidos_Frecuentes_No_Mensual/Criterio_Inclusión_Exclusión_Escuelas/Centros_Escolares_B1.xlsx"
OUTPUT_B2   <- "C:/Users/USUARIO.DESKTOP-KJ9DUHE.000/Documents/PAARS_System/Pedidos_Frecuentes_No_Mensual/Criterio_Inclusión_Exclusión_Escuelas/Centros_Escolares_B2.xlsx"

# Read the CSV 
df <- read_delim(INPUT_FILE, delim = ";", show_col_types = FALSE)

# ── EXCLUDE VIRTUAL SCHOOL (99999) ──────────────────────────────────
if ("CODIGO" %in% names(df)) {
  df <- df %>% filter(as.character(CODIGO) != "99999")
}
if ("Nro de centro" %in% names(df)) {
  df <- df %>% filter(as.character(`Nro de centro`) != "99999")
}
# ────────────────────────────────────────────────────────────────────

# ── Determine the definitive group for each school ──────────────────
school_group <- df %>%
  count(CODIGO, GRUPO_Limpio) %>%
  group_by(CODIGO) %>%
  slice_max(n, n = 1, with_ties = FALSE) %>%
  ungroup() %>%
  select(CODIGO, GRUPO_Escuela = GRUPO_Limpio)

df <- df %>% left_join(school_group, by = "CODIGO")

valid_b1 <- df %>% filter(GRUPO_Escuela == "B1")
valid_b2 <- df %>% filter(GRUPO_Escuela == "B2")

# ─────────────────────────────────────────────
# 2. SCHOOL TYPE CLASSIFICATION
# ─────────────────────────────────────────────
classify_school <- function(name) {
  n <- toupper(as.character(name))
  case_when(
    str_detect(n, "INSTITUTO") ~ "Instituto",
    str_detect(n, "COMPLEJO EDUCATIVO") ~ "Complejo Educativo",
    TRUE ~ "Centro Escolar"
  )
}

valid_b1 <- valid_b1 %>% mutate(Tipo_Escuela = classify_school(NOMBRE))
valid_b2 <- valid_b2 %>% mutate(Tipo_Escuela = classify_school(NOMBRE))

# ─────────────────────────────────────────────
# 3. CONSTANTS
# ─────────────────────────────────────────────
NIVEL_COLS_B1 <- list(
  mat = c("Nivel Matemática (Mes 1 (Marzo))", "Nivel Matemática (Mes 2 (Abril))", "Nivel Matemática (Mes 3 (Mayo))"),
  len = c("Nivel Lengua (Mes 1 (Marzo))", "Nivel Lengua (Mes 2 (Abril))", "Nivel Lengua (Mes 3 (Mayo))")
)
SCORE_COLS_B1 <- list(
  mat = c("Puntaje Matemática (Mes 1 (Marzo))", "Puntaje Matemática (Mes 2 (Abril))", "Puntaje Matemática (Mes 3 (Mayo))"),
  len = c("Puntaje Lengua (Mes 1 (Marzo))", "Puntaje Lengua (Mes 2 (Abril))", "Puntaje Lengua (Mes 3 (Mayo))")
)

NIVEL_COLS_B2 <- list(
  mat = c("Nivel Matemática (Mes 3 (Mayo))"),
  len = c("Nivel Lengua (Mes 3 (Mayo))")
)
SCORE_COLS_B2 <- list(
  mat = c("Puntaje Matemática (Mes 3 (Mayo))"),
  len = c("Puntaje Lengua (Mes 3 (Mayo))")
)

GRADE_ORDER <- c("2º", "3º", "4º", "5º", "6º", "7°", "8°", "9°", "10°", "11°")
STATUS_ORDER <- c("Alerta", "Regular", "Bueno", "Excelente")

# ─────────────────────────────────────────────
# 4. CRITICAL % CALCULATIONS
# ─────────────────────────────────────────────
pct_critico <- function(vec) {
  valid <- na.omit(vec)
  if (length(valid) == 0) return(NA_real_)
  sum(valid == "Crítico", na.rm = TRUE) / length(valid)
}

compute_status_b1 <- function(mat_wins, len_wins) {
  if (mat_wins >= 2 && len_wins >= 2) return("Excelente")
  if ((mat_wins == 3 && len_wins == 1) || (mat_wins == 1 && len_wins == 3)) return("Bueno")
  if ((mat_wins == 2 && len_wins == 1) || (mat_wins == 1 && len_wins == 2)) return("Regular")
  return("Alerta")
}

compute_status_b2 <- function(mat_wins, len_wins) {
  if (mat_wins == 1 && len_wins == 1) return("Excelente")
  if (mat_wins == 1 || len_wins == 1) return("Regular")
  return("Alerta")
}

# ─────────────────────────────────────────────
# 5. BUILD UNIVERSE % CRÍTICO
# ─────────────────────────────────────────────
build_universe <- function(data, nivel_cols) {
  types <- c("Instituto", "Complejo Educativo", "Centro Escolar")
  universe <- list()
  for (tipo in types) {
    subset_data <- data %>% filter(Tipo_Escuela == tipo)
    univ_type <- list()
    for (subj in names(nivel_cols)) {
      univ_type[[subj]] <- map_dbl(nivel_cols[[subj]], ~pct_critico(subset_data[[.x]]))
    }
    universe[[tipo]] <- univ_type
  }
  return(universe)
}

universes_b1 <- build_universe(valid_b1, NIVEL_COLS_B1)
universes_b2 <- build_universe(valid_b2, NIVEL_COLS_B2)

# ─────────────────────────────────────────────
# 6. COMPUTE SCHOOL-LEVEL STATUS
# ─────────────────────────────────────────────
count_participations <- function(grp, nivel_cols) {
  participated <- 0
  for (i in seq_along(nivel_cols$mat)) {
    has_mat <- any(!is.na(grp[[ nivel_cols$mat[i] ]]))
    has_len <- any(!is.na(grp[[ nivel_cols$len[i] ]]))
    if (has_mat || has_len) participated <- participated + 1
  }
  return(participated)
}

build_school_records <- function(valid_data, nivel_cols, universes, compute_status_fn, include_participations = FALSE) {
  school_list <- list()
  split_data <- split(valid_data, list(valid_data$CODIGO, valid_data$NOMBRE), drop = TRUE)
  
  for (grp in split_data) {
    codigo <- grp$CODIGO[1]
    nombre <- grp$NOMBRE[1]
    tipo <- grp$Tipo_Escuela[1]
    univ <- universes[[tipo]]
    
    wins <- list(mat = 0, len = 0)
    for (subj in c("mat", "len")) {
      for (i in seq_along(nivel_cols[[subj]])) {
        col <- nivel_cols[[subj]][i]
        s_pct <- pct_critico(grp[[col]])
        u_pct <- univ[[subj]][i]
        if (!is.na(s_pct) && !is.na(u_pct) && s_pct <= u_pct) {
          wins[[subj]] <- wins[[subj]] + 1
        }
      }
    }
    
    estatus <- compute_status_fn(wins$mat, wins$len)
    record <- data.frame(
      Código = codigo, `Centro Escolar` = nombre, Tipo = tipo, Estatus = estatus,
      mat_wins = wins$mat, len_wins = wins$len, stringsAsFactors = FALSE, check.names = FALSE
    )
    
    if (include_participations) record$`N° Aplicaciones` <- count_participations(grp, nivel_cols)
    
    for (grade in GRADE_ORDER) {
      grade_grp <- grp %>% filter(Grado_Str == grade)
      if (nrow(grade_grp) == 0) {
        record[[paste0("grado_", grade)]] <- ""
        next
      }
      grade_univ_data <- valid_data %>% filter(Tipo_Escuela == tipo, Grado_Str == grade)
      g_wins <- list(mat = 0, len = 0)
      for (subj in c("mat", "len")) {
        for (col in nivel_cols[[subj]]) {
          s_pct <- pct_critico(grade_grp[[col]])
          u_pct <- pct_critico(grade_univ_data[[col]])
          if (!is.na(s_pct) && !is.na(u_pct) && s_pct <= u_pct) {
            g_wins[[subj]] <- g_wins[[subj]] + 1
          }
        }
      }
      record[[paste0("grado_", grade)]] <- compute_status_fn(g_wins$mat, g_wins$len)
    }
    school_list[[length(school_list) + 1]] <- record
  }
  
  df_out <- bind_rows(school_list) %>%
    mutate(status_sort = match(Estatus, STATUS_ORDER)) %>%
    arrange(status_sort) %>% select(-status_sort)
  return(df_out)
}

schools_b1 <- build_school_records(valid_b1, NIVEL_COLS_B1, universes_b1, compute_status_b1, TRUE)
schools_b2 <- build_school_records(valid_b2, NIVEL_COLS_B2, universes_b2, compute_status_b2, FALSE)

# ─────────────────────────────────────────────
# 6b. BUILD VERIFICATION DATA
# ─────────────────────────────────────────────
build_verif_df <- function(valid_data, nivel_cols, universes, schools_df, month_labels, max_wins) {
  subj_labels <- c(mat = "Matemática", len = "Lenguaje")
  verif_list <- list()
  
  split_data <- split(valid_data, list(valid_data$CODIGO, valid_data$NOMBRE), drop = TRUE)
  for (grp in split_data) {
    codigo <- grp$CODIGO[1]
    nombre <- grp$NOMBRE[1]
    tipo <- grp$Tipo_Escuela[1]
    univ <- universes[[tipo]]
    
    sch_row <- schools_df %>% filter(Código == codigo)
    
    for (subj in c("mat", "len")) {
      for (i in seq_along(nivel_cols[[subj]])) {
        col <- nivel_cols[[subj]][i]
        s_pct <- pct_critico(grp[[col]])
        u_pct <- univ[[subj]][i]
        
        win_val <- "N/D"
        if (!is.na(s_pct) && !is.na(u_pct)) {
          win_val <- if_else(s_pct <= u_pct, "✔ Sí", "✘ No")
        }
        
        verif_list[[length(verif_list) + 1]] <- data.frame(
          Código = codigo, `Centro Escolar` = nombre, Tipo = tipo, Estatus = sch_row$Estatus[1],
          `Mat Wins` = paste0(sch_row$mat_wins[1], "/", max_wins),
          `Len Wins` = paste0(sch_row$len_wins[1], "/", max_wins),
          Asignatura = subj_labels[[subj]], Mes = month_labels[i],
          `% Crítico Escuela` = if(!is.na(s_pct)) round(s_pct * 100, 2) else NA,
          `% Crítico Universo (Ref.)` = if(!is.na(u_pct)) round(u_pct * 100, 2) else NA,
          `Escuela ≤ Universo` = win_val, stringsAsFactors = FALSE, check.names = FALSE
        )
      }
    }
  }
  
  bind_rows(verif_list) %>%
    mutate(status_sort = match(Estatus, STATUS_ORDER),
           subj_sort = match(Asignatura, c("Matemática", "Lenguaje")),
           mes_sort = match(Mes, month_labels)) %>%
    arrange(status_sort, `Centro Escolar`, subj_sort, mes_sort) %>%
    select(-status_sort, -subj_sort, -mes_sort)
}

MONTH_LABELS_B1 <- c('Mes 1 (Marzo)', 'Mes 2 (Abril)', 'Mes 3 (Mayo)')
MONTH_LABELS_B2 <- c('Mes 3 (Mayo)')

verif_b1 <- build_verif_df(valid_b1, NIVEL_COLS_B1, universes_b1, schools_b1, MONTH_LABELS_B1, "3")
verif_b2 <- build_verif_df(valid_b2, NIVEL_COLS_B2, universes_b2, schools_b2, MONTH_LABELS_B2, "1")

# ─────────────────────────────────────────────
# 7. BUILD STUDENTS SHEET DATA
# ─────────────────────────────────────────────
build_students_df <- function(valid_data, schools_df, score_cols_dict) {
  all_score_cols <- character()
  for (m_col in score_cols_dict$mat) {
    all_score_cols <- c(all_score_cols, m_col, str_replace(m_col, "Puntaje", "Nivel"))
  }
  for (l_col in score_cols_dict$len) {
    all_score_cols <- c(all_score_cols, l_col, str_replace(l_col, "Puntaje", "Nivel"))
  }
  
  valid_cols <- intersect(all_score_cols, names(valid_data))
  
  # Standardize base column names 
  base_cols <- c("CODIGO", "NOMBRE", "Tipo_Escuela", "NIE", "PRIMER_NOMBRE", 
                 "SEGUNDO_NOMBRE", "PRIMER_APELLIDO", "SEGUNDO_APELLIDO", "Grado_Str")
  if ("CÓDIGO_SECCIÓN" %in% names(valid_data)) base_cols <- c(base_cols, "CÓDIGO_SECCIÓN")
  
  students <- valid_data %>% select(all_of(c(base_cols, valid_cols))) %>%
    rename(
      Código = CODIGO, `Centro Escolar` = NOMBRE, Tipo = Tipo_Escuela, NIE = NIE,
      `Primer Nombre` = PRIMER_NOMBRE, `Segundo Nombre` = SEGUNDO_NOMBRE,
      `Primer Apellido` = PRIMER_APELLIDO, `Segundo Apellido` = SEGUNDO_APELLIDO, Grado = Grado_Str
    )
  
  if ("CÓDIGO_SECCIÓN" %in% names(valid_data)) {
    students <- students %>% rename(`Código Sección LXP` = CÓDIGO_SECCIÓN)
  } else {
    students <- students %>% mutate(`Código Sección LXP` = NA_character_)
  }
  
  grade_cols <- paste0("grado_", GRADE_ORDER)
  students <- students %>% left_join(schools_df %>% select(Código, Estatus, any_of(grade_cols)), by = "Código")
  
  # Extract specific grade status row by row
  students$`Estatus Grado` <- NA_character_
  for (i in 1:nrow(students)) {
    col_name <- paste0("grado_", students$Grado[i])
    if (col_name %in% names(students)) students$`Estatus Grado`[i] <- students[[col_name]][i]
  }
  
  students <- students %>%
    rename(`Estatus Centro Escolar` = Estatus) %>%
    mutate(status_sort = match(`Estatus Centro Escolar`, STATUS_ORDER),
           grade_sort = match(Grado, GRADE_ORDER)) %>%
    arrange(status_sort, `Centro Escolar`, grade_sort) %>%
    select(-status_sort, -grade_sort, -any_of(grade_cols))
  
  # Reorder
  col_order <- c("Código", "Centro Escolar", "Tipo", "NIE", "Primer Nombre", "Segundo Nombre",
                 "Primer Apellido", "Segundo Apellido", "Grado", "Estatus Centro Escolar", 
                 "Estatus Grado", "Código Sección LXP", valid_cols)
  
  list(df = students %>% select(all_of(col_order)), cols = valid_cols)
}

st_b1_info <- build_students_df(valid_b1, schools_b1, SCORE_COLS_B1)
students_b1 <- st_b1_info$df
score_cols_b1 <- st_b1_info$cols

st_b2_info <- build_students_df(valid_b2, schools_b2, SCORE_COLS_B2)
students_b2 <- st_b2_info$df
score_cols_b2 <- st_b2_info$cols


# ─────────────────────────────────────────────
# 8. EXCEL EXPORT WITH OPENXLSX (ALL 3 SHEETS + COLORS)
# ─────────────────────────────────────────────
st_header <- createStyle(fontColour = "#FFFFFF", fgFill = "#2E4057", halign = "center", valign = "center", textDecoration = "bold", border = "TopBottomLeftRight")
st_body_alt <- createStyle(fgFill = "#F2F2F2", border = "TopBottomLeftRight", halign = "center")
st_body_white <- createStyle(fgFill = "#FFFFFF", border = "TopBottomLeftRight", halign = "center")

status_styles <- list(
  Excelente = createStyle(fgFill = "#1E7145", fontColour = "#FFFFFF", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight"),
  Bueno     = createStyle(fgFill = "#70AD47", fontColour = "#FFFFFF", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight"),
  Regular   = createStyle(fgFill = "#FFD966", fontColour = "#000000", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight"),
  Alerta    = createStyle(fgFill = "#FF4B4B", fontColour = "#FFFFFF", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight")
)

nivel_styles <- list(
  "Crítico"   = createStyle(fgFill = "#FF4B4B", fontColour = "#FFFFFF", halign = "center", border = "TopBottomLeftRight"),
  "Bajo"      = createStyle(fgFill = "#FF9966", fontColour = "#FFFFFF", halign = "center", border = "TopBottomLeftRight"),
  "Medio"     = createStyle(fgFill = "#FFD966", fontColour = "#000000", halign = "center", border = "TopBottomLeftRight"),
  "Bueno"     = createStyle(fgFill = "#A9D18E", fontColour = "#000000", halign = "center", border = "TopBottomLeftRight"),
  "Excelente" = createStyle(fgFill = "#1E7145", fontColour = "#FFFFFF", halign = "center", border = "TopBottomLeftRight")
)

win_styles <- list(
  "✔ Sí" = createStyle(fgFill = "#C6EFCE", fontColour = "#276221", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight"),
  "✘ No" = createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006", textDecoration = "bold", halign = "center", border = "TopBottomLeftRight"),
  "N/D"  = createStyle(fgFill = "#EEEEEE", fontColour = "#888888", halign = "center", border = "TopBottomLeftRight")
)

build_workbook <- function(students_df, schools_df, verif_df, output_path) {
  wb <- createWorkbook()
  
  # ── SHEET 1: ESTUDIANTES ───────────────────
  addWorksheet(wb, "Estudiantes")
  freezePane(wb, "Estudiantes", firstActiveRow = 2)
  writeData(wb, "Estudiantes", students_df, startRow = 1, colNames = TRUE)
  addStyle(wb, "Estudiantes", st_header, rows = 1, cols = 1:ncol(students_df), gridExpand = TRUE)
  
  for (r in 1:nrow(students_df)) {
    # Apply status colors to 'Estatus Centro Escolar' & 'Estatus Grado'
    for (col_name in c("Estatus Centro Escolar", "Estatus Grado")) {
      col_idx <- which(names(students_df) == col_name)
      val <- students_df[[col_idx]][r]
      if (!is.na(val) && val %in% names(status_styles)) {
        addStyle(wb, "Estudiantes", status_styles[[val]], rows = r + 1, cols = col_idx)
      }
    }
    # Apply Nivel categorization colors to Nivel columns
    nivel_cols <- which(str_detect(names(students_df), "Nivel"))
    for (col_idx in nivel_cols) {
      val <- students_df[[col_idx]][r]
      if (!is.na(val) && val %in% names(nivel_styles)) {
        addStyle(wb, "Estudiantes", nivel_styles[[val]], rows = r + 1, cols = col_idx)
      }
    }
  }
  
  # ── SHEET 2: CENTROS ESCOLARES ─────────────
  addWorksheet(wb, "Centros Escolares")
  freezePane(wb, "Centros Escolares", firstActiveRow = 2)
  writeData(wb, "Centros Escolares", schools_df, startRow = 1, colNames = TRUE)
  addStyle(wb, "Centros Escolares", st_header, rows = 1, cols = 1:ncol(schools_df), gridExpand = TRUE)
  
  estatus_cols <- c(which(names(schools_df) == "Estatus"), which(str_detect(names(schools_df), "grado_")))
  for (col in estatus_cols) {
    for (r in 1:nrow(schools_df)) {
      val <- schools_df[[col]][r]
      if (!is.na(val) && val %in% names(status_styles)) {
        addStyle(wb, "Centros Escolares", status_styles[[val]], rows = r + 1, cols = col)
      }
    }
  }
  
  # ── SHEET 3: VERIFICACIÓN ──────────────────
  addWorksheet(wb, "Verificación de Cálculo")
  freezePane(wb, "Verificación de Cálculo", firstActiveRow = 2)
  writeData(wb, "Verificación de Cálculo", verif_df, startRow = 1, colNames = TRUE)
  addStyle(wb, "Verificación de Cálculo", st_header, rows = 1, cols = 1:ncol(verif_df), gridExpand = TRUE)
  
  win_col <- which(names(verif_df) == "Escuela ≤ Universo")
  for (r in 1:nrow(verif_df)) {
    val <- verif_df[[win_col]][r]
    if (!is.na(val) && val %in% names(win_styles)) {
      addStyle(wb, "Verificación de Cálculo", win_styles[[val]], rows = r + 1, cols = win_col)
    }
  }
  
  # Adjust column widths dynamically for all sheets
  for(sheet in c("Estudiantes", "Centros Escolares", "Verificación de Cálculo")) {
    setColWidths(wb, sheet, cols = 1:ncol(get(ifelse(sheet=="Estudiantes", "students_df", ifelse(sheet=="Centros Escolares", "schools_df", "verif_df")))), widths = "auto")
  }
  
  saveWorkbook(wb, output_path, overwrite = TRUE)
  cat("Guardado:", output_path, "\n")
}

# ─────────────────────────────────────────────
# 9. EXECUTE AND SAVE
# ─────────────────────────────────────────────
cat("\nGenerando archivo B1...\n")
build_workbook(students_b1, schools_b1, verif_b1, OUTPUT_B1)

cat("\nGenerando archivo B2...\n")
build_workbook(students_b2, schools_b2, verif_b2, OUTPUT_B2)

cat("\nProceso finalizado con éxito. Revisa la carpeta de salida.\n")