# ============================================================
# PAARS — Script 4: Boxplots de score.final por Asignatura y Grado
#
# PRE-REQUISITO: haber ejecutado los Scripts 2 y 3, que generan:
#   results/reporte_estudiantes.xlsx
#
# QUÉ HACE ESTE SCRIPT:
#   Lee el reporte de estudiantes y produce dos boxplots:
#     1. score.final (escala 0-100) — Lectura (LEC) por grado
#     2. score.final (escala 0-100) — Matemática (MAT) por grado
#
#   Los boxplots se pueden comparar entre grados porque los puntajes
#   están equiparados a una escala común (media=50, SD=17, referencia
#   2do Grado, método Stocking-Lord).
#
# SALIDA:
#   results/boxplots_score_final.png   ← ambos boxplots en una imagen
#   results/boxplot_LEC.png            ← boxplot LEC individual
#   results/boxplot_MAT.png            ← boxplot MAT individual
#
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(ggplot2)
  library(dplyr)
})

# ============================================================
# CONFIGURACIÓN
# ============================================================
CARPETA_RESULTADOS <- "results"

ARCHIVO_REPORTE <- file.path(CARPETA_RESULTADOS, "reporte_estudiantes.xlsx")

ARCHIVO_SALIDA_COMBINADO  <- file.path(CARPETA_RESULTADOS, "boxplots_score_final.png")
ARCHIVO_SALIDA_LEC        <- file.path(CARPETA_RESULTADOS, "boxplot_LEC.png")
ARCHIVO_SALIDA_MAT        <- file.path(CARPETA_RESULTADOS, "boxplot_MAT.png")

# Colores por grado
COLORES_GRADO <- c(
  "Segundo Grado" = "#1A6FA8",   # azul
  "Tercer Grado"  = "#1A9E6F",   # verde
  "Cuarto Grado"  = "#D4A017"    # dorado
)

# Colores de cabecera por asignatura
COLOR_LEC <- "#033B6D"
COLOR_MAT <- "#085041"

# Etiquetas cortas de grado para los ejes
ETIQUETAS_GRADO <- c(
  "Segundo Grado" = "2do Grado",
  "Tercer Grado"  = "3er Grado",
  "Cuarto Grado"  = "4to Grado"
)

# ============================================================
# VERIFICAR ARCHIVO
# ============================================================
if (!file.exists(ARCHIVO_REPORTE)) {
  stop(sprintf(
    "Archivo no encontrado: %s\nEjecute primero reporte_estudiantes.R (Script 3).",
    ARCHIVO_REPORTE
  ), call. = FALSE)
}

# ============================================================
# LEER DATOS
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  PAARS — Script 4: Boxplots score.final por Grado\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")

cat("  Leyendo", ARCHIVO_REPORTE, "...\n")
df <- as.data.frame(read_excel(ARCHIVO_REPORTE, sheet = "Estudiantes"))

# Verificar que la columna Grado exista
if (!"Grado" %in% names(df)) {
  stop("Columna 'Grado' no encontrada en el archivo.", call. = FALSE)
}

# Verificar que existan columnas de puntaje LEC y MAT
if (length(grep("LEC", names(df))) == 0 || length(grep("MAT", names(df))) == 0) {
  stop(sprintf(
    "No se encontraron columnas de puntaje LEC/MAT. Columnas disponibles:\n  %s",
    paste(names(df), collapse = "\n  ")
  ), call. = FALSE)
}

cat(sprintf("  Estudiantes cargados: %d\n", nrow(df)))
cat(sprintf("  Columnas disponibles: %s\n\n",
            paste(names(df), collapse = ", ")))

# ============================================================
# PREPARAR DATOS
# ============================================================

# Limpiar nombres de columnas: eliminar espacios y saltos de linea
names(df) <- trimws(gsub("\n", " ", names(df)))

# Filtrar fila de nota al pie (Grado = NA)
df <- df[!is.na(df$Grado) & df$Grado != "", ]

# Orden y etiquetas de grado
orden_grados <- c("Segundo Grado", "Tercer Grado", "Cuarto Grado")
df$Grado <- factor(df$Grado, levels = orden_grados)

# Detectar columnas de puntaje aunque tengan espacios u otras variaciones
col_lec <- grep("LEC", names(df), value = TRUE, ignore.case = FALSE)[1]
col_mat <- grep("MAT", names(df), value = TRUE, ignore.case = FALSE)[1]

if (is.na(col_lec) || is.na(col_mat)) {
  stop(sprintf(
    "No se encontraron columnas de puntaje. Columnas disponibles:\n  %s",
    paste(names(df), collapse = "\n  ")
  ), call. = FALSE)
}

cat(sprintf("  Columna LEC detectada: '%s'\n", col_lec))
cat(sprintf("  Columna MAT detectada: '%s'\n\n", col_mat))

names(df)[names(df) == col_lec] <- "score_LEC"
names(df)[names(df) == col_mat] <- "score_MAT"

# Estadísticos de referencia por asignatura y grado
resumen <- df %>%
  group_by(Grado) %>%
  summarise(
    n_LEC    = sum(!is.na(score_LEC)),
    med_LEC  = round(median(score_LEC, na.rm = TRUE), 1),
    mean_LEC = round(mean(score_LEC,   na.rm = TRUE), 1),
    n_MAT    = sum(!is.na(score_MAT)),
    med_MAT  = round(median(score_MAT, na.rm = TRUE), 1),
    mean_MAT = round(mean(score_MAT,   na.rm = TRUE), 1),
    .groups = "drop"
  )

cat("  Estadísticos por grado:\n")
cat(sprintf("  %-16s %6s %8s %8s %8s %8s %8s\n",
            "Grado", "n", "Med.LEC", "Avg.LEC", "Med.MAT", "Avg.MAT", ""))
for (i in seq_len(nrow(resumen))) {
  r <- resumen[i, ]
  cat(sprintf("  %-16s %6d %8.1f %8.1f %8.1f %8.1f\n",
              as.character(r$Grado), r$n_LEC,
              r$med_LEC, r$mean_LEC,
              r$med_MAT, r$mean_MAT))
}
cat("\n")

# ============================================================
# FUNCIÓN: construir_boxplot
# ============================================================
construir_boxplot <- function(data, col_score, titulo, subtitulo,
                              color_titulo, etiqueta_y) {
  
  # Filtrar NAs en el score y en Grado
  data <- data[!is.na(data[[col_score]]) & !is.na(data$Grado), ]
  
  # Etiquetas del eje X con n por grado — base R para evitar
  # problemas de tidy evaluation dentro de funciones
  conteos_n <- table(data$Grado)
  label_map <- sapply(orden_grados, function(g) {
    n_g <- if (g %in% names(conteos_n)) conteos_n[[g]] else 0
    paste0(ETIQUETAS_GRADO[[g]], "\n(n=", n_g, ")")
  })
  names(label_map) <- orden_grados
  
  data$eje_x <- factor(
    label_map[as.character(data$Grado)],
    levels = label_map[orden_grados]
  )
  
  # Columna auxiliar para el fill: copia de Grado como character
  data$grado_fill <- as.character(data$Grado)
  
  ggplot(data, aes(x = eje_x,
                   y = .data[[col_score]],
                   fill = grado_fill)) +
    
    # Línea de referencia: media poblacional = 50
    geom_hline(yintercept = 50, linetype = "dashed",
               colour = "#999999", linewidth = 0.6) +
    annotate("text", x = Inf, y = 51.5,
             label = "Media poblacional (50)  ",
             hjust = 1, size = 3, colour = "#888888",
             fontface = "italic") +
    
    # Boxplot
    geom_boxplot(
      colour      = "#333333",
      linewidth   = 0.5,
      outlier.shape  = 21,
      outlier.size   = 1.8,
      outlier.colour = "#333333",
      outlier.fill   = "white",
      outlier.alpha  = 0.7,
      width       = 0.55,
      notch       = FALSE
    ) +
    
    # Media como punto
    stat_summary(
      fun     = mean,
      geom    = "point",
      shape   = 23,
      size    = 3,
      fill    = "white",
      colour  = "#333333",
      stroke  = 0.8
    ) +
    
    # Etiqueta de mediana sobre cada caja
    stat_summary(
      fun      = median,
      geom     = "text",
      aes(label = round(after_stat(y), 1)),
      vjust    = -0.6,
      size     = 3.2,
      fontface = "bold",
      colour   = "#333333"
    ) +
    
    scale_fill_manual(values = COLORES_GRADO, guide = "none") +
    
    scale_y_continuous(
      limits = c(0, 100),
      breaks = seq(0, 100, by = 10),
      expand = expansion(mult = c(0.01, 0.04))
    ) +
    
    labs(
      title    = titulo,
      subtitle = subtitulo,
      x        = NULL,
      y        = etiqueta_y,
      caption  = paste0(
        "Puntajes equiparados (Stocking-Lord, referencia 2do Grado). ",
        "Escala: media=50, SD=17.\n",
        "◆ = media aritmética  |  línea central = mediana  |",
        "  ○ = valores atípicos"
      )
    ) +
    
    theme_minimal(base_size = 12) +
    theme(
      plot.title      = element_text(face = "bold", size = 14,
                                     colour = color_titulo,
                                     margin = margin(b = 4)),
      plot.subtitle   = element_text(size = 10, colour = "#555555",
                                     margin = margin(b = 10)),
      plot.caption    = element_text(size = 8, colour = "#888888",
                                     hjust = 0,
                                     margin = margin(t = 8)),
      axis.text.x     = element_text(size = 11, face = "bold",
                                     colour = "#333333"),
      axis.text.y     = element_text(size = 10, colour = "#555555"),
      axis.title.y    = element_text(size = 10, colour = "#555555",
                                     margin = margin(r = 8)),
      panel.grid.major.x = element_blank(),
      panel.grid.major.y = element_line(colour = "#EEEEEE",
                                        linewidth = 0.5),
      panel.grid.minor   = element_blank(),
      plot.background    = element_rect(fill = "white", colour = NA),
      panel.background   = element_rect(fill = "white", colour = NA),
      plot.margin        = margin(16, 20, 12, 16)
    )
}

# ============================================================
# BOXPLOT LEC
# ============================================================
cat("  Generando boxplot LEC...\n")

p_lec <- construir_boxplot(
  data         = df,
  col_score    = "score_LEC",
  titulo       = "Lectura (LEC) — Distribución del puntaje final por grado",
  subtitulo    = "score.final (escala 0-100) equiparado a escala común",
  color_titulo = COLOR_LEC,
  etiqueta_y   = "score.final (escala 0-100)"
)

ggsave(ARCHIVO_SALIDA_LEC, plot = p_lec,
       width = 8, height = 6, dpi = 180, bg = "white")
cat(sprintf("  ✓ %s\n", ARCHIVO_SALIDA_LEC))

# ============================================================
# BOXPLOT MAT
# ============================================================
cat("  Generando boxplot MAT...\n")

p_mat <- construir_boxplot(
  data         = df,
  col_score    = "score_MAT",
  titulo       = "Matemática (MAT) — Distribución del puntaje final por grado",
  subtitulo    = "score.final (escala 0-100) equiparado a escala común",
  color_titulo = COLOR_MAT,
  etiqueta_y   = "score.final (escala 0-100)"
)

ggsave(ARCHIVO_SALIDA_MAT, plot = p_mat,
       width = 8, height = 6, dpi = 180, bg = "white")
cat(sprintf("  ✓ %s\n", ARCHIVO_SALIDA_MAT))

# ============================================================
# IMAGEN COMBINADA (LEC + MAT lado a lado)
# ============================================================
cat("  Generando imagen combinada...\n")

# Usar patchwork si está disponible, si no usar cowplot, si no gridExtra
pkg_ok <- suppressWarnings(
  tryCatch({ library(patchwork); "patchwork" }, error = function(e)
    tryCatch({ library(cowplot);   "cowplot"   }, error = function(e)
      tryCatch({ library(gridExtra); "gridExtra" }, error = function(e) "none")))
)

if (pkg_ok == "patchwork") {
  p_combined <- p_lec + p_mat +
    plot_annotation(
      title   = "Distribución del puntaje final equiparado por asignatura y grado",
      caption = "PAARS — Prueba de Fundamentos",
      theme   = theme(
        plot.title = element_text(face = "bold", size = 15,
                                  hjust = 0.5, margin = margin(b = 8)),
        plot.caption = element_text(size = 8, colour = "#999999",
                                    hjust = 1)
      )
    )
  ggsave(ARCHIVO_SALIDA_COMBINADO, plot = p_combined,
         width = 16, height = 6, dpi = 180, bg = "white")
  
} else if (pkg_ok == "cowplot") {
  p_combined <- cowplot::plot_grid(p_lec, p_mat, ncol = 2,
                                   labels = c("A", "B"),
                                   label_size = 13)
  ggsave(ARCHIVO_SALIDA_COMBINADO, plot = p_combined,
         width = 16, height = 6, dpi = 180, bg = "white")
  
} else if (pkg_ok == "gridExtra") {
  png(ARCHIVO_SALIDA_COMBINADO, width = 1600, height = 600, res = 100)
  gridExtra::grid.arrange(p_lec, p_mat, ncol = 2)
  dev.off()
  
} else {
  # Fallback: guardar los dos gráficos en una sola imagen con par()
  png(ARCHIVO_SALIDA_COMBINADO, width = 1600, height = 620, res = 100)
  gridExtra::grid.arrange(p_lec, p_mat, ncol = 2)
  dev.off()
  cat("  [i] Ningún paquete de composición disponible.",
      "Se recomienda instalar 'patchwork'.\n")
}

cat(sprintf("  ✓ %s\n", ARCHIVO_SALIDA_COMBINADO))

# ============================================================
# RESUMEN FINAL
# ============================================================
cat("\n", paste(rep("=", 60), collapse = ""), "\n")
cat("  [✓] Script 4 terminado.\n\n")
cat("  Archivos generados:\n")
cat(sprintf("    %s\n", ARCHIVO_SALIDA_LEC))
cat(sprintf("    %s\n", ARCHIVO_SALIDA_MAT))
cat(sprintf("    %s\n", ARCHIVO_SALIDA_COMBINADO))
cat("\n  Interpretación de los boxplots:\n")
cat("    - La caja abarca el rango intercuartil (Q1 a Q3)\n")
cat("    - La línea dentro de la caja es la MEDIANA\n")
cat("    - El número sobre cada caja es el valor de la mediana\n")
cat("    - El rombo (◆) es la MEDIA aritmética\n")
cat("    - Los círculos son valores atípicos\n")
cat("    - La línea punteada horizontal marca la media\n")
cat("      poblacional combinada (50)\n")
cat("    - Los grados son comparables porque los puntajes\n")
cat("      están equiparados a una escala común\n")
cat(paste(rep("=", 60), collapse = ""), "\n\n")