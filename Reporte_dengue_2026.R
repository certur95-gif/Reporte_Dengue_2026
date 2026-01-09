# ========================================================== #
# REPORTE SEMANAL DENGUE - NETLABv2 (CORREGIDO POR COLECCIÓN)
# - Gráfico 1: SE según Fecha de Colección
# - Gráfico 2: Acumulado 2026 según Fecha de Verificación
# - Tabla 03: Positivos SE 2026 según Colección
# - Tabla 04: Pruebas SE actual según Colección
# ========================================================== #

rm(list = ls())

# --------------------------- #
# 0) PAQUETES
# --------------------------- #
req_pkgs <- c(
  "readxl", "dplyr", "stringr", "lubridate", "tidyr", "ggplot2",
  "patchwork", "gridExtra", "scales", "openxlsx", "tibble",
  "officer", "flextable", "webshot2"
)

# --------------------------- #
# 1) CONFIGURACIÓN (EDITA SOLO ESTO)
# --------------------------- #

# Archivo / hoja
archivo <- "dengue_2025.xlsx"
hoja <- 1

# Exámenes permitidos (SOLO estos dos)
examenes_permitidos <- c(
  "Virus Dengue Ag NS1 [Presencia] en Suero o Plasma por Inmunoensayo",
  "Virus Dengue Ac. IgM [Presencia] en Suero o Plasma por Inmunoensayo"
)

# Unidad global
unidad_global <- "examen"

# Filtro opcional por Laboratorio Destino (GLOBAL):
lab_destino_global <- "TODOS"

# Semana epidemiológica: MMWR (Domingo–Sábado)
week_system <- "MMWR"

# Carpeta del reporte
incluir_anio_en_carpeta <- FALSE

# ------------- GRÁFICO 1 (IP% por SE - SEGÚN COLECCIÓN) -------------
g1_anio <- "AUTO"
g1_se_inicio <- "20"
g1_unidad <- unidad_global

# ------------- GRÁFICO 2 (Procesamiento acumulado 2026 - SEGÚN VERIFICACIÓN) -------------
g2_anio <- 2026  # FIJO: año actual según tu solicitud
g2_se_inicio <- 1  # Desde SE 1 del año 2026
g2_se_fin <- "AUTO"  # Hasta SE actual según verificación
g2_unidad <- "examen"
g2_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")
g2_lab_solo <- "TODOS"

# ------------- TABLA POSITIVOS por PROVINCIA (SEGÚN COLECCIÓN 2026) -------------
tabprov_anio <- 2026  # FIJO: solo año 2026 según tu solicitud
tabprov_se_inicio <- 1  # Desde SE 1
tabprov_se_fin <- "AUTO"  # Hasta SE actual
tabprov_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")
tabprov_lab_solo <- "TODOS"
tabprov_unidad <- "examen"

# ------------- TABLA POR SE (MicroRED x EE.SS. Origen - SEGÚN COLECCIÓN) -------------
tabse_anio <- "AUTO"  # Año actual según colección
tabse_se <- "AUTO"  # SE actual según colección (debe coincidir con Gráfico 1)
tabse_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")
tabse_lab_solo <- "TODOS"
tabse_unidad <- "examen"

# --------------------------- #
# 2) FUNCIONES (NO EDITAR)
# --------------------------- #

load_packages <- function(packages) {
  to_install <- packages[!sapply(packages, requireNamespace, quietly = TRUE)]
  if (length(to_install) > 0) {
    install.packages(to_install)
  }
  invisible(lapply(packages, library, character.only = TRUE))
  message("Librerías cargadas correctamente.")
}

validate_inputs <- function(archivo, hoja, week_system, unidad_global, unidades_validas) {
  if (!file.exists(archivo)) {
    stop("No se encontró el archivo de entrada: ", archivo)
  }
  sheets <- readxl::excel_sheets(archivo)
  if (is.numeric(hoja)) {
    if (hoja < 1 || hoja > length(sheets)) {
      stop("La hoja indicada está fuera de rango. Hojas disponibles: ", paste(sheets, collapse = ", "))
    }
  } else if (!hoja %in% sheets) {
    stop("No se encontró la hoja '", hoja, "'. Hojas disponibles: ", paste(sheets, collapse = ", "))
  }
  if (!week_system %in% c("MMWR", "ISO")) {
    stop("week_system debe ser 'MMWR' o 'ISO'. Valor recibido: ", week_system)
  }
  if (!unidad_global %in% unidades_validas) {
    stop("unidad_global debe ser 'examen' o 'muestra'. Valor recibido: ", unidad_global)
  }
}

norm_txt <- function(x) {
  x <- as.character(x)
  x <- iconv(x, from = "", to = "ASCII//TRANSLIT")
  x <- tolower(x)
  stringr::str_squish(x)
}

find_col <- function(df, patterns) {
  nms <- norm_txt(names(df))
  pats <- norm_txt(patterns)
  for (p in pats) {
    idx <- which(stringr::str_detect(nms, fixed(p)))
    if (length(idx) > 0) return(names(df)[idx[1]])
  }
  NA_character_
}

must_col <- function(df, patterns, label = "columna") {
  col <- find_col(df, patterns)
  if (is.na(col)) {
    stop("No se encontró ", label, ". Patrones buscados: ",
         paste(patterns, collapse = " | "),
         "\nSugerencia: revisa names(raw).")
  }
  col
}

parse_excel_date <- function(x) {
  if (inherits(x, "Date")) return(as.Date(x))
  if (inherits(x, "POSIXct") || inherits(x, "POSIXt")) return(as.Date(x))
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))
  as.Date(lubridate::parse_date_time(
    as.character(x),
    orders = c(
      "dmy", "dmY", "ymd", "Ymd", "d/m/Y", "Y-m-d", "d-m-Y", "Y/m/d",
      "ymd HMS", "Ymd HMS", "dmy HMS", "dmY HMS", "ymd HM", "Ymd HM", "dmy HM", "dmY HM",
      "d/m/Y HMS", "d/m/Y HM", "d-m-Y HMS", "d-m-Y HM", "Y/m/d HMS", "Y/m/d HM", "Y-m-d HMS", "Y-m-d HM"
    ),
    tz = "UTC"
  ))
}

add_epi <- function(df, date_col, week_system = c("MMWR", "ISO")) {
  week_system <- match.arg(week_system)
  if (week_system == "ISO") {
    df %>% mutate(se = lubridate::isoweek(.data[[date_col]]),
                  anio = lubridate::isoyear(.data[[date_col]]))
  } else {
    df %>% mutate(se = lubridate::epiweek(.data[[date_col]]),
                  anio = lubridate::epiyear(.data[[date_col]]))
  }
}

apply_filters <- function(df, anio_sel = NULL, se_inicio = NULL, se_fin = NULL, labs_excluir = NULL, labs_solo = "TODOS") {
  out <- df
  if (!is.null(labs_excluir) && length(labs_excluir) > 0) {
    ex <- stringr::str_to_upper(stringr::str_squish(labs_excluir))
    out <- out %>% dplyr::filter(!(.data$lab_destino_std %in% ex))
  }
  if (!(length(labs_solo) == 1 && toupper(labs_solo) == "TODOS")) {
    inc <- stringr::str_to_upper(stringr::str_squish(labs_solo))
    out <- out %>% dplyr::filter(.data$lab_destino_std %in% inc)
  }
  if (!is.null(anio_sel)) {
    out <- out %>% dplyr::filter(.data$anio %in% anio_sel)
  }
  if (!is.null(se_inicio)) out <- out %>% dplyr::filter(.data$se >= se_inicio)
  if (!is.null(se_fin)) out <- out %>% dplyr::filter(.data$se <= se_fin)
  out
}

dedup_by_unit <- function(df, unit = c("examen", "muestra")) {
  unit <- match.arg(unit)
  if (unit == "examen") return(df)
  df %>%
    group_by(cod_muestra, anio, se) %>%
    summarise(
      clasif = case_when(
        any(clasif == "POSITIVO") ~ "POSITIVO",
        all(clasif == "NEGATIVO") ~ "NEGATIVO",
        any(clasif == "INDETERMINADO") ~ "INDETERMINADO",
        TRUE ~ "OTRO"
      ),
      .groups = "drop"
    )
}

style_xlsx <- function(df, path, pct_cols = character(), total_row_idx = NA_integer_) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Reporte")
  openxlsx::writeData(wb, "Reporte", df, withFilter = TRUE)
  headerStyle <- openxlsx::createStyle(
    fontSize = 11, fontColour = "#000000", textDecoration = "bold",
    fgFill = "#D9E1F2", halign = "center", valign = "center", border = "Bottom"
  )
  openxlsx::addStyle(wb, "Reporte", headerStyle, rows = 1, cols = 1:ncol(df), gridExpand = TRUE)
  openxlsx::freezePane(wb, "Reporte", firstRow = TRUE)
  openxlsx::setColWidths(wb, "Reporte", cols = 1:ncol(df), widths = "auto")
  if (length(pct_cols) > 0) {
    pctStyle <- openxlsx::createStyle(numFmt = "0.00\"%\"")
    for (cn in pct_cols) {
      if (cn %in% names(df)) {
        col_i <- which(names(df) == cn)
        openxlsx::addStyle(wb, "Reporte", pctStyle, rows = 2:(nrow(df) + 1), cols = col_i, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  if (!is.na(total_row_idx)) {
    totalStyle <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9E1F2", border = "Top")
    openxlsx::addStyle(wb, "Reporte", totalStyle, rows = total_row_idx + 1, cols = 1:ncol(df), gridExpand = TRUE, stack = TRUE)
  }
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  path
}

resolve_auto <- function(x, auto_value) {
  if (is.character(x) && length(x) == 1 && toupper(x) == "AUTO") return(auto_value)
  x
}

# --------------------------- #
# 3) EJECUCIÓN PRINCIPAL
# --------------------------- #

load_packages(req_pkgs)
validate_inputs(archivo, hoja, week_system, unidad_global, c("examen", "muestra"))

# Leer datos
raw <- readxl::read_excel(archivo, sheet = hoja)

# Detectar columnas
detect_columns <- function(raw) {
  col_fecha <- must_col(raw, c("Fecha Colección", "Fecha Coleccion"), "Fecha de colección")
  col_fecha_verif <- must_col(raw, c("Fecha Verificación", "Fecha Verificacion"), "Fecha de verificación")
  col_result <- must_col(raw, c("resultado"), "Resultado")
  col_examen <- must_col(raw, c("nombre de examen", "examen"), "Examen")
  col_estatus <- must_col(raw, c("estatus resultado", "estado resultado", "estatus"), "Estatus")
  col_cod <- must_col(raw, c("codigo de muestra", "codigo muestra", "cod muestra", "muestra"), "Código de muestra")
  col_labdest <- must_col(raw, c("laboratorio destino", "lab destino", "destino"), "Laboratorio destino")
  col_prov <- find_col(raw, c("provincia", "provincia procedencia", "provincia de procedencia", "provincia domicilio"))
  col_micro <- find_col(raw, c("Micro Red EE.SS Origen", "Micro Red", "MicroRED", "Microred"))
  col_estab <- find_col(raw, c("Establecimiento de Origen", "IPRESS Origen", "Origen"))
  list(
    col_fecha = col_fecha,
    col_fecha_verif = col_fecha_verif,
    col_result = col_result,
    col_examen = col_examen,
    col_estatus = col_estatus,
    col_cod = col_cod,
    col_labdest = col_labdest,
    col_prov = col_prov,
    col_micro = col_micro,
    col_estab = col_estab
  )
}

cols <- detect_columns(raw)

# Construir dataset filtrado
build_dataset <- function(raw, cols, examenes_permitidos, lab_destino_global) {
  examenes_key <- norm_txt(examenes_permitidos)
  dat <- raw %>%
    mutate(
      fecha_coleccion = parse_excel_date(.data[[cols$col_fecha]]),
      fecha_verificacion = parse_excel_date(.data[[cols$col_fecha_verif]]),
      resultado_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_result]]))),
      examen_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_examen]]))),
      examen_key = norm_txt(examen_std),
      estatus_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_estatus]]))),
      cod_muestra = stringr::str_squish(as.character(.data[[cols$col_cod]])),
      lab_destino_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_labdest]])))
    ) %>%
    filter(!is.na(fecha_coleccion)) %>%
    filter(!is.na(fecha_verificacion)) %>%
    filter(estatus_std == "RESULTADO VERIFICADO") %>%
    filter(examen_key %in% examenes_key) %>%
    mutate(
      clasif = case_when(
        str_detect(resultado_std, "POSITIV") ~ "POSITIVO",
        str_detect(resultado_std, "NEGATIV") ~ "NEGATIVO",
        str_detect(resultado_std, "INDETERMIN") ~ "INDETERMINADO",
        TRUE ~ "OTRO"
      )
    )
  if (!(length(lab_destino_global) == 1 && toupper(lab_destino_global) == "TODOS")) {
    lab_ok <- stringr::str_to_upper(stringr::str_squish(lab_destino_global))
    dat <- dat %>% filter(lab_destino_std %in% lab_ok)
  }
  if (nrow(dat) == 0) {
    stop("No quedan registros tras filtros: verificado + exámenes permitidos + lab destino (si aplica).")
  }
  dat
}

dat <- build_dataset(raw, cols, examenes_permitidos, lab_destino_global)

# Agregar SE según Fecha de Colección (para Gráfico 1 y Tablas)
dat <- add_epi(dat, "fecha_coleccion", week_system = week_system) %>%
  rename(se_colec = se, anio_colec = anio)

# Agregar SE según Fecha de Verificación (para Gráfico 2)
dat <- dat %>%
  mutate(
    se_verif = if (week_system == "ISO") lubridate::isoweek(fecha_verificacion) else lubridate::epiweek(fecha_verificacion),
    anio_verif = if (week_system == "ISO") lubridate::isoyear(fecha_verificacion) else lubridate::epiyear(fecha_verificacion)
  )

# Configurar contexto del reporte (según Colección)
fecha_max_colec <- max(dat$fecha_coleccion, na.rm = TRUE)
se_reporte_colec <- if (week_system == "ISO") lubridate::isoweek(fecha_max_colec) else lubridate::epiweek(fecha_max_colec)
anio_rep_colec <- if (week_system == "ISO") lubridate::isoyear(fecha_max_colec) else lubridate::epiyear(fecha_max_colec)

# Para Gráfico 2 (según Verificación)
fecha_max_verif <- max(dat$fecha_verificacion, na.rm = TRUE)
se_reporte_verif <- if (week_system == "ISO") lubridate::isoweek(fecha_max_verif) else lubridate::epiweek(fecha_max_verif)
anio_rep_verif <- if (week_system == "ISO") lubridate::isoyear(fecha_max_verif) else lubridate::epiyear(fecha_max_verif)

# Crear carpeta de salida
carpeta <- if (incluir_anio_en_carpeta) sprintf("%d_SE %02d", anio_rep_colec, se_reporte_colec) else sprintf("SE %02d", se_reporte_colec)
dir.create(carpeta, showWarnings = FALSE, recursive = TRUE)

# Resolver AUTO en configuraciones
g1_anio <- resolve_auto(g1_anio, anio_rep_colec)
g1_se_inicio <- resolve_auto(g1_se_inicio, 1)

g2_se_fin <- resolve_auto(g2_se_fin, se_reporte_verif)

tabprov_se_fin <- resolve_auto(tabprov_se_fin, se_reporte_colec)
tabse_anio <- resolve_auto(tabse_anio, anio_rep_colec)
tabse_se <- resolve_auto(tabse_se, se_reporte_colec)

# --------------------------- #
# 5) GRÁFICO 1 (SEGÚN COLECCIÓN)
# --------------------------- #

create_graph1 <- function(dat, cols, g1_anio, g1_se_inicio, g1_unidad, se_reporte_colec, carpeta, week_system) {
  # Usar SE según colección
  base_g1 <- dat %>%
    filter(clasif %in% c("NEGATIVO", "POSITIVO")) %>%
    mutate(se = se_colec, anio = anio_colec) %>%
    mutate(
      provincia = if (!is.na(cols$col_prov)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_prov]]))) else NA_character_,
      micro_red = if (!is.na(cols$col_micro)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_micro]]))) else NA_character_,
      estab_origen = if (!is.na(cols$col_estab)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_estab]]))) else NA_character_
    )
  
  base_g1 <- if (g1_unidad == "muestra") dedup_by_unit(base_g1, "muestra") else base_g1
  
  sem <- base_g1 %>%
    count(anio, se, clasif, name = "n") %>%
    tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>%
    mutate(total = NEGATIVO + POSITIVO, IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>%
    arrange(anio, se) %>%
    group_by(anio) %>%
    tidyr::complete(se = 1:53, fill = list(NEGATIVO = 0, POSITIVO = 0)) %>%
    mutate(total = NEGATIVO + POSITIVO, IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>%
    ungroup()
  
  max_anio_g1 <- max(g1_anio, na.rm = TRUE)
  sem_plot <- sem %>%
    filter(anio < max_anio_g1 | (anio == max_anio_g1 & se <= se_reporte_colec))
  
  if (!is.null(g1_se_inicio)) {
    sem_plot <- sem_plot %>% filter(anio < max_anio_g1 | se >= g1_se_inicio)
  }
  
  if (nrow(sem_plot) == 0) stop("Gráfico 1: el filtro dejó el dataset vacío (revisa año/SE inicio).")
  
  sem_plot <- sem_plot %>%
    arrange(anio, se) %>%
    mutate(se_label = sprintf("%02d\n%d", se, anio))
  
  bars <- sem_plot %>%
    select(se_label, NEGATIVO, POSITIVO) %>%
    pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n")
  
  max_count <- max(bars$n, na.rm = TRUE)
  max_ip <- max(sem_plot$IP, na.rm = TRUE)
  scale_factor <- ifelse(is.finite(max_ip) && max_ip > 0, max_count / max_ip, 1)
  
  se_levels <- unique(sem_plot$se_label)
  bars <- bars %>% mutate(se_f = factor(se_label, levels = se_levels))
  sem_plot <- sem_plot %>% mutate(se_f = factor(se_label, levels = se_levels))
  
  sem_plot <- sem_plot %>%
    mutate(
      es_max_ip = IP == max(IP, na.rm = TRUE),
      etiqueta_ip = if_else(es_max_ip, sprintf("IP: %.1f%%", IP), NA_character_)
    )
  
  color_ip_impacto <- "#00C853"
  
  p1 <- ggplot() +
    geom_col(data = bars, aes(x = se_f, y = n, fill = tipo), position = position_dodge(width = 0.8), width = 0.7) +
    geom_line(data = sem_plot, aes(x = se_f, y = IP * scale_factor, group = 1), linewidth = 1.5, color = color_ip_impacto) +
    geom_point(data = sem_plot, aes(x = se_f, y = IP * scale_factor), size = 4, color = color_ip_impacto, fill = "white", shape = 21, stroke = 2) +
    geom_text(data = sem_plot, aes(x = se_f, y = IP * scale_factor, label = round(IP, 1)), vjust = -0.8, size = 5, fontface = "bold", color = color_ip_impacto) +
    scale_y_continuous(
      name = ifelse(g1_unidad == "examen", "Número de procesamientos", "Número de muestras"),
      sec.axis = sec_axis(~ . / scale_factor, name = "Índice de positividad (IP%)", labels = function(x) paste0(round(x, 1), "%"))
    ) +
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) +
    labs(x = NULL, fill = NULL) +
    theme_minimal(base_size = 16) +
    theme(panel.grid.minor = element_blank(), legend.position = "none", 
          axis.text = element_text(face = "bold"), axis.title = element_text(face = "bold"), 
          axis.text.x = element_blank())
  
  dat_tabla <- sem_plot %>%
    select(se_f, NEGATIVO, POSITIVO, IP) %>%
    mutate(IP = sprintf("%.1f", IP), NEGATIVO = as.character(NEGATIVO), POSITIVO = as.character(POSITIVO)) %>%
    pivot_longer(cols = c(NEGATIVO, POSITIVO, IP), names_to = "variable", values_to = "valor") %>%
    mutate(variable = factor(variable, levels = c("IP", "POSITIVO", "NEGATIVO")), 
           color_txt = case_when(variable == "NEGATIVO" ~ "#1F77B4", 
                                 variable == "POSITIVO" ~ "#D62728", 
                                 variable == "IP" ~ color_ip_impacto))
  
  p_tabla <- ggplot(dat_tabla, aes(x = se_f, y = variable)) +
    geom_tile(aes(fill = variable), alpha = 0.08, color = "white", linewidth = 0.5) +
    geom_text(aes(label = valor, color = variable), size = 4.5, fontface = "bold") +
    scale_color_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) +
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) +
    scale_y_discrete(labels = c("NEGATIVO" = "NEGATIVO (-)", "POSITIVO" = "POSITIVO (-)", "IP" = "IP (%)")) +
    labs(x = "Semana Epidemiológica") +
    theme_minimal(base_size = 14) +
    theme(panel.grid = element_blank(), legend.position = "none", 
          axis.title.y = element_blank(), 
          axis.text.y = element_text(face = "bold", color = "black", hjust = 1), 
          axis.text.x = element_text(face = "bold"), 
          plot.margin = margin(t = -5, r = 0, b = 0, l = 0))
  
  fig1 <- p1 / p_tabla + patchwork::plot_layout(heights = c(4, 1.3))
  out_g1 <- file.path(carpeta, "01_IP_por_SE_HighImpact.png")
  ggsave(out_g1, fig1, width = 18, height = 10, dpi = 300)
  
  list(sem_plot = sem_plot, out_g1 = out_g1)
}

graph1 <- create_graph1(dat, cols, g1_anio, g1_se_inicio, g1_unidad, se_reporte_colec, carpeta, week_system)
sem_plot <- graph1$sem_plot
out_g1 <- graph1$out_g1

# --------------------------- #
# 6) GRÁFICO 2 (ACUMULADO 2026 - SEGÚN VERIFICACIÓN)
# --------------------------- #

create_graph2 <- function(dat, g2_anio, g2_se_inicio, g2_se_fin, g2_excluir_labs, g2_lab_solo, g2_unidad, carpeta) {
  dat_g2 <- dat %>%
    mutate(
      prueba = case_when(
        str_detect(examen_key, "ac\\.?\\s*igm") ~ "Virus Dengue Ac. IgM",
        str_detect(examen_key, "ag\\s*ns1") ~ "Virus Dengue Ag NS1",
        TRUE ~ examen_std
      ),
      # Usar SE y año de VERIFICACIÓN para el gráfico 2
      se = se_verif,
      anio = anio_verif
    ) %>%
    filter(clasif %in% c("NEGATIVO", "POSITIVO"))
  
  # Aplicar filtros: año 2026 completo
  dat_g2 <- apply_filters(dat_g2, anio_sel = g2_anio, se_inicio = g2_se_inicio, se_fin = g2_se_fin, 
                          labs_excluir = g2_excluir_labs, labs_solo = g2_lab_solo)
  
  if (nrow(dat_g2) == 0) stop("Gráfico 2: no quedan datos tras filtros (año/SE/labs).")
  
  base_g2 <- if (g2_unidad == "muestra") dedup_by_unit(dat_g2, "muestra") else dat_g2
  
  res_g2 <- base_g2 %>%
    count(prueba, clasif, name = "n") %>%
    tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>%
    mutate(TOTAL = NEGATIVO + POSITIVO)
  
  bars_g2 <- res_g2 %>%
    select(prueba, NEGATIVO, POSITIVO) %>%
    pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n")
  
  ymax <- max(res_g2$TOTAL, na.rm = TRUE)
  offset_total <- ymax * 0.05
  
  # Calcular SE mínima y máxima de 2026 según verificación
  se_min_2026 <- min(dat_g2$se, na.rm = TRUE)
  se_max_2026 <- max(dat_g2$se, na.rm = TRUE)
  
  subtitulo_dinamico <- sprintf("Acumulado año %d (SE %02d a SE %02d según verificación)", 
                                g2_anio, se_min_2026, se_max_2026)
  
  p2 <- ggplot() +
    geom_col(data = bars_g2, aes(x = prueba, y = n, fill = tipo), position = position_dodge(width = 0.8), width = 0.7) +
    geom_text(data = bars_g2, aes(x = prueba, y = n, label = scales::comma(n), group = tipo), 
              position = position_dodge(width = 0.8), vjust = -0.5, size = 6, fontface = "bold") +
    geom_text(data = res_g2, aes(x = prueba, y = TOTAL + offset_total, 
                                 label = paste0("TOTAL:\n", scales::comma(TOTAL))), 
              size = 8, fontface = "bold", color = "black", lineheight = 0.8) +
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) +
    scale_y_continuous(labels = scales::comma, expand = expansion(mult = c(0, 0.2))) +
    labs(x = NULL, y = "NÚMERO DE PRUEBAS", fill = NULL, 
         title = "PROCESAMIENTO DE MUESTRAS POR TIPO DE PRUEBA - ACUMULADO 2026",
         subtitle = subtitulo_dinamico) +
    theme_minimal(base_size = 18) +
    theme(legend.position = "top", panel.grid.major.x = element_blank(), 
          panel.grid.minor = element_blank(), axis.text.x = element_text(face = "bold", size = 14), 
          axis.text.y = element_text(face = "bold"), 
          plot.title = element_text(face = "bold", hjust = 0.5, size = 22), 
          plot.subtitle = element_text(face = "italic", hjust = 0.5, size = 16, color = "#555555"))
  
  out_g2 <- file.path(carpeta, "02_Procesamiento_por_prueba_HighImpact.png")
  ggsave(out_g2, p2, width = 16, height = 9, dpi = 300)
  
  list(out_g2 = out_g2, res_g2 = res_g2, dat_g2 = dat_g2)
}

graph2 <- create_graph2(dat, g2_anio, g2_se_inicio, g2_se_fin, g2_excluir_labs, g2_lab_solo, g2_unidad, carpeta)
out_g2 <- graph2$out_g2
res_g2 <- graph2$res_g2

# --------------------------- #
# 7) TABLA 03: POSITIVOS POR PROVINCIA (SEGÚN COLECCIÓN 2026)
# --------------------------- #

create_table_prov <- function(dat, col_prov, tabprov_anio, tabprov_se_inicio, tabprov_se_fin, 
                              tabprov_excluir_labs, tabprov_lab_solo, tabprov_unidad, carpeta) {
  if (is.na(col_prov)) {
    warning("No se detectó columna PROVINCIA. Se omite tabla por provincia.")
    return(NULL)
  }
  
  dat_tabprov <- dat %>%
    mutate(
      provincia = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_prov]]))),
      provincia = if_else(is.na(provincia) | provincia == "", "SIN DATO", provincia),
      prueba = case_when(str_detect(examen_key, "ac\\.?\\s*igm") ~ "IgM", 
                         str_detect(examen_key, "ag\\s*ns1") ~ "NS1", 
                         TRUE ~ "OTRO"),
      # Usar SE según colección para esta tabla
      se = se_colec,
      anio = anio_colec
    ) %>%
    filter(prueba %in% c("IgM", "NS1"))
  
  # Aplicar filtros: solo año 2026 según colección
  dat_tabprov <- apply_filters(dat_tabprov, anio_sel = tabprov_anio, se_inicio = tabprov_se_inicio, 
                               se_fin = tabprov_se_fin, labs_excluir = tabprov_excluir_labs, 
                               labs_solo = tabprov_lab_solo)
  
  if (nrow(dat_tabprov) == 0) stop("Tabla provincia: no quedan datos tras filtros.")
  
  # Filtrar solo positivos
  base_pos <- dat_tabprov %>% filter(clasif == "POSITIVO")
  
  if (tabprov_unidad == "muestra") {
    base_pos <- dat_tabprov %>%
      group_by(cod_muestra, provincia, prueba) %>%
      summarise(clasif = if_else(any(clasif == "POSITIVO"), "POSITIVO", "NEGATIVO"), .groups = "drop") %>%
      filter(clasif == "POSITIVO") %>%
      select(provincia, prueba)
  } else {
    base_pos <- base_pos %>% select(provincia, prueba)
  }
  
  tab_pos <- base_pos %>%
    count(provincia, prueba, name = "n") %>%
    tidyr::pivot_wider(names_from = prueba, values_from = n, values_fill = 0) %>%
    mutate(`Total / POSITIVOS` = IgM + NS1, 
           `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0), 
           `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0)) %>%
    arrange(desc(`Total / POSITIVOS`))
  
  fila_total <- tibble::tibble(PROVINCIA = "Total", IgM = sum(tab_pos$IgM, na.rm = TRUE), 
                               NS1 = sum(tab_pos$NS1, na.rm = TRUE)) %>%
    mutate(`Total / POSITIVOS` = IgM + NS1, 
           `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0), 
           `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0))
  
  tab_final <- tab_pos %>%
    transmute(PROVINCIA = provincia, IgM, NS1, `% IgM`, `% NS1`, `Total / POSITIVOS`) %>%
    bind_rows(fila_total)
  
  out_tabprov <- file.path(carpeta, "03_Tabla_positivos_por_provincia.xlsx")
  style_xlsx(tab_final, out_tabprov, pct_cols = c("% IgM", "% NS1"), total_row_idx = nrow(tab_final))
  
  tab_final
}

tab_final <- create_table_prov(dat, cols$col_prov, tabprov_anio, tabprov_se_inicio, tabprov_se_fin, 
                               tabprov_excluir_labs, tabprov_lab_solo, tabprov_unidad, carpeta)

# --------------------------- #
# 8) TABLA 04: MICRORED x ESTABLECIMIENTO (SEGÚN COLECCIÓN ACTUAL)
# --------------------------- #

create_table_se <- function(dat, col_micro, col_estab, tabse_anio, tabse_se, 
                            tabse_excluir_labs, tabse_lab_solo, tabse_unidad, carpeta) {
  if (is.na(col_micro) || is.na(col_estab)) {
    warning("No se detectó Micro RED y/o Establecimiento de Origen. Se omite tabla por SE.")
    return(NULL)
  }
  
  dat_tabse <- dat %>%
    mutate(
      micro_red = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_micro]]))),
      estab_origen = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_estab]]))),
      micro_red = if_else(is.na(micro_red) | micro_red == "", "SIN DATO", micro_red),
      estab_origen = if_else(is.na(estab_origen) | estab_origen == "", "SIN DATO", estab_origen),
      # Usar SE según colección para esta tabla
      se = se_colec,
      anio = anio_colec
    ) %>%
    filter(clasif %in% c("NEGATIVO", "POSITIVO"))
  
  # Aplicar filtros: solo SE actual según colección
  dat_tabse <- apply_filters(dat_tabse, anio_sel = tabse_anio, se_inicio = tabse_se, se_fin = tabse_se, 
                             labs_excluir = tabse_excluir_labs, labs_solo = tabse_lab_solo)
  
  if (nrow(dat_tabse) == 0) stop("Tabla SE: no hay registros para ese año/SE con los filtros actuales.")
  
  if (tabse_unidad == "muestra") {
    dat_tabse <- dedup_by_unit(dat_tabse, "muestra") %>% filter(clasif %in% c("NEGATIVO", "POSITIVO"))
  }
  
  tabla_se <- dat_tabse %>%
    group_by(micro_red, estab_origen) %>%
    summarise(`NEGATIVO -` = sum(clasif == "NEGATIVO", na.rm = TRUE), 
              `POSITIVO -` = sum(clasif == "POSITIVO", na.rm = TRUE), .groups = "drop") %>%
    mutate(`Total general` = `NEGATIVO -` + `POSITIVO -`, 
           IP = if_else(`Total general` > 0, 100 * `POSITIVO -` / `Total general`, NA_real_), 
           SE = sprintf("SE %02d", as.integer(tabse_se))) %>%
    relocate(SE, .before = micro_red) %>%
    arrange(desc(`Total general`), micro_red, estab_origen)
  
  tot_neg <- sum(tabla_se$`NEGATIVO -`, na.rm = TRUE)
  tot_pos <- sum(tabla_se$`POSITIVO -`, na.rm = TRUE)
  tot_all <- sum(tabla_se$`Total general`, na.rm = TRUE)
  tot_ip <- ifelse(tot_all > 0, 100 * tot_pos / tot_all, NA_real_)
  
  fila_total <- tibble::tibble(SE = sprintf("SE %02d", as.integer(tabse_se)), 
                               micro_red = "Total general", estab_origen = "", 
                               `NEGATIVO -` = tot_neg, `POSITIVO -` = tot_pos, 
                               `Total general` = tot_all, IP = tot_ip)
  
  tabla_se_final <- bind_rows(tabla_se, fila_total) %>% 
    rename(`Micro RED` = micro_red, `Establecimiento de Origen` = estab_origen)
  
  out_tabse <- file.path(carpeta, sprintf("04_Tabla_SE%02d_Microred_Establecimiento.xlsx", as.integer(tabse_se)))
  style_xlsx(tabla_se_final, out_tabse, pct_cols = c("IP"), total_row_idx = nrow(tabla_se_final))
  
  tabla_se_final
}

tabla_se_final <- create_table_se(dat, cols$col_micro, cols$col_estab, tabse_anio, tabse_se, 
                                  tabse_excluir_labs, tabse_lab_solo, tabse_unidad, carpeta)

# --------------------------- #
# 9) MENSAJES FINALES
# --------------------------- #

message("OK. Carpeta del reporte: ", carpeta)
message("SE del reporte según Colección: SE ", sprintf("%02d", se_reporte_colec), " - Año ", anio_rep_colec)
message("SE del reporte según Verificación: SE ", sprintf("%02d", se_reporte_verif), " - Año ", anio_rep_verif)
message("Gráfico 1 (según Colección): ", out_g1)
message("Gráfico 2 (Acumulado 2026 según Verificación): ", out_g2)
message("Tabla 03 (Positivos 2026 según Colección): 03_Tabla_positivos_por_provincia.xlsx")
message("Tabla 04 (Microred x Establecimiento SE actual según Colección): 04_Tabla_SE_Microred_Establecimiento.xlsx")
message("Listo.")
