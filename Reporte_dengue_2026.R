# ==========================================================
# REPORTE SEMANAL DENGUE - NETLABv2 (OPTIMIZADO) [CORREGIDO]
# - SE: MMWR (Domingo–Sábado) usando epiweek/epiyear
# - Filtros: Solo 2 exámenes dengue (NS1 e IgM) + RESULTADO VERIFICADO
# - Salidas: 2 gráficos + 2 tablas Excel formateadas
# ==========================================================

rm(list = ls())



# ---------------------------
# 0) PAQUETES
# ---------------------------
req_pkgs <- c(
  "readxl","dplyr","stringr","lubridate","tidyr","ggplot2",
  "patchwork","gridExtra","scales","openxlsx","tibble",
  "officer", "flextable", "webshot2"
)

to_install <- req_pkgs[!sapply(req_pkgs, requireNamespace, quietly = TRUE)]
if (length(to_install) > 0) install.packages(to_install)

# Cargar librerías
invisible(lapply(req_pkgs, library, character.only = TRUE))

message("Librerías cargadas correctamente. Puedes continuar con la Sección 1.")
# ---------------------------
# 1) CONFIGURACIÓN (EDITA SOLO ESTO)
# ---------------------------

# Archivo / hoja (ajusta a tu descarga NETLAB)
archivo <- "dengue_2025.xlsx"  # <- AJUSTA AQUÍ si tu nombre/ruta es diferente
hoja    <- 1                      # o "NET LAB 2025"

# Exámenes permitidos (SOLO estos dos)
examenes_permitidos <- c(
  "Virus Dengue Ag NS1 [Presencia] en Suero o Plasma por Inmunoensayo",
  "Virus Dengue Ac. IgM [Presencia] en Suero o Plasma por Inmunoensayo"
)

# Unidad global (por defecto: procesamiento total)
unidad_global <- "examen"     # "examen" = procesamiento total (default)
# unidad_global <- "muestra"  # "muestra" = muestra única (dedup por código)

# Filtro opcional por Laboratorio Destino (GLOBAL):
lab_destino_global <- "TODOS"
# lab_destino_global <- c("LABORATORIO DE REFERENCIA REGIONAL DE LORETO")  # descomenta para filtrar

# Semana epidemiológica: MMWR (Domingo–Sábado)
week_system <- "MMWR"   # "MMWR" o "ISO"

# Carpeta del reporte: por defecto "SE XX"
incluir_anio_en_carpeta <- FALSE

# ------------- GRÁFICO 1 (IP% por SE) -------------
g1_anio <- "AUTO"     # o fija: 2025
g1_se_inicio <- 20    # ejemplo: 40 para que empiece desde SE40
# g1_se_inicio <- NULL
g1_unidad <- unidad_global

# ------------- GRÁFICO 2 (Procesamiento por tipo de prueba) -------------
g2_anio <- "AUTO"     # o fija: 2025 | o NULL = todos los años
g2_se_inicio <- NULL
g2_se_fin    <- NULL
g2_unidad    <- "examen"  # recomendado: procesamiento total

# Default: excluir LRNMZVIR
g2_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")

g2_lab_solo <- "TODOS"
# g2_lab_solo <- c("LABORATORIO DE REFERENCIA REGIONAL DE LORETO")

# ------------- TABLA POSITIVOS por PROVINCIA -------------
tabprov_anio <- "AUTO"   # <- SI QUIERES FILTRAR 2025: tabprov_anio <- 2025
tabprov_se_inicio <- NULL
tabprov_se_fin    <- NULL
tabprov_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")
tabprov_lab_solo <- "TODOS"
tabprov_unidad <- "examen"

# ------------- TABLA POR SE (MicroRED x EE.SS. Origen) -------------
tabse_anio <- 2025        # <- AQUÍ FILTRAS AÑO 2025 (o usa "AUTO")
tabse_se   <- "AUTO"      # <- AQUÍ PONES LA SE (ej: 51) o "AUTO"
tabse_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES")
tabse_lab_solo <- "TODOS"
tabse_unidad <- "examen"

# ---------------------------
# 2) FUNCIONES (NO EDITAR)
# ---------------------------

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
    orders = c("dmy","dmY","ymd","Ymd","d/m/Y","Y-m-d","d-m-Y","Y/m/d"),
    tz = "UTC"
  ))
}

add_epi <- function(df, date_col, week_system = c("MMWR","ISO")) {
  week_system <- match.arg(week_system)
  if (week_system == "ISO") {
    df %>% mutate(se = lubridate::isoweek(.data[[date_col]]),
                  anio = lubridate::isoyear(.data[[date_col]]))
  } else {
    df %>% mutate(se = lubridate::epiweek(.data[[date_col]]),
                  anio = lubridate::epiyear(.data[[date_col]]))
  }
}

# --------- CORRECCIÓN CLAVE: apply_filters() ----------
apply_filters <- function(df,
                          anio_sel = NULL,
                          se_inicio = NULL,
                          se_fin = NULL,
                          labs_excluir = NULL,
                          labs_solo = "TODOS") {
  
  out <- df
  
  # Excluir labs
  if (!is.null(labs_excluir) && length(labs_excluir) > 0) {
    ex <- stringr::str_to_upper(stringr::str_squish(labs_excluir))
    out <- out %>% dplyr::filter(!(.data$lab_destino_std %in% ex))
  }
  
  # Solo labs
  if (!(length(labs_solo) == 1 && toupper(labs_solo) == "TODOS")) {
    inc <- stringr::str_to_upper(stringr::str_squish(labs_solo))
    out <- out %>% dplyr::filter(.data$lab_destino_std %in% inc)
  }
  
  # Año(s)  ✅ (ANTES estaba mal)
  if (!is.null(anio_sel)) {
    out <- out %>% dplyr::filter(.data$anio %in% anio_sel)
  }
  
  # Rango SE
  if (!is.null(se_inicio)) out <- out %>% dplyr::filter(.data$se >= se_inicio)
  if (!is.null(se_fin))    out <- out %>% dplyr::filter(.data$se <= se_fin)
  
  out
}

dedup_by_unit <- function(df, unit = c("examen","muestra")) {
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
        openxlsx::addStyle(wb, "Reporte", pctStyle,
                           rows = 2:(nrow(df) + 1), cols = col_i,
                           gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  if (!is.na(total_row_idx)) {
    totalStyle <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9E1F2", border = "Top")
    openxlsx::addStyle(wb, "Reporte", totalStyle,
                       rows = total_row_idx + 1, cols = 1:ncol(df),
                       gridExpand = TRUE, stack = TRUE)
  }
  
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  path
}

# ---------------------------
# 3) CARGA + DETECCIÓN COLUMNAS
# ---------------------------

raw <- readxl::read_excel(archivo, sheet = hoja)

col_fecha   <- must_col(raw, c("Fecha Colección","Fecha Coleccion"), "Fecha de colección")
col_result  <- must_col(raw, c("resultado"), "Resultado")
col_examen  <- must_col(raw, c("nombre de examen","examen"), "Examen")
col_estatus <- must_col(raw, c("estatus resultado","estado resultado","estatus"), "Estatus")
col_cod     <- must_col(raw, c("codigo de muestra","codigo muestra","cod muestra","muestra"), "Código de muestra")
col_labdest <- must_col(raw, c("laboratorio destino","lab destino","destino"), "Laboratorio destino")

# Opcionales
col_prov  <- find_col(raw, c("provincia", "provincia procedencia", "provincia de procedencia", "provincia domicilio"))
col_micro <- find_col(raw, c("Micro Red EE.SS Origen","Micro Red","MicroRED","Microred"))
col_estab <- find_col(raw, c("Establecimiento de Origen","IPRESS Origen","Origen"))

# ---------------------------
# 4) LIMPIEZA + FILTRO DENGUE
# ---------------------------

examenes_key <- norm_txt(examenes_permitidos)

dat <- raw %>%
  mutate(
    fecha_coleccion = parse_excel_date(.data[[col_fecha]]),
    resultado_std   = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_result]]))),
    examen_std      = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_examen]]))),
    examen_key      = norm_txt(examen_std),
    estatus_std     = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_estatus]]))),
    cod_muestra     = stringr::str_squish(as.character(.data[[col_cod]])),
    lab_destino_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_labdest]])))
  ) %>%
  filter(!is.na(fecha_coleccion)) %>%
  filter(estatus_std == "RESULTADO VERIFICADO") %>%
  filter(examen_key %in% examenes_key) %>%
  mutate(
    clasif = case_when(
      str_detect(resultado_std, "POSITIV")    ~ "POSITIVO",
      str_detect(resultado_std, "NEGATIV")    ~ "NEGATIVO",
      str_detect(resultado_std, "INDETERMIN") ~ "INDETERMINADO",
      TRUE ~ "OTRO"
    )
  )

# Filtro global opcional por lab destino
if (!(length(lab_destino_global) == 1 && toupper(lab_destino_global) == "TODOS")) {
  lab_ok <- stringr::str_to_upper(stringr::str_squish(lab_destino_global))
  dat <- dat %>% filter(lab_destino_std %in% lab_ok)
}

if (nrow(dat) == 0) stop("No quedan registros tras filtros: verificado + exámenes permitidos + lab destino (si aplica).")

# ---------------------------
# 5) AÑADIR SE/AÑO EPIDEMIOLÓGICO
# ---------------------------
dat <- add_epi(dat, "fecha_coleccion", week_system = week_system)

# ---------------------------
# 6) DEFINIR SE/AÑO DEL REPORTE + CARPETA
# ---------------------------
fecha_max <- max(dat$fecha_coleccion, na.rm = TRUE)
se_reporte <- if (week_system == "ISO") lubridate::isoweek(fecha_max) else lubridate::epiweek(fecha_max)
anio_rep   <- if (week_system == "ISO") lubridate::isoyear(fecha_max) else lubridate::epiyear(fecha_max)

carpeta <- if (incluir_anio_en_carpeta) sprintf("%d_SE %02d", anio_rep, se_reporte) else sprintf("SE %02d", se_reporte)
dir.create(carpeta, showWarnings = FALSE, recursive = TRUE)

resolve_auto <- function(x, auto_value) {
  if (is.character(x) && length(x) == 1 && toupper(x) == "AUTO") return(auto_value)
  x
}

g1_anio      <- resolve_auto(g1_anio, anio_rep)
g2_anio      <- resolve_auto(g2_anio, anio_rep)
tabprov_anio <- resolve_auto(tabprov_anio, anio_rep)
tabse_anio   <- resolve_auto(tabse_anio, anio_rep)
tabse_se     <- resolve_auto(tabse_se, se_reporte)

# ==========================================================
# GRÁFICO 1: NEG/POS + IP% por SE (con tabla inferior) - MEJORADO
# ==========================================================


base_g1 <- dat %>%
  filter(clasif %in% c("NEGATIVO","POSITIVO")) %>%
  mutate(
    provincia = if (!is.na(col_prov)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_prov]]))) else NA_character_,
    micro_red = if (!is.na(col_micro)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_micro]]))) else NA_character_,
    estab_origen = if (!is.na(col_estab)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_estab]]))) else NA_character_
  )

base_g1 <- if (g1_unidad == "muestra") dedup_by_unit(base_g1, "muestra") else base_g1

sem <- base_g1 %>%
  count(anio, se, clasif, name = "n") %>%
  tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>%
  mutate(total = NEGATIVO + POSITIVO,
         IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>%
  arrange(anio, se) %>%
  group_by(anio) %>%
  tidyr::complete(se = 1:53, fill = list(NEGATIVO = 0, POSITIVO = 0)) %>%
  mutate(total = NEGATIVO + POSITIVO,
         IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>%
  ungroup()

sem_plot <- sem %>% filter(anio == g1_anio, se <= se_reporte)
if (!is.null(g1_se_inicio)) sem_plot <- sem_plot %>% filter(se >= g1_se_inicio)
if (nrow(sem_plot) == 0) stop("Gráfico 1: el filtro dejó el dataset vacío (revisa año/SE inicio).")

bars <- sem_plot %>%
  select(se, NEGATIVO, POSITIVO) %>%
  pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n")

max_count <- max(bars$n, na.rm = TRUE)
max_ip    <- max(sem_plot$IP, na.rm = TRUE)
scale_factor <- ifelse(is.finite(max_ip) && max_ip > 0, max_count / max_ip, 1)

se_levels <- sort(unique(sem_plot$se))
bars <- bars %>% mutate(se_f = factor(se, levels = se_levels))
sem_plot <- sem_plot %>% mutate(se_f = factor(se, levels = se_levels))

# ==========================================================
# GRÁFICO 1: NEG/POS + IP% por SE (con tabla inferior) - MEJORADO
# ==========================================================

# Identificar el IP máximo para destacarlo
sem_plot <- sem_plot %>%
  mutate(
    es_max_ip = IP == max(IP, na.rm = TRUE),
    etiqueta_ip = if_else(es_max_ip, sprintf("IP: %.1f%%", IP), NA_character_)
  )

max_ip_point <- sem_plot %>% filter(es_max_ip)

# --- INICIO REEMPLAZO BLOQUE GRÁFICO 1 (VERDE NEÓN + TABLA ALINEADA) ---

# COLOR DE IMPACTO: Verde Neón "High Voltage"
# Usamos un verde muy saturado (#00E600) que resalta sobre blanco.
color_ip_impacto <- "#00C853" 

# 1.1 Gráfico Principal (p1)
p1 <- ggplot() +
  geom_col(data = bars, aes(x = se_f, y = n, fill = tipo),
           position = position_dodge(width = 0.8), width = 0.7) +
  # LÍNEA IP: Verde Neón y gruesa
  geom_line(data = sem_plot, aes(x = se_f, y = IP * scale_factor, group = 1), 
            linewidth = 1.5, color = color_ip_impacto) +
  # PUNTOS IP: Círculo con relleno blanco y borde verde neón grueso
  geom_point(data = sem_plot, aes(x = se_f, y = IP * scale_factor), 
             size = 4, color = color_ip_impacto, fill = "white", shape = 21, stroke = 2) +
  # ETIQUETAS IP: En negrita y color verde neón
  geom_text(data = sem_plot, aes(x = se_f, y = IP * scale_factor, label = round(IP, 1)),
            vjust = -0.8, size = 5, fontface = "bold", color = color_ip_impacto) +
  scale_y_continuous(
    name = ifelse(g1_unidad == "examen", "Número de procesamientos", "Número de muestras"),
    sec.axis = sec_axis(~ . / scale_factor, name = "Índice de positividad (IP%)",
                        labels = function(x) paste0(round(x, 1), "%"))
  ) +
  scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) +
  labs(x = NULL, fill = NULL) + 
  theme_minimal(base_size = 16) +
  theme(
    panel.grid.minor = element_blank(),
    legend.position = "none", 
    axis.text = element_text(face = "bold"),
    axis.title = element_text(face = "bold"),
    # Ocultamos el texto del eje X aquí para ponerlo en la tabla si se desea, 
    # o lo dejamos para referencia visual. Lo dejamos por claridad.
    axis.text.x = element_blank() 
  )

# 1.2 Tabla construida como GGPLOT (Truco para ancho exacto)
# [CORREGIDO: Convertimos todo a texto para evitar error de tipos]

dat_tabla <- sem_plot %>%
  select(se_f, NEGATIVO, POSITIVO, IP) %>%
  mutate(
    # Convertimos todo a caracteres explícitamente para que pivot_longer no falle
    IP = sprintf("%.1f", IP),
    NEGATIVO = as.character(NEGATIVO),
    POSITIVO = as.character(POSITIVO)
  ) %>%
  pivot_longer(cols = c(NEGATIVO, POSITIVO, IP), names_to = "variable", values_to = "valor") %>%
  mutate(
    # Ordenar eje Y de la tabla: Negativo arriba, IP abajo
    variable = factor(variable, levels = c("IP", "POSITIVO", "NEGATIVO")),
    # Lógica de color (variable auxiliar, aunque usaremos scale_color_manual directo en el plot)
    color_txt = case_when(
      variable == "NEGATIVO" ~ "#1F77B4",
      variable == "POSITIVO" ~ "#D62728",
      variable == "IP" ~ color_ip_impacto
    )
  )

p_tabla <- ggplot(dat_tabla, aes(x = se_f, y = variable)) +
  # Fondo suave alternado (geom_tile)
  geom_tile(aes(fill = variable), alpha = 0.08, color = "white", linewidth = 0.5) +
  # Los números
  geom_text(aes(label = valor, color = variable), size = 4.5, fontface = "bold") +
  # Escalas de color fijas
  scale_color_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) +
  scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) +
  # Etiquetas eje Y de la tabla
  scale_y_discrete(labels = c("NEGATIVO" = "NEGATIVO (-)", "POSITIVO" = "POSITIVO (-)", "IP" = "IP (%)")) +
  labs(x = "Semana Epidemiológica") +
  theme_minimal(base_size = 14) +
  theme(
    panel.grid = element_blank(),
    legend.position = "none",
    axis.title.y = element_blank(),
    axis.text.y = element_text(face = "bold", color = "black", hjust = 1),
    axis.text.x = element_text(face = "bold"), # Muestra SE abajo
    plot.margin = margin(t = -5, r = 0, b = 0, l = 0) # Unir visualmente arriba
  )

# 1.3 Unir con Patchwork (Garantiza anchos idénticos)
fig1 <- p1 / p_tabla + 
  patchwork::plot_layout(heights = c(4, 1.3))

out_g1 <- file.path(carpeta, "01_IP_por_SE_HighImpact.png")
ggsave(out_g1, fig1, width = 18, height = 10, dpi = 300)



# --- FIN REEMPLAZO BLOQUE GRÁFICO 1 ---

# ==========================================================
# GRÁFICO 2: Procesamiento por tipo de prueba - ALTO IMPACTO
# ==========================================================

dat_g2 <- dat %>%
  mutate(
    prueba = case_when(
      str_detect(examen_key, "ac\\.?\\s*igm") ~ "Virus Dengue Ac. IgM",
      str_detect(examen_key, "ag\\s*ns1")     ~ "Virus Dengue Ag NS1",
      TRUE ~ examen_std
    )
  ) %>%
  filter(clasif %in% c("NEGATIVO","POSITIVO"))

anio_f_g2 <- if (is.null(g2_anio)) NULL else g2_anio

dat_g2 <- apply_filters(
  dat_g2,
  anio_sel = anio_f_g2,
  se_inicio = g2_se_inicio,
  se_fin = g2_se_fin,
  labs_excluir = g2_excluir_labs,
  labs_solo = g2_lab_solo
)

if (nrow(dat_g2) == 0) stop("Gráfico 2: no quedan datos tras filtros (año/SE/labs).")

# Unidad para G2 (por defecto: examen)
base_g2 <- if (g2_unidad == "muestra") dedup_by_unit(dat_g2, "muestra") else dat_g2

res_g2 <- base_g2 %>%
  count(prueba, clasif, name = "n") %>%
  tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>%
  mutate(TOTAL = NEGATIVO + POSITIVO)

bars_g2 <- res_g2 %>%
  select(prueba, NEGATIVO, POSITIVO) %>%
  pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n")

# --- INICIO REEMPLAZO BLOQUE GRÁFICO 2 ---

# Cálculo de límites para el título dinámico
ymax <- max(res_g2$TOTAL, na.rm = TRUE)
offset_total <- ymax * 0.05
min_se_g2 <- min(dat_g2$se, na.rm = TRUE)
max_se_g2 <- max(dat_g2$se, na.rm = TRUE)

# Texto dinámico: "Desde SE XX hasta la SE YY Año ZZZZ"
subtitulo_dinamico <- sprintf("Desde SE %02d hasta la SE %02d %s", 
                              min_se_g2, max_se_g2, g2_anio)

p2 <- ggplot() +
  geom_col(data = bars_g2,
           aes(x = prueba, y = n, fill = tipo),
           position = position_dodge(width = 0.8), width = 0.7) +
  geom_text(data = bars_g2,
            aes(x = prueba, y = n, label = scales::comma(n), group = tipo),
            position = position_dodge(width = 0.8),
            vjust = -0.5, size = 6, fontface = "bold") +
  # Etiqueta TOTAL mejorada
  geom_text(data = res_g2,
            aes(x = prueba, y = TOTAL + offset_total, 
                label = paste0("TOTAL:\n", scales::comma(TOTAL))),
            size = 8, fontface = "bold", color = "black", lineheight = 0.8) +
  scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) +
  scale_y_continuous(labels = scales::comma, expand = expansion(mult = c(0, 0.2))) +
  labs(x = NULL, 
       y = "NÚMERO DE PRUEBAS",
       fill = NULL,
       title = "PROCESAMIENTO DE MUESTRAS POR TIPO DE PRUEBA",
       subtitle = subtitulo_dinamico) + # Aquí insertamos el texto que pediste
  theme_minimal(base_size = 18) +
  theme(legend.position = "top",
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        axis.text.x = element_text(face = "bold", size = 14),
        axis.text.y = element_text(face = "bold"),
        plot.title = element_text(face = "bold", hjust = 0.5, size = 22),
        plot.subtitle = element_text(face = "italic", hjust = 0.5, size = 16, color = "#555555"))

out_g2 <- file.path(carpeta, "02_Procesamiento_por_prueba_HighImpact.png")
ggsave(out_g2, p2, width = 16, height = 9, dpi = 300)

# --- FIN REEMPLAZO BLOQUE GRÁFICO 2 ---
# ==========================================================
# ==========================================================
# TABLA: SOLO POSITIVOS por PROVINCIA (IgM vs NS1) + % dentro provincia
# ==========================================================

if (is.na(col_prov)) {
  warning("No se detectó columna PROVINCIA. Se omite tabla por provincia.")
} else {
  
  dat_tabprov <- dat %>%
    mutate(
      provincia = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_prov]]))),
      provincia = if_else(is.na(provincia) | provincia == "", "SIN DATO", provincia),
      prueba = case_when(
        str_detect(examen_key, "ac\\.?\\s*igm") ~ "IgM",
        str_detect(examen_key, "ag\\s*ns1")     ~ "NS1",
        TRUE ~ "OTRO"
      )
    ) %>%
    filter(prueba %in% c("IgM","NS1"))
  
  anio_f_tabprov <- if (is.null(tabprov_anio)) NULL else tabprov_anio
  
  dat_tabprov <- apply_filters(
    dat_tabprov,
    anio_sel = anio_f_tabprov,   # ✅ CORREGIDO
    se_inicio = tabprov_se_inicio,
    se_fin = tabprov_se_fin,
    labs_excluir = tabprov_excluir_labs,
    labs_solo = tabprov_lab_solo
  )
  
  if (nrow(dat_tabprov) == 0) stop("Tabla provincia: no quedan datos tras filtros.")
  
  base_pos <- dat_tabprov %>% filter(clasif == "POSITIVO")
  
  # Unidad (provincia)
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
    mutate(
      `Total / POSITIVOS` = IgM + NS1,
      `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0),
      `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0)
    ) %>%
    arrange(desc(`Total / POSITIVOS`))
  
  fila_total <- tibble::tibble(
    PROVINCIA = "Total",
    IgM = sum(tab_pos$IgM, na.rm = TRUE),
    NS1 = sum(tab_pos$NS1, na.rm = TRUE)
  ) %>%
    mutate(
      `Total / POSITIVOS` = IgM + NS1,
      `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0),
      `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0)
    )
  
  tab_final <- tab_pos %>%
    transmute(
      PROVINCIA = provincia,
      IgM, NS1,
      `% IgM`, `% NS1`,
      `Total / POSITIVOS`
    ) %>%
    bind_rows(fila_total)
  
  out_tabprov <- file.path(carpeta, "03_Tabla_positivos_por_provincia.xlsx")
  style_xlsx(tab_final, out_tabprov, pct_cols = c("% IgM","% NS1"), total_row_idx = nrow(tab_final))
}


# ==========================================================
# TABLA POR SE: Micro RED x Establecimiento de Origen (NEG/POS/Total/IP)
# ==========================================================

if (is.na(col_micro) || is.na(col_estab)) {
  warning("No se detectó Micro RED y/o Establecimiento de Origen. Se omite tabla por SE.")
} else {
  
  dat_tabse <- dat %>%
    mutate(
      micro_red = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_micro]]))),
      estab_origen = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_estab]]))),
      micro_red = if_else(is.na(micro_red) | micro_red == "", "SIN DATO", micro_red),
      estab_origen = if_else(is.na(estab_origen) | estab_origen == "", "SIN DATO", estab_origen)
    ) %>%
    filter(clasif %in% c("NEGATIVO","POSITIVO"))
  
  anio_f_tabse <- if (is.null(tabse_anio)) NULL else tabse_anio
  
  dat_tabse <- apply_filters(
    dat_tabse,
    anio_sel = anio_f_tabse,   # ✅ CORREGIDO
    se_inicio = tabse_se,
    se_fin = tabse_se,
    labs_excluir = tabse_excluir_labs,
    labs_solo = tabse_lab_solo
  )
  
  if (nrow(dat_tabse) == 0) stop("Tabla SE: no hay registros para ese año/SE con los filtros actuales.")
  
  if (tabse_unidad == "muestra") {
    dat_tabse <- dedup_by_unit(dat_tabse, "muestra") %>%
      filter(clasif %in% c("NEGATIVO","POSITIVO"))
  }
  
  tabla_se <- dat_tabse %>%
    group_by(micro_red, estab_origen) %>%
    summarise(
      `NEGATIVO -` = sum(clasif == "NEGATIVO", na.rm = TRUE),
      `POSITIVO -` = sum(clasif == "POSITIVO", na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      `Total general` = `NEGATIVO -` + `POSITIVO -`,
      IP = if_else(`Total general` > 0, 100 * `POSITIVO -` / `Total general`, NA_real_),
      SE = sprintf("SE %02d", as.integer(tabse_se))
    ) %>%
    relocate(SE, .before = micro_red) %>%
    arrange(desc(`Total general`), micro_red, estab_origen)
  
  tot_neg <- sum(tabla_se$`NEGATIVO -`, na.rm = TRUE)
  tot_pos <- sum(tabla_se$`POSITIVO -`, na.rm = TRUE)
  tot_all <- sum(tabla_se$`Total general`, na.rm = TRUE)
  tot_ip  <- ifelse(tot_all > 0, 100 * tot_pos / tot_all, NA_real_)
  
  fila_total <- tibble::tibble(
    SE = sprintf("SE %02d", as.integer(tabse_se)),
    micro_red = "Total general",
    estab_origen = "",
    `NEGATIVO -` = tot_neg,
    `POSITIVO -` = tot_pos,
    `Total general` = tot_all,
    IP = tot_ip
  )
  
  tabla_se_final <- bind_rows(tabla_se, fila_total) %>%
    rename(`Micro RED` = micro_red, `Establecimiento de Origen` = estab_origen)
  
  out_tabse <- file.path(carpeta, sprintf("04_Tabla_SE%02d_Microred_Establecimiento.xlsx", as.integer(tabse_se)))
  style_xlsx(tabla_se_final, out_tabse, pct_cols = c("IP"), total_row_idx = nrow(tabla_se_final))
}

# ---------------------------
# 7) MENSAJES FINALES
# ---------------------------
message("OK. Carpeta del reporte: ", carpeta)
message("SE del reporte (según max Fecha Colección): SE ", sprintf("%02d", se_reporte), " - Año ", anio_rep)
message("Gráfico 1: ", out_g1)
message("Gráfico 2: ", out_g2)
message("Listo.")

# ==========================================================
# 8) GENERACIÓN PPT – ESTILO MEJORADO (FLEXTABLE)
# ==========================================================
suppressPackageStartupMessages({
  library(officer)
  library(dplyr)
  library(flextable)
  library(webshot2)
})

message("Generando PPT con tablas de alto impacto...")

# ---------------------------
# 8.0) Asegurar carpeta de salida
# ---------------------------
if (!exists("carpeta")) carpeta <- getwd()
dir.create(carpeta, showWarnings = FALSE, recursive = TRUE)

# ---------------------------
# 8.1) Cargar plantilla
# ---------------------------
ruta_plantilla <- "plantilla_base.pptx"
my_ppt <- if (file.exists(ruta_plantilla)) read_pptx(ruta_plantilla) else read_pptx()

# ---------------------------
# 8.2) Detectar layout
# ---------------------------
ls <- layout_summary(my_ppt)
candidatos_layout <- c("Blank","En blanco","Vacío","Vacio","Title Only", "Título solamente")
LAYOUT <- candidatos_layout[candidatos_layout %in% ls$layout][1]
if (is.na(LAYOUT)) LAYOUT <- ls$layout[1]
MASTER <- ls$master[ls$layout == LAYOUT][1]
if (is.na(MASTER)) MASTER <- ls$master[1]

dims <- slide_size(my_ppt)
W_SLIDE <- dims$width
H_SLIDE <- dims$height

# Constantes de diseño
H_BANNER_GRAF <- 0.75
H_BANNER_TAB  <- 0.85 # Un poco más alto para interpretaciones
H_FOOTER      <- 0.45
PAD           <- 0.12
COLOR_FONDO_TITULO <- "#D9E1F2"
COLOR_FONDO_FOOTER <- "#F2F2F2"

# --- FUNCIONES AUXILIARES DE DISEÑO PPT ---

add_rect <- function(ppt, left, top, width, height, fill) {
  # Crea un rectángulo sólido usando flextable vacío (truco rápido) o bloque XML
  # Para simplificar, usamos ph_with con un bloque de texto vacío con fondo
  ppt <- ph_with(
    ppt, 
    value = fpar(ftext(" ", fp_text(font.size = 1))), 
    location = ph_location(left = left, top = top, width = width, height = height, bg = fill)
  )
  ppt
}

add_banner2 <- function(ppt, title, subtitle = NULL, h = 0.75) {
  # Fondo
  ppt <- add_rect(ppt, 0, 0, W_SLIDE, h, COLOR_FONDO_TITULO)
  
  # Título Principal
  ppt <- ph_with(
    ppt,
    value = fpar(
      ftext(title, prop = fp_text(font.size = 20, bold = TRUE, color = "black", font.family = "Arial")),
      fp_p = fp_par(text.align = "center")
    ),
    location = ph_location(left = PAD, top = 0.05, width = W_SLIDE - 2*PAD, height = h * 0.6)
  )
  
  # Subtítulo / Interpretación
  if (!is.null(subtitle) && nchar(trimws(subtitle)) > 0) {
    ppt <- ph_with(
      ppt,
      value = fpar(
        ftext(subtitle, prop = fp_text(font.size = 12, italic = FALSE, color = "#333333", font.family = "Arial")),
        fp_p = fp_par(text.align = "center")
      ),
      location = ph_location(left = PAD, top = h * 0.55, width = W_SLIDE - 2*PAD, height = h * 0.4)
    )
  }
  ppt
}

add_footer <- function(ppt, texto) {
  ppt <- add_rect(ppt, 0, H_SLIDE - H_FOOTER, W_SLIDE, H_FOOTER, COLOR_FONDO_FOOTER)
  ppt <- ph_with(
    ppt,
    value = fpar(
      ftext(texto, prop = fp_text(font.size = 10, italic = TRUE, color = "#555555", font.family = "Arial")),
      fp_p = fp_par(text.align = "center")
    ),
    location = ph_location(left = PAD, top = H_SLIDE - H_FOOTER + 0.1, width = W_SLIDE - 2*PAD, height = H_FOOTER - 0.1)
  )
  ppt
}

add_img_in_body <- function(ppt, img_path, banner_h, footer = TRUE) {
  if (!file.exists(img_path)) return(ppt)
  h_body <- H_SLIDE - banner_h - if (footer) H_FOOTER else 0
  ph_with(
    ppt,
    value = external_img(img_path),
    location = ph_location(left = 0, top = banner_h, width = W_SLIDE, height = h_body)
  )
}

# -------------------------------------------------------------------------
# NUEVA FUNCIÓN: Generar Imagen de Tabla PRETTY con Flextable
# -------------------------------------------------------------------------
df_to_flextable_img <- function(df, path, width_in, height_in, base_size = 11) {
  
  # 1. Crear objeto flextable
  ft <- flextable(df) %>%
    # Tema limpio
    theme_vanilla() %>%
    # Fuente
    font(fontname = "Arial", part = "all") %>%
    fontsize(size = base_size, part = "all") %>%
    # Alineación: Texto a la izq, Números al centro
    align(align = "center", part = "all") %>%
    align(j = 1, align = "left", part = "all") %>% # Primera col (nombres) a la izq
    # Ajuste de ancho
    autofit()
  
  # 2. Estilo del Encabezado (HEADER)
  ft <- ft %>%
    bg(part = "header", bg = "#1F4E79") %>%  # Azul Institucional Oscuro
    color(part = "header", color = "white") %>%
    bold(part = "header") %>%
    border_remove() %>%
    hline_bottom(part = "header", border = fp_border(color = "white", width = 2))
  
  # 3. Estilo Cebra (Zebra Striping) en el cuerpo
  ft <- ft %>%
    bg(i = seq(1, nrow(df), 2), bg = "#F2F2F2", part = "body") %>% # Gris muy claro
    bg(i = seq(2, nrow(df), 2), bg = "white", part = "body")
  
  # 4. Detectar fila de TOTALES y estilizar
  # Buscamos filas que contengan "Total" o "General" en la primera columna
  idx_total <- which(str_detect(toupper(as.character(df[[1]])), "TOTAL") | 
                       str_detect(toupper(as.character(df[[1]])), "GENERAL"))
  
  if (length(idx_total) > 0) {
    ft <- ft %>%
      bold(i = idx_total, part = "body") %>%
      bg(i = idx_total, bg = "#D9E1F2", part = "body") %>% # Azul claro para resaltar
      hline(i = idx_total, border = fp_border(color = "#1F4E79", width = 1.5))
  }
  
  # 5. Formato Condicional (Ej: Colorear IP alto si existe la columna "IP")
  # (Opcional: puedes descomentar si quieres alertas rojas)
  # if ("IP" %in% names(df)) {
  #   ft <- color(ft, i = ~ IP > 20, j = "IP", color = "red", part = "body")
  # }
  
  # 6. Ajuste final de dimensiones para la imagen
  # Si la tabla es muy larga, reducimos un poco la fuente
  if (nrow(df) > 25) ft <- fontsize(ft, size = 9, part = "all")
  if (nrow(df) > 40) ft <- fontsize(ft, size = 8, part = "all")
  
  # Ajustar alto de filas para que sea compacto
  ft <- padding(ft, padding = 3, part = "all")
  
  # 7. Guardar como Imagen de Alta Calidad
  # flextable::save_as_image usa webshot2 por debajo
  save_as_image(ft, path = path, webshot = "webshot2", zoom = 3) # zoom=3 para alta resolución
  
  return(path)
}

# Wrapper para añadir tabla al PPT
add_table_slide_pretty <- function(ppt, df, title, subtitle, prefix, banner_h = 0.85) {
  if (is.null(df) || nrow(df) == 0) return(ppt)
  
  box_h <- H_SLIDE - banner_h
  
  # Nombre del archivo temporal
  out_png <- file.path(carpeta, sprintf("%s_pretty.png", prefix))
  
  # Generar la imagen bonita
  tryCatch({
    df_to_flextable_img(df, out_png, width_in = W_SLIDE, height_in = box_h)
  }, error = function(e) {
    warning("Error generando flextable: ", e$message)
  })
  
  # Añadir Slide
  ppt <- add_slide(ppt, layout = LAYOUT, master = MASTER)
  ppt <- add_banner2(ppt, title, subtitle, h = banner_h)
  
  # Insertar imagen si se creó
  if (file.exists(out_png)) {
    # Calcular posición para centrar verticalmente si sobra espacio
    ppt <- ph_with(
      ppt, 
      value = external_img(out_png),
      location = ph_location(left = 0, top = banner_h, width = W_SLIDE, height = box_h)
    )
  }
  ppt
}

# ==========================================================
# 8) GENERACIÓN PPT – ESTILO EJECUTIVO (CORREGIDO - FONDO SÓLIDO)
# ==========================================================

# --- PASO 0: FUNCIONES DE DISEÑO ROBUSTAS ---

# 1. Función para crear imagen de color sólido (Truco para fondos infalibles)
get_solid_img <- function(color, width_in, height_in) {
  # Nombre de archivo único por color
  fname <- paste0("solid_", gsub("#", "", color), ".png")
  fpath <- file.path(tempdir(), fname)
  
  if (!file.exists(fpath)) {
    png(fpath, width = 100, height = 100) # Pequeño, se estira luego
    par(mar = c(0,0,0,0))
    plot(0,0, type="n", xlim=c(0,1), ylim=c(0,1), axes=FALSE, xlab="", ylab="")
    rect(0,0,1,1, col=color, border=NA)
    dev.off()
  }
  return(fpath)
}

# 2. Función para el Banner Superior (Título Ejecutivo)
add_executive_banner <- function(ppt, title) {
  # A. Poner el FONDO AZUL (Imagen estirada)
  bg_img <- get_solid_img("#1F4E79", W_SLIDE, 0.85)
  ppt <- ph_with(
    ppt, 
    value = external_img(bg_img), 
    location = ph_location(left = 0, top = 0, width = W_SLIDE, height = 0.85)
  )
  
  # B. Poner el TEXTO encima (Blanco)
  ppt <- ph_with(
    ppt,
    value = fpar(
      ftext(title, prop = fp_text(font.size = 20, bold = TRUE, color = "white", font.family = "Arial")),
      fp_p = fp_par(text.align = "center")
    ),
    location = ph_location(left = 0.2, top = 0.1, width = W_SLIDE - 0.4, height = 0.7)
  )
  ppt
}

# 3. Función para la Caja de Interpretación (Debajo del Gráfico)
add_executive_interpretation <- function(ppt, text) {
  top_pos <- 6.1
  height_box <- 1.1 # Un poco más alto para asegurar que entre todo
  
  # A. Fondo Gris (Imagen)
  bg_img <- get_solid_img("#F2F2F2", W_SLIDE - 1.0, height_box)
  ppt <- ph_with(
    ppt, 
    value = external_img(bg_img), 
    location = ph_location(left = 0.5, top = top_pos, width = W_SLIDE - 1.0, height = height_box)
  )
  
  # B. Texto
  ppt <- ph_with(
    ppt,
    value = fpar(
      ftext("INTERPRETACIÓN:", prop = fp_text(font.size = 10, bold = TRUE, color = "#1F4E79", font.family = "Arial")),
      ftext("\r\n", prop = fp_text(font.size = 5)), 
      ftext(text, prop = fp_text(font.size = 11, color = "#333333", font.family = "Arial"))
    ),
    location = ph_location(left = 0.6, top = top_pos + 0.1, width = W_SLIDE - 1.2, height = height_box - 0.2)
  )
  ppt
}

# 4. Función para Imagen Central
add_centered_image <- function(ppt, img_path) {
  if (!file.exists(img_path)) return(ppt)
  ph_with(
    ppt,
    value = external_img(img_path),
    location = ph_location(left = 0, top = 0.9, width = W_SLIDE, height = 5.1)
  )
}

# 5. Función para Footer
add_executive_footer <- function(ppt, text) {
  ph_with(
    ppt,
    value = fpar(
      ftext(text, prop = fp_text(font.size = 9, italic = FALSE, color = "#777777", font.family = "Arial")),
      fp_p = fp_par(text.align = "right")
    ),
    location = ph_location(left = 0, top = 7.15, width = W_SLIDE - 0.3, height = 0.35)
  )
}


# --- PASO A: CÁLCULOS DE TEXTOS DINÁMICOS ---

# 1. Fechas y Semanas
min_se_g1 <- min(sem_plot$se)
max_se_g1 <- max(sem_plot$se) 
anio_g1   <- unique(sem_plot$anio)[1]

# Fecha de corte
fecha_corte_raw <- max(dat$fecha_coleccion, na.rm = TRUE)
fecha_corte_txt <- format(fecha_corte_raw, "%d.%m.%Y")

# 2. Datos para la Interpretación
fila_ultima_se <- sem_plot %>% filter(se == max_se_g1)
ip_actual <- if(nrow(fila_ultima_se) > 0) round(fila_ultima_se$IP, 1) else 0
ip_max_rango <- max(sem_plot$IP, na.rm = TRUE)
techo_ip <- ceiling(ip_max_rango)

# --- PASO B: CONSTRUCCIÓN DE FRASES ---

# Título G1
titulo_dinamico_g1 <- sprintf(
  "NÚMERO DE MUESTRAS DE DENGUE E INDICE DE POSITIVIDAD (IP) POR SEMANAS EPIDEMIOLÓGICAS. SE %d-%d – SE %d-%d. (CORTE %s)",
  min_se_g1, anio_g1, max_se_g1, anio_g1, fecha_corte_txt
)

# Interpretación G1
interp_dinamica_g1 <- sprintf(
  "El Índice de Positividad muestra el porcentaje de resultados positivos con respecto al número total de pruebas procesadas. A partir de la SE %d el IP se mantuvo constante menor al %d%%. En la SE %d el IP es %.1f%%.",
  min_se_g1, techo_ip, max_se_g1, ip_actual
)

# Título G2
titulo_dinamico_g2 <- sprintf(
  "PROCESAMIENTO DE MUESTRAS (NS1 E IGM) POR SEMANAS EPIDEMIOLÓGICAS. SE %d-%d – SE %d-%d. (CORTE %s)",
  min_se_g1, anio_g1, max_se_g1, anio_g1, fecha_corte_txt
)

# Interpretación G2
total_muestras_g2 <- sum(res_g2$TOTAL, na.rm = TRUE)
interp_dinamica_g2 <- sprintf(
  "Durante el periodo analizado (SE %d a %d), se procesaron un total de %s muestras. Se observa el comportamiento comparativo entre las pruebas de NS1 e IgM para la vigilancia virológica.",
  min_se_g1, max_se_g1, scales::comma(total_muestras_g2)
)

# Footer Exacto
txt_footer <- sprintf("Fuente: NETLABv2 | Corte: %s", fecha_corte_txt)


# --- PASO C: ARMADO DEL PPT ---

message("Generando diapositivas ejecutivas (con fondo corregido)...")

# 1. PORTADA
my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER)
my_ppt <- ph_with(my_ppt, value = fpar(ftext(" ", fp_text(font.size = 1))), location = ph_location_fullsize(bg = "white"))

my_ppt <- ph_with(
  my_ppt,
  value = fpar(
    ftext("VIGILANCIA DE DENGUE", fp_text(font.size = 44, bold = TRUE, color = "#1F4E79", font.family = "Arial")),
    fp_p = fp_par(text.align = "center")
  ),
  location = ph_location(left = 0, top = 2.5, width = W_SLIDE, height = 1.5)
)

subtitulo_portada <- sprintf("SEMANA %02d - %s", max_se_g1, anio_g1)
my_ppt <- ph_with(
  my_ppt,
  value = fpar(
    ftext(subtitulo_portada, fp_text(font.size = 24, color = "#666666", font.family = "Arial")),
    fp_p = fp_par(text.align = "center")
  ),
  location = ph_location(left = 0, top = 3.8, width = W_SLIDE, height = 1.0)
)
my_ppt <- add_executive_footer(my_ppt, txt_footer)


# 2. DIAPO GRÁFICO 1 (IP)
if (file.exists(out_g1)) {
  my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER)
  my_ppt <- add_executive_banner(my_ppt, titulo_dinamico_g1)
  my_ppt <- add_centered_image(my_ppt, out_g1)
  my_ppt <- add_executive_interpretation(my_ppt, interp_dinamica_g1)
  my_ppt <- add_executive_footer(my_ppt, txt_footer)
}

# 3. DIAPO GRÁFICO 2 (Pruebas)
if (file.exists(out_g2)) {
  my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER)
  my_ppt <- add_executive_banner(my_ppt, titulo_dinamico_g2)
  my_ppt <- add_centered_image(my_ppt, out_g2)
  my_ppt <- add_executive_interpretation(my_ppt, interp_dinamica_g2)
  my_ppt <- add_executive_footer(my_ppt, txt_footer)
}

# 4. DIAPO TABLA PROVINCIA
if (exists("tab_final") && nrow(tab_final) > 0) {
  col_target <- "Total / POSITIVOS"
  interp_prov <- "Distribución de casos positivos."
  
  if (col_target %in% names(tab_final)) {
    vals <- as.numeric(tab_final[[col_target]])
    total_pos <- sum(vals, na.rm = TRUE)
    if(length(vals) > 0) {
      top1_n <- vals[1]
      top1_name <- as.character(tab_final[[1]][1])
      pct_top1 <- if(total_pos > 0) round(100 * top1_n / total_pos, 1) else 0
      interp_prov <- sprintf("La provincia de %s concentra el %.1f%% (%s) de los casos positivos reportados.", top1_name, pct_top1, scales::comma(top1_n))
    }
  }
  
  my_ppt <- add_table_slide_pretty(
    ppt = my_ppt,
    df = tab_final,
    title = "POSITIVOS POR PROVINCIA",
    subtitle = interp_prov,
    prefix = "03_Tabla_Prov"
  )
  my_ppt <- add_executive_footer(my_ppt, txt_footer)
}

# 5. DIAPO TABLA MICRORED
if (exists("tabla_se_final") && nrow(tabla_se_final) > 0) {
  my_ppt <- add_table_slide_pretty(
    ppt = my_ppt,
    df = tabla_se_final,
    title = "DETALLE POR MICRORED Y EE.SS.",
    subtitle = sprintf("Desglose detallado SE %02d (Ordenado por mayor carga)", max_se_g1),
    prefix = "04_Tabla_Microred"
  )
  my_ppt <- add_executive_footer(my_ppt, txt_footer)
}

# --- GUARDAR PPT FINAL ---
nombre_ppt <- file.path(carpeta, sprintf("Reporte_Dengue_SE%02d_Ejecutivo_V2.pptx", max_se_g1))
print(my_ppt, target = nombre_ppt)
message("PPT generado correctamente: ", nombre_ppt)