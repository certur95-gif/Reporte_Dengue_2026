# ========================================================== # Encabezado del script.
# REPORTE SEMANAL DENGUE - NETLABv2 (OPTIMIZADO) [CORREGIDO] # Título descriptivo del reporte.
# - SE: MMWR (Domingo–Sábado) usando epiweek/epiyear         # Regla de semana epidemiológica.
# - Filtros: Solo 2 exámenes dengue (NS1 e IgM) + RESULTADO VERIFICADO # Filtros principales.
# - Salidas: 2 gráficos + 2 tablas Excel formateadas         # Artefactos esperados.
# ========================================================== # Separador visual.

rm(list = ls()) # Limpiar el entorno para evitar variables residuales.

# --------------------------- # Separador de sección.
# 0) PAQUETES # Título de la sección de paquetes.
# --------------------------- # Separador visual.
req_pkgs <- c( # Vector de paquetes requeridos.
  "readxl", "dplyr", "stringr", "lubridate", "tidyr", "ggplot2", # Paquetes para lectura y manipulación.
  "patchwork", "gridExtra", "scales", "openxlsx", "tibble", # Paquetes para gráficos y salida Excel.
  "officer", "flextable", "webshot2" # Paquetes para PPT y tablas.
) # Fin del vector de paquetes.

# --------------------------- # Separador de sección.
# 1) CONFIGURACIÓN (EDITA SOLO ESTO) # Título de configuración editable.
# --------------------------- # Separador visual.

# Archivo / hoja (ajusta a tu descarga NETLAB) # Nota de configuración.
archivo <- "dengue_2025.xlsx" # Ruta al archivo fuente.
hoja <- 1 # Índice o nombre de hoja en Excel.

# Exámenes permitidos (SOLO estos dos) # Lista de exámenes válidos.
examenes_permitidos <- c( # Definición de exámenes permitidos.
  "Virus Dengue Ag NS1 [Presencia] en Suero o Plasma por Inmunoensayo", # Examen NS1.
  "Virus Dengue Ac. IgM [Presencia] en Suero o Plasma por Inmunoensayo" # Examen IgM.
) # Fin del vector de exámenes.

# Unidad global (por defecto: procesamiento total) # Definición de unidad global.
unidad_global <- "examen" # "examen" = procesamiento total (default).
# unidad_global <- "muestra"  # "muestra" = muestra única (dedup por código).

# Filtro opcional por Laboratorio Destino (GLOBAL): # Configuración de filtro global.
lab_destino_global <- "TODOS" # Valor por defecto (sin filtrar).
# lab_destino_global <- c("LABORATORIO DE REFERENCIA REGIONAL DE LORETO")  # Descomenta para filtrar.

# Semana epidemiológica: MMWR (Domingo–Sábado) # Configuración de sistema de semanas.
week_system <- "MMWR" # "MMWR" o "ISO".

# Carpeta del reporte: por defecto "SE XX" # Configuración de carpeta de salida.
incluir_anio_en_carpeta <- FALSE # Si TRUE incluye año en el nombre.

# ------------- GRÁFICO 1 (IP% por SE) ------------- # Sección de gráfico 1.
g1_anio <- "AUTO" # Año para gráfico 1 o "AUTO".
g1_se_inicio <- 20 # Semana de inicio del gráfico 1.
# g1_se_inicio <- NULL # Usar NULL para no filtrar por semana de inicio.
g1_unidad <- unidad_global # Unidad a utilizar en gráfico 1.

# ------------- GRÁFICO 2 (Procesamiento por tipo de prueba) ------------- # Sección de gráfico 2.
g2_anio <- "AUTO" # Año para gráfico 2 o "AUTO".
g2_se_inicio <- NULL # Semana inicial para gráfico 2.
g2_se_fin <- NULL # Semana final para gráfico 2.
g2_unidad <- "examen" # Unidad para gráfico 2.

# Default: excluir LRNMZVIR # Laboratorios a excluir por defecto.
g2_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES") # Labs excluidos.

g2_lab_solo <- "TODOS" # Filtrar por labs específicos o "TODOS".
# g2_lab_solo <- c("LABORATORIO DE REFERENCIA REGIONAL DE LORETO") # Ejemplo de filtro.

# ------------- TABLA POSITIVOS por PROVINCIA ------------- # Sección tabla provincia.
tabprov_anio <- "AUTO" # Año para tabla provincia.
tabprov_se_inicio <- NULL # Semana inicial para tabla provincia.
tabprov_se_fin <- NULL # Semana final para tabla provincia.
tabprov_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES") # Labs excluidos.
tabprov_lab_solo <- "TODOS" # Filtrar por labs específicos o "TODOS".
tabprov_unidad <- "examen" # Unidad para tabla provincia.

# ------------- TABLA POR SE (MicroRED x EE.SS. Origen) ------------- # Sección tabla SE.
tabse_anio <- 2025 # Año a filtrar para tabla SE.
tabse_se <- "AUTO" # Semana a filtrar para tabla SE.
tabse_excluir_labs <- c("LRNMZVIR - LRN DE METAXENICAS Y ZOONOSIS VIRALES") # Labs excluidos.
tabse_lab_solo <- "TODOS" # Filtrar por labs específicos o "TODOS".
tabse_unidad <- "examen" # Unidad para tabla SE.

# --------------------------- # Separador de sección.
# 2) FUNCIONES (NO EDITAR) # Definición de funciones base.
# --------------------------- # Separador visual.

load_packages <- function(packages) { # Función para instalar y cargar paquetes.
  to_install <- packages[!sapply(packages, requireNamespace, quietly = TRUE)] # Detectar faltantes.
  if (length(to_install) > 0) { # Verificar si hay paquetes faltantes.
    install.packages(to_install) # Instalar paquetes faltantes.
  } # Fin de la condición de instalación.
  invisible(lapply(packages, library, character.only = TRUE)) # Cargar paquetes sin imprimir salida.
  message("Librerías cargadas correctamente. Puedes continuar con la Sección 1.") # Mensaje de confirmación.
} # Fin de la función load_packages.

validate_inputs <- function(archivo, hoja, week_system, unidad_global, unidades_validas) { # Función para validar inputs.
  if (!file.exists(archivo)) { # Validar existencia del archivo.
    stop("No se encontró el archivo de entrada: ", archivo) # Error claro si falta el archivo.
  } # Fin de validación de archivo.
  sheets <- readxl::excel_sheets(archivo) # Obtener nombres de hojas disponibles.
  if (is.numeric(hoja)) { # Validar si hoja es numérica.
    if (hoja < 1 || hoja > length(sheets)) { # Comprobar rango válido.
      stop("La hoja indicada está fuera de rango. Hojas disponibles: ", paste(sheets, collapse = ", ")) # Mensaje de error.
    } # Fin de la validación de rango de hoja.
  } else if (!hoja %in% sheets) { # Validar si hoja por nombre existe.
    stop("No se encontró la hoja '", hoja, "'. Hojas disponibles: ", paste(sheets, collapse = ", ")) # Error claro.
  } # Fin de validación de hoja.
  if (!week_system %in% c("MMWR", "ISO")) { # Validar sistema de semanas.
    stop("week_system debe ser 'MMWR' o 'ISO'. Valor recibido: ", week_system) # Mensaje de error.
  } # Fin de validación de week_system.
  if (!unidad_global %in% unidades_validas) { # Validar unidad global.
    stop("unidad_global debe ser 'examen' o 'muestra'. Valor recibido: ", unidad_global) # Mensaje de error.
  } # Fin de validación de unidad global.
} # Fin de la función validate_inputs.

norm_txt <- function(x) { # Normaliza texto para comparaciones.
  x <- as.character(x) # Convertir a carácter.
  x <- iconv(x, from = "", to = "ASCII//TRANSLIT") # Eliminar tildes.
  x <- tolower(x) # Convertir a minúsculas.
  stringr::str_squish(x) # Eliminar espacios duplicados.
} # Fin de norm_txt.

find_col <- function(df, patterns) { # Buscar columna por patrones.
  nms <- norm_txt(names(df)) # Normalizar nombres de columnas.
  pats <- norm_txt(patterns) # Normalizar patrones.
  for (p in pats) { # Iterar por patrones.
    idx <- which(stringr::str_detect(nms, fixed(p))) # Buscar coincidencias.
    if (length(idx) > 0) return(names(df)[idx[1]]) # Retornar primera coincidencia.
  } # Fin de iteración.
  NA_character_ # Retornar NA si no hay coincidencias.
} # Fin de find_col.

must_col <- function(df, patterns, label = "columna") { # Obtener columna requerida.
  col <- find_col(df, patterns) # Buscar columna con patrones.
  if (is.na(col)) { # Validar que exista.
    stop("No se encontró ", label, ". Patrones buscados: ", # Mensaje base del error.
         paste(patterns, collapse = " | "), # Listar patrones.
         "\nSugerencia: revisa names(raw).") # Sugerencia adicional.
  } # Fin de validación.
  col # Retornar nombre de columna.
} # Fin de must_col.

parse_excel_date <- function(x) { # Parsear fecha desde Excel.
  if (inherits(x, "Date")) return(as.Date(x)) # Devolver si ya es Date.
  if (inherits(x, "POSIXct") || inherits(x, "POSIXt")) return(as.Date(x)) # Convertir datetime.
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30")) # Convertir fecha Excel numérica.
  as.Date(lubridate::parse_date_time( # Parsear texto con varios formatos.
    as.character(x), # Convertir a carácter.
    orders = c("dmy", "dmY", "ymd", "Ymd", "d/m/Y", "Y-m-d", "d-m-Y", "Y/m/d"), # Formatos permitidos.
    tz = "UTC" # Zona horaria fija.
  )) # Fin del parseo.
} # Fin de parse_excel_date.

add_epi <- function(df, date_col, week_system = c("MMWR", "ISO")) { # Agregar SE y año.
  week_system <- match.arg(week_system) # Validar argumento.
  if (week_system == "ISO") { # Lógica para ISO.
    df %>% mutate(se = lubridate::isoweek(.data[[date_col]]), # Semana ISO.
                  anio = lubridate::isoyear(.data[[date_col]])) # Año ISO.
  } else { # Lógica para MMWR.
    df %>% mutate(se = lubridate::epiweek(.data[[date_col]]), # Semana MMWR.
                  anio = lubridate::epiyear(.data[[date_col]])) # Año MMWR.
  } # Fin de condición.
} # Fin de add_epi.

apply_filters <- function(df, anio_sel = NULL, se_inicio = NULL, se_fin = NULL, labs_excluir = NULL, labs_solo = "TODOS") { # Filtro global.
  out <- df # Copia de trabajo.
  if (!is.null(labs_excluir) && length(labs_excluir) > 0) { # Excluir labs.
    ex <- stringr::str_to_upper(stringr::str_squish(labs_excluir)) # Normalizar labs excluidos.
    out <- out %>% dplyr::filter(!(.data$lab_destino_std %in% ex)) # Aplicar filtro.
  } # Fin de exclusión.
  if (!(length(labs_solo) == 1 && toupper(labs_solo) == "TODOS")) { # Incluir solo labs.
    inc <- stringr::str_to_upper(stringr::str_squish(labs_solo)) # Normalizar labs incluidos.
    out <- out %>% dplyr::filter(.data$lab_destino_std %in% inc) # Aplicar filtro.
  } # Fin de inclusión.
  if (!is.null(anio_sel)) { # Filtrar por año.
    out <- out %>% dplyr::filter(.data$anio %in% anio_sel) # Aplicar filtro por año.
  } # Fin de filtro de año.
  if (!is.null(se_inicio)) out <- out %>% dplyr::filter(.data$se >= se_inicio) # Filtrar por semana inicio.
  if (!is.null(se_fin)) out <- out %>% dplyr::filter(.data$se <= se_fin) # Filtrar por semana fin.
  out # Devolver resultado filtrado.
} # Fin de apply_filters.

dedup_by_unit <- function(df, unit = c("examen", "muestra")) { # Dedupe por unidad.
  unit <- match.arg(unit) # Validar unidad.
  if (unit == "examen") return(df) # Si es examen no se deduplica.
  df %>% # Iniciar pipeline.
    group_by(cod_muestra, anio, se) %>% # Agrupar por muestra, año y semana.
    summarise( # Resumir por grupo.
      clasif = case_when( # Clasificación agregada.
        any(clasif == "POSITIVO") ~ "POSITIVO", # Priorizar positivos.
        all(clasif == "NEGATIVO") ~ "NEGATIVO", # Si todos negativos.
        any(clasif == "INDETERMINADO") ~ "INDETERMINADO", # Si hay indeterminados.
        TRUE ~ "OTRO" # Caso por defecto.
      ), # Fin de case_when.
      .groups = "drop" # Desagrupar.
    ) # Fin de summarise.
} # Fin de dedup_by_unit.

style_xlsx <- function(df, path, pct_cols = character(), total_row_idx = NA_integer_) { # Formatear Excel.
  wb <- openxlsx::createWorkbook() # Crear workbook.
  openxlsx::addWorksheet(wb, "Reporte") # Agregar hoja.
  openxlsx::writeData(wb, "Reporte", df, withFilter = TRUE) # Escribir datos.
  headerStyle <- openxlsx::createStyle( # Estilo de encabezado.
    fontSize = 11, fontColour = "#000000", textDecoration = "bold", # Configuración de fuente.
    fgFill = "#D9E1F2", halign = "center", valign = "center", border = "Bottom" # Estilo de celdas.
  ) # Fin del estilo.
  openxlsx::addStyle(wb, "Reporte", headerStyle, rows = 1, cols = 1:ncol(df), gridExpand = TRUE) # Aplicar estilo.
  openxlsx::freezePane(wb, "Reporte", firstRow = TRUE) # Congelar encabezado.
  openxlsx::setColWidths(wb, "Reporte", cols = 1:ncol(df), widths = "auto") # Autoajuste de columnas.
  if (length(pct_cols) > 0) { # Verificar columnas con porcentaje.
    pctStyle <- openxlsx::createStyle(numFmt = "0.00\"%\"") # Estilo de porcentaje.
    for (cn in pct_cols) { # Iterar por columnas de porcentaje.
      if (cn %in% names(df)) { # Verificar existencia de columna.
        col_i <- which(names(df) == cn) # Índice de columna.
        openxlsx::addStyle(wb, "Reporte", pctStyle, rows = 2:(nrow(df) + 1), cols = col_i, gridExpand = TRUE, stack = TRUE) # Aplicar estilo.
      } # Fin de condición de columna.
    } # Fin del bucle.
  } # Fin del bloque de porcentajes.
  if (!is.na(total_row_idx)) { # Aplicar estilo de totales.
    totalStyle <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9E1F2", border = "Top") # Estilo total.
    openxlsx::addStyle(wb, "Reporte", totalStyle, rows = total_row_idx + 1, cols = 1:ncol(df), gridExpand = TRUE, stack = TRUE) # Aplicar estilo.
  } # Fin del bloque de totales.
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE) # Guardar archivo.
  path # Retornar ruta.
} # Fin de style_xlsx.

resolve_auto <- function(x, auto_value) { # Resolver valores AUTO.
  if (is.character(x) && length(x) == 1 && toupper(x) == "AUTO") return(auto_value) # Reemplazar por auto_value.
  x # Retornar valor original.
} # Fin de resolve_auto.

load_raw_data <- function(archivo, hoja) { # Función para leer Excel.
  readxl::read_excel(archivo, sheet = hoja) # Leer datos de Excel.
} # Fin de load_raw_data.

detect_columns <- function(raw) { # Detectar columnas requeridas y opcionales.
  col_fecha <- must_col(raw, c("Fecha Colección", "Fecha Coleccion"), "Fecha de colección") # Columna fecha.
  col_fecha_verif <- must_col(raw, c("Fecha Verificación", "Fecha Verificacion"), "Fecha de verificación") # Columna fecha verificación.
  col_result <- must_col(raw, c("resultado"), "Resultado") # Columna resultado.
  col_examen <- must_col(raw, c("nombre de examen", "examen"), "Examen") # Columna examen.
  col_estatus <- must_col(raw, c("estatus resultado", "estado resultado", "estatus"), "Estatus") # Columna estatus.
  col_cod <- must_col(raw, c("codigo de muestra", "codigo muestra", "cod muestra", "muestra"), "Código de muestra") # Columna código.
  col_labdest <- must_col(raw, c("laboratorio destino", "lab destino", "destino"), "Laboratorio destino") # Columna lab destino.
  col_prov <- find_col(raw, c("provincia", "provincia procedencia", "provincia de procedencia", "provincia domicilio")) # Columna provincia.
  col_micro <- find_col(raw, c("Micro Red EE.SS Origen", "Micro Red", "MicroRED", "Microred")) # Columna micro red.
  col_estab <- find_col(raw, c("Establecimiento de Origen", "IPRESS Origen", "Origen")) # Columna establecimiento.
  list( # Retornar lista de columnas.
    col_fecha = col_fecha, # Fecha.
    col_fecha_verif = col_fecha_verif, # Fecha verificación.
    col_result = col_result, # Resultado.
    col_examen = col_examen, # Examen.
    col_estatus = col_estatus, # Estatus.
    col_cod = col_cod, # Código de muestra.
    col_labdest = col_labdest, # Laboratorio destino.
    col_prov = col_prov, # Provincia.
    col_micro = col_micro, # Micro red.
    col_estab = col_estab # Establecimiento.
  ) # Fin de lista.
} # Fin de detect_columns.

build_dataset <- function(raw, cols, examenes_permitidos, lab_destino_global) { # Construir dataset filtrado.
  examenes_key <- norm_txt(examenes_permitidos) # Normalizar exámenes permitidos.
  dat <- raw %>% # Iniciar pipeline.
    mutate( # Crear columnas estándar.
      fecha_coleccion = parse_excel_date(.data[[cols$col_fecha]]), # Parsear fecha.
      fecha_verificacion = parse_excel_date(.data[[cols$col_fecha_verif]]), # Parsear fecha verificación.
      resultado_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_result]]))), # Normalizar resultado.
      examen_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_examen]]))), # Normalizar examen.
      examen_key = norm_txt(examen_std), # Normalizar examen para comparación.
      estatus_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_estatus]]))), # Normalizar estatus.
      cod_muestra = stringr::str_squish(as.character(.data[[cols$col_cod]])), # Normalizar código de muestra.
      lab_destino_std = stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_labdest]]))) # Normalizar lab destino.
    ) %>% # Fin de mutate.
    filter(!is.na(fecha_coleccion)) %>% # Filtrar filas con fecha colección válida.
    filter(!is.na(fecha_verificacion)) %>% # Filtrar filas con fecha verificación válida.
    filter(estatus_std == "RESULTADO VERIFICADO") %>% # Filtrar estatus verificado.
    filter(examen_key %in% examenes_key) %>% # Filtrar exámenes permitidos.
    mutate( # Crear clasificación.
      clasif = case_when( # Definir clasif.
        str_detect(resultado_std, "POSITIV") ~ "POSITIVO", # Detectar positivos.
        str_detect(resultado_std, "NEGATIV") ~ "NEGATIVO", # Detectar negativos.
        str_detect(resultado_std, "INDETERMIN") ~ "INDETERMINADO", # Detectar indeterminado.
        TRUE ~ "OTRO" # Caso por defecto.
      ) # Fin de case_when.
    ) # Fin de mutate.
  if (!(length(lab_destino_global) == 1 && toupper(lab_destino_global) == "TODOS")) { # Filtro global opcional.
    lab_ok <- stringr::str_to_upper(stringr::str_squish(lab_destino_global)) # Normalizar labs permitidos.
    dat <- dat %>% filter(lab_destino_std %in% lab_ok) # Aplicar filtro global.
  } # Fin de filtro global.
  if (nrow(dat) == 0) { # Validar que queden datos.
    stop("No quedan registros tras filtros: verificado + exámenes permitidos + lab destino (si aplica).") # Error claro.
  } # Fin de validación.
  dat # Retornar dataset filtrado.
} # Fin de build_dataset.

setup_report_context <- function(dat, week_system, incluir_anio_en_carpeta) { # Configurar contexto del reporte.
  fecha_max <- max(dat$fecha_verificacion, na.rm = TRUE) # Fecha máxima disponible.
  se_reporte <- if (week_system == "ISO") lubridate::isoweek(fecha_max) else lubridate::epiweek(fecha_max) # Semana de reporte.
  anio_rep <- if (week_system == "ISO") lubridate::isoyear(fecha_max) else lubridate::epiyear(fecha_max) # Año del reporte.
  carpeta <- if (incluir_anio_en_carpeta) sprintf("%d_SE %02d", anio_rep, se_reporte) else sprintf("SE %02d", se_reporte) # Nombre de carpeta.
  dir.create(carpeta, showWarnings = FALSE, recursive = TRUE) # Crear carpeta si no existe.
  list( # Retornar lista de contexto.
    fecha_max = fecha_max, # Fecha máxima.
    se_reporte = se_reporte, # Semana reporte.
    anio_rep = anio_rep, # Año reporte.
    carpeta = carpeta # Carpeta salida.
  ) # Fin de lista.
} # Fin de setup_report_context.

create_graph1 <- function(dat, cols, g1_anio, g1_se_inicio, g1_unidad, se_reporte, carpeta, week_system) { # Generar gráfico 1.
  base_g1 <- dat %>% # Base para gráfico 1.
    filter(clasif %in% c("NEGATIVO", "POSITIVO")) %>% # Filtrar negativos y positivos.
    mutate( # Agregar columnas opcionales.
      provincia = if (!is.na(cols$col_prov)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_prov]]))) else NA_character_, # Provincia.
      micro_red = if (!is.na(cols$col_micro)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_micro]]))) else NA_character_, # Micro red.
      estab_origen = if (!is.na(cols$col_estab)) stringr::str_to_upper(stringr::str_squish(as.character(.data[[cols$col_estab]]))) else NA_character_ # Establecimiento.
    ) # Fin de mutate.
  base_g1 <- if (g1_unidad == "muestra") dedup_by_unit(base_g1, "muestra") else base_g1 # Deduplicar si corresponde.
  sem <- base_g1 %>% # Calcular resumen semanal.
    count(anio, se, clasif, name = "n") %>% # Contar por año, semana y clasificación.
    tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>% # Pivot a columnas.
    mutate(total = NEGATIVO + POSITIVO, IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>% # Calcular total e IP.
    arrange(anio, se) %>% # Ordenar por año y semana.
    group_by(anio) %>% # Agrupar por año.
    tidyr::complete(se = 1:53, fill = list(NEGATIVO = 0, POSITIVO = 0)) %>% # Completar semanas.
    mutate(total = NEGATIVO + POSITIVO, IP = if_else(total > 0, 100 * POSITIVO / total, NA_real_)) %>% # Recalcular tras completar.
    ungroup() # Desagrupar.
  max_anio_g1 <- max(g1_anio, na.rm = TRUE) # Año máximo para gráfico.
  sem_plot <- sem %>% # Filtrar para gráfico.
    filter(anio < max_anio_g1 | (anio == max_anio_g1 & se <= se_reporte)) # Incluir años previos y cortar año actual.
  if (!is.null(g1_se_inicio)) { # Aplicar semana de inicio solo para año actual.
    sem_plot <- sem_plot %>% filter(anio < max_anio_g1 | se >= g1_se_inicio) # Mantener años previos completos.
  } # Fin filtro se inicio.
  if (nrow(sem_plot) == 0) stop("Gráfico 1: el filtro dejó el dataset vacío (revisa año/SE inicio).") # Error si no hay datos.
  sem_plot <- sem_plot %>% # Ordenar para continuidad entre años.
    arrange(anio, se) %>% # Ordenar por año y semana.
    mutate(se_label = sprintf("%02d\n%d", se, anio)) # Etiqueta SE con año.
  bars <- sem_plot %>% # Preparar barras.
    select(se_label, NEGATIVO, POSITIVO) %>% # Seleccionar columnas.
    pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n") # Pasar a formato largo.
  max_count <- max(bars$n, na.rm = TRUE) # Máximo de barras.
  max_ip <- max(sem_plot$IP, na.rm = TRUE) # Máximo de IP.
  scale_factor <- ifelse(is.finite(max_ip) && max_ip > 0, max_count / max_ip, 1) # Factor de escala.
  se_levels <- unique(sem_plot$se_label) # Niveles de semana.
  bars <- bars %>% mutate(se_f = factor(se_label, levels = se_levels)) # Factor para barras.
  sem_plot <- sem_plot %>% mutate(se_f = factor(se_label, levels = se_levels)) # Factor para línea.
  sem_plot <- sem_plot %>% mutate( # Marcar máximo IP.
    es_max_ip = IP == max(IP, na.rm = TRUE), # Indicador de máximo.
    etiqueta_ip = if_else(es_max_ip, sprintf("IP: %.1f%%", IP), NA_character_) # Etiqueta para máximo.
  ) # Fin de mutate.
  color_ip_impacto <- "#00C853" # Color de impacto.
  p1 <- ggplot() + # Construir gráfico principal.
    geom_col(data = bars, aes(x = se_f, y = n, fill = tipo), position = position_dodge(width = 0.8), width = 0.7) + # Barras.
    geom_line(data = sem_plot, aes(x = se_f, y = IP * scale_factor, group = 1), linewidth = 1.5, color = color_ip_impacto) + # Línea IP.
    geom_point(data = sem_plot, aes(x = se_f, y = IP * scale_factor), size = 4, color = color_ip_impacto, fill = "white", shape = 21, stroke = 2) + # Puntos.
    geom_text(data = sem_plot, aes(x = se_f, y = IP * scale_factor, label = round(IP, 1)), vjust = -0.8, size = 5, fontface = "bold", color = color_ip_impacto) + # Etiquetas IP.
    scale_y_continuous( # Ejes.
      name = ifelse(g1_unidad == "examen", "Número de procesamientos", "Número de muestras"), # Eje Y principal.
      sec.axis = sec_axis(~ . / scale_factor, name = "Índice de positividad (IP%)", labels = function(x) paste0(round(x, 1), "%")) # Eje secundario.
    ) + # Fin de escala.
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) + # Colores.
    labs(x = NULL, fill = NULL) + # Etiquetas.
    theme_minimal(base_size = 16) + # Tema base.
    theme(panel.grid.minor = element_blank(), legend.position = "none", axis.text = element_text(face = "bold"), axis.title = element_text(face = "bold"), axis.text.x = element_blank()) # Ajustes.
  dat_tabla <- sem_plot %>% # Tabla inferior.
    select(se_f, NEGATIVO, POSITIVO, IP) %>% # Selección.
    mutate(IP = sprintf("%.1f", IP), NEGATIVO = as.character(NEGATIVO), POSITIVO = as.character(POSITIVO)) %>% # Convertir a texto.
    pivot_longer(cols = c(NEGATIVO, POSITIVO, IP), names_to = "variable", values_to = "valor") %>% # Formato largo.
    mutate(variable = factor(variable, levels = c("IP", "POSITIVO", "NEGATIVO")), color_txt = case_when(variable == "NEGATIVO" ~ "#1F77B4", variable == "POSITIVO" ~ "#D62728", variable == "IP" ~ color_ip_impacto)) # Colores.
  p_tabla <- ggplot(dat_tabla, aes(x = se_f, y = variable)) + # Construir tabla con ggplot.
    geom_tile(aes(fill = variable), alpha = 0.08, color = "white", linewidth = 0.5) + # Fondo.
    geom_text(aes(label = valor, color = variable), size = 4.5, fontface = "bold") + # Texto.
    scale_color_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) + # Colores texto.
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728", "IP" = color_ip_impacto)) + # Colores fondo.
    scale_y_discrete(labels = c("NEGATIVO" = "NEGATIVO (-)", "POSITIVO" = "POSITIVO (-)", "IP" = "IP (%)")) + # Etiquetas eje Y.
    labs(x = "Semana Epidemiológica") + # Etiqueta eje X.
    theme_minimal(base_size = 14) + # Tema base.
    theme(panel.grid = element_blank(), legend.position = "none", axis.title.y = element_blank(), axis.text.y = element_text(face = "bold", color = "black", hjust = 1), axis.text.x = element_text(face = "bold"), plot.margin = margin(t = -5, r = 0, b = 0, l = 0)) # Ajustes de tema.
  fig1 <- p1 / p_tabla + patchwork::plot_layout(heights = c(4, 1.3)) # Combinar gráficos.
  out_g1 <- file.path(carpeta, "01_IP_por_SE_HighImpact.png") # Ruta de salida.
  ggsave(out_g1, fig1, width = 18, height = 10, dpi = 300) # Guardar gráfico.
  list(sem_plot = sem_plot, out_g1 = out_g1) # Retornar sem_plot y ruta.
} # Fin de create_graph1.

create_graph2 <- function(dat, g2_anio, g2_se_inicio, g2_se_fin, g2_excluir_labs, g2_lab_solo, g2_unidad, carpeta) { # Generar gráfico 2.
  dat_g2 <- dat %>% # Datos para gráfico 2.
    mutate( # Crear tipo de prueba.
      prueba = case_when( # Clasificar pruebas.
        str_detect(examen_key, "ac\\.?\\s*igm") ~ "Virus Dengue Ac. IgM", # IgM.
        str_detect(examen_key, "ag\\s*ns1") ~ "Virus Dengue Ag NS1", # NS1.
        TRUE ~ examen_std # Por defecto.
      ) # Fin de case_when.
    ) %>% # Fin de mutate.
    filter(clasif %in% c("NEGATIVO", "POSITIVO")) %>% # Filtrar negativos y positivos.
    mutate( # Usar semana/año de verificación para filtros del gráfico 2.
      se = se_verif, # SE según fecha de verificación.
      anio = anio_verif # Año según fecha de verificación.
    ) # Fin de mutate.
  anio_f_g2 <- if (is.null(g2_anio)) NULL else g2_anio # Definir año.
  dat_g2 <- apply_filters(dat_g2, anio_sel = anio_f_g2, se_inicio = g2_se_inicio, se_fin = g2_se_fin, labs_excluir = g2_excluir_labs, labs_solo = g2_lab_solo) # Aplicar filtros.
  if (nrow(dat_g2) == 0) stop("Gráfico 2: no quedan datos tras filtros (año/SE/labs).") # Error si no hay datos.
  base_g2 <- if (g2_unidad == "muestra") dedup_by_unit(dat_g2, "muestra") else dat_g2 # Dedupe si aplica.
  res_g2 <- base_g2 %>% # Resumen por prueba.
    count(prueba, clasif, name = "n") %>% # Conteo por prueba/clasif.
    tidyr::pivot_wider(names_from = clasif, values_from = n, values_fill = 0) %>% # Pivot a columnas.
    mutate(TOTAL = NEGATIVO + POSITIVO) # Total por prueba.
  bars_g2 <- res_g2 %>% # Preparar barras.
    select(prueba, NEGATIVO, POSITIVO) %>% # Seleccionar columnas.
    pivot_longer(cols = c(NEGATIVO, POSITIVO), names_to = "tipo", values_to = "n") # Formato largo.
  ymax <- max(res_g2$TOTAL, na.rm = TRUE) # Máximo para texto total.
  offset_total <- ymax * 0.05 # Offset para etiqueta total.
  min_se_g2 <- min(dat_g2$se, na.rm = TRUE) # Semana mínima.
  max_se_g2 <- max(dat_g2$se, na.rm = TRUE) # Semana máxima.
  subtitulo_dinamico <- sprintf("Desde SE %02d hasta la SE %02d %s", min_se_g2, max_se_g2, g2_anio) # Subtítulo.
  p2 <- ggplot() + # Construir gráfico.
    geom_col(data = bars_g2, aes(x = prueba, y = n, fill = tipo), position = position_dodge(width = 0.8), width = 0.7) + # Barras.
    geom_text(data = bars_g2, aes(x = prueba, y = n, label = scales::comma(n), group = tipo), position = position_dodge(width = 0.8), vjust = -0.5, size = 6, fontface = "bold") + # Etiquetas.
    geom_text(data = res_g2, aes(x = prueba, y = TOTAL + offset_total, label = paste0("TOTAL:\n", scales::comma(TOTAL))), size = 8, fontface = "bold", color = "black", lineheight = 0.8) + # Etiqueta total.
    scale_fill_manual(values = c("NEGATIVO" = "#1F77B4", "POSITIVO" = "#D62728")) + # Colores.
    scale_y_continuous(labels = scales::comma, expand = expansion(mult = c(0, 0.2))) + # Escala Y.
    labs(x = NULL, y = "NÚMERO DE PRUEBAS", fill = NULL, title = "PROCESAMIENTO DE MUESTRAS POR TIPO DE PRUEBA", subtitle = subtitulo_dinamico) + # Etiquetas.
    theme_minimal(base_size = 18) + # Tema base.
    theme(legend.position = "top", panel.grid.major.x = element_blank(), panel.grid.minor = element_blank(), axis.text.x = element_text(face = "bold", size = 14), axis.text.y = element_text(face = "bold"), plot.title = element_text(face = "bold", hjust = 0.5, size = 22), plot.subtitle = element_text(face = "italic", hjust = 0.5, size = 16, color = "#555555")) # Ajustes.
  out_g2 <- file.path(carpeta, "02_Procesamiento_por_prueba_HighImpact.png") # Ruta de salida.
  ggsave(out_g2, p2, width = 16, height = 9, dpi = 300) # Guardar gráfico.
  list(out_g2 = out_g2, res_g2 = res_g2, dat_g2 = dat_g2) # Retornar resultados.
} # Fin de create_graph2.

create_table_prov <- function(dat, col_prov, tabprov_anio, tabprov_se_inicio, tabprov_se_fin, tabprov_excluir_labs, tabprov_lab_solo, tabprov_unidad, carpeta) { # Crear tabla provincia.
  if (is.na(col_prov)) { # Validar columna provincia.
    warning("No se detectó columna PROVINCIA. Se omite tabla por provincia.") # Advertencia.
    return(NULL) # Retornar NULL.
  } # Fin de validación.
  dat_tabprov <- dat %>% # Datos para tabla provincia.
    mutate( # Crear columnas.
      provincia = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_prov]]))), # Normalizar provincia.
      provincia = if_else(is.na(provincia) | provincia == "", "SIN DATO", provincia), # Reemplazar vacíos.
      prueba = case_when(str_detect(examen_key, "ac\\.?\\s*igm") ~ "IgM", str_detect(examen_key, "ag\\s*ns1") ~ "NS1", TRUE ~ "OTRO") # Clasificar prueba.
    ) %>% # Fin de mutate.
    filter(prueba %in% c("IgM", "NS1")) # Filtrar pruebas.
  anio_f_tabprov <- if (is.null(tabprov_anio)) NULL else tabprov_anio # Definir año.
  dat_tabprov <- apply_filters(dat_tabprov, anio_sel = anio_f_tabprov, se_inicio = tabprov_se_inicio, se_fin = tabprov_se_fin, labs_excluir = tabprov_excluir_labs, labs_solo = tabprov_lab_solo) # Aplicar filtros.
  if (nrow(dat_tabprov) == 0) stop("Tabla provincia: no quedan datos tras filtros.") # Error si no hay datos.
  base_pos <- dat_tabprov %>% filter(clasif == "POSITIVO") # Base de positivos.
  if (tabprov_unidad == "muestra") { # Dedupe por muestra.
    base_pos <- dat_tabprov %>% # Recalcular positivos por muestra.
      group_by(cod_muestra, provincia, prueba) %>% # Agrupar por muestra.
      summarise(clasif = if_else(any(clasif == "POSITIVO"), "POSITIVO", "NEGATIVO"), .groups = "drop") %>% # Resumir.
      filter(clasif == "POSITIVO") %>% # Mantener positivos.
      select(provincia, prueba) # Seleccionar columnas.
  } else { # Si unidad examen.
    base_pos <- base_pos %>% select(provincia, prueba) # Seleccionar columnas.
  } # Fin de condición.
  tab_pos <- base_pos %>% # Construir tabla.
    count(provincia, prueba, name = "n") %>% # Contar.
    tidyr::pivot_wider(names_from = prueba, values_from = n, values_fill = 0) %>% # Pivot.
    mutate(`Total / POSITIVOS` = IgM + NS1, `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0), `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0)) %>% # Calcular totales y porcentajes.
    arrange(desc(`Total / POSITIVOS`)) # Ordenar.
  fila_total <- tibble::tibble(PROVINCIA = "Total", IgM = sum(tab_pos$IgM, na.rm = TRUE), NS1 = sum(tab_pos$NS1, na.rm = TRUE)) %>% # Fila total.
    mutate(`Total / POSITIVOS` = IgM + NS1, `% IgM` = if_else(`Total / POSITIVOS` > 0, 100 * IgM / `Total / POSITIVOS`, 0), `% NS1` = if_else(`Total / POSITIVOS` > 0, 100 * NS1 / `Total / POSITIVOS`, 0)) # Calcular porcentajes.
  tab_final <- tab_pos %>% # Tabla final.
    transmute(PROVINCIA = provincia, IgM, NS1, `% IgM`, `% NS1`, `Total / POSITIVOS`) %>% # Seleccionar columnas.
    bind_rows(fila_total) # Agregar fila total.
  out_tabprov <- file.path(carpeta, "03_Tabla_positivos_por_provincia.xlsx") # Ruta de salida.
  style_xlsx(tab_final, out_tabprov, pct_cols = c("% IgM", "% NS1"), total_row_idx = nrow(tab_final)) # Guardar con estilo.
  tab_final # Retornar tabla final.
} # Fin de create_table_prov.

create_table_se <- function(dat, col_micro, col_estab, tabse_anio, tabse_se, tabse_excluir_labs, tabse_lab_solo, tabse_unidad, carpeta) { # Crear tabla por SE.
  if (is.na(col_micro) || is.na(col_estab)) { # Validar columnas.
    warning("No se detectó Micro RED y/o Establecimiento de Origen. Se omite tabla por SE.") # Advertencia.
    return(NULL) # Retornar NULL.
  } # Fin de validación.
  dat_tabse <- dat %>% # Datos para tabla SE.
    mutate( # Normalizar campos.
      micro_red = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_micro]]))), # Micro red.
      estab_origen = stringr::str_to_upper(stringr::str_squish(as.character(.data[[col_estab]]))), # Establecimiento.
      micro_red = if_else(is.na(micro_red) | micro_red == "", "SIN DATO", micro_red), # Reemplazar vacíos.
      estab_origen = if_else(is.na(estab_origen) | estab_origen == "", "SIN DATO", estab_origen) # Reemplazar vacíos.
    ) %>% # Fin de mutate.
    filter(clasif %in% c("NEGATIVO", "POSITIVO")) # Filtrar negativos y positivos.
  anio_f_tabse <- if (is.null(tabse_anio)) NULL else tabse_anio # Definir año.
  dat_tabse <- apply_filters(dat_tabse, anio_sel = anio_f_tabse, se_inicio = tabse_se, se_fin = tabse_se, labs_excluir = tabse_excluir_labs, labs_solo = tabse_lab_solo) # Aplicar filtros.
  if (nrow(dat_tabse) == 0) stop("Tabla SE: no hay registros para ese año/SE con los filtros actuales.") # Error si no hay datos.
  if (tabse_unidad == "muestra") { # Dedupe por muestra.
    dat_tabse <- dedup_by_unit(dat_tabse, "muestra") %>% filter(clasif %in% c("NEGATIVO", "POSITIVO")) # Mantener clasif.
  } # Fin de dedupe.
  tabla_se <- dat_tabse %>% # Construir tabla.
    group_by(micro_red, estab_origen) %>% # Agrupar.
    summarise(`NEGATIVO -` = sum(clasif == "NEGATIVO", na.rm = TRUE), `POSITIVO -` = sum(clasif == "POSITIVO", na.rm = TRUE), .groups = "drop") %>% # Resumen.
    mutate(`Total general` = `NEGATIVO -` + `POSITIVO -`, IP = if_else(`Total general` > 0, 100 * `POSITIVO -` / `Total general`, NA_real_), SE = sprintf("SE %02d", as.integer(tabse_se))) %>% # Calcular totales.
    relocate(SE, .before = micro_red) %>% # Mover SE al inicio.
    arrange(desc(`Total general`), micro_red, estab_origen) # Ordenar.
  tot_neg <- sum(tabla_se$`NEGATIVO -`, na.rm = TRUE) # Total negativos.
  tot_pos <- sum(tabla_se$`POSITIVO -`, na.rm = TRUE) # Total positivos.
  tot_all <- sum(tabla_se$`Total general`, na.rm = TRUE) # Total general.
  tot_ip <- ifelse(tot_all > 0, 100 * tot_pos / tot_all, NA_real_) # IP total.
  fila_total <- tibble::tibble(SE = sprintf("SE %02d", as.integer(tabse_se)), micro_red = "Total general", estab_origen = "", `NEGATIVO -` = tot_neg, `POSITIVO -` = tot_pos, `Total general` = tot_all, IP = tot_ip) # Fila total.
  tabla_se_final <- bind_rows(tabla_se, fila_total) %>% rename(`Micro RED` = micro_red, `Establecimiento de Origen` = estab_origen) # Tabla final.
  out_tabse <- file.path(carpeta, sprintf("04_Tabla_SE%02d_Microred_Establecimiento.xlsx", as.integer(tabse_se))) # Ruta de salida.
  style_xlsx(tabla_se_final, out_tabse, pct_cols = c("IP"), total_row_idx = nrow(tabla_se_final)) # Guardar tabla.
  tabla_se_final # Retornar tabla final.
} # Fin de create_table_se.

generate_ppt <- function(carpeta, out_g1, out_g2, sem_plot, dat, res_g2, tab_final, tabla_se_final) { # Generar PPT.
  suppressPackageStartupMessages({ # Suprimir mensajes.
    library(officer) # Cargar officer.
    library(dplyr) # Cargar dplyr.
    library(flextable) # Cargar flextable.
    library(webshot2) # Cargar webshot2.
  }) # Fin de suppressPackageStartupMessages.
  message("Generando PPT con tablas de alto impacto...") # Mensaje de inicio.
  if (!exists("carpeta")) carpeta <- getwd() # Asegurar carpeta.
  dir.create(carpeta, showWarnings = FALSE, recursive = TRUE) # Crear carpeta si no existe.
  ruta_plantilla <- "plantilla_base.pptx" # Ruta de plantilla.
  my_ppt <- if (file.exists(ruta_plantilla)) read_pptx(ruta_plantilla) else read_pptx() # Cargar plantilla.
  ls <- layout_summary(my_ppt) # Resumen de layouts.
  candidatos_layout <- c("Blank", "En blanco", "Vacío", "Vacio", "Title Only", "Título solamente") # Layouts preferidos.
  LAYOUT <- candidatos_layout[candidatos_layout %in% ls$layout][1] # Escoger layout.
  if (is.na(LAYOUT)) LAYOUT <- ls$layout[1] # Fallback layout.
  MASTER <- ls$master[ls$layout == LAYOUT][1] # Seleccionar master.
  if (is.na(MASTER)) MASTER <- ls$master[1] # Fallback master.
  dims <- slide_size(my_ppt) # Tamaño de diapositiva.
  W_SLIDE <- dims$width # Ancho.
  H_SLIDE <- dims$height # Alto.
  H_BANNER_GRAF <- 0.75 # Alto de banner gráfico.
  H_BANNER_TAB <- 0.85 # Alto de banner tabla.
  H_FOOTER <- 0.45 # Alto footer.
  PAD <- 0.12 # Padding.
  COLOR_FONDO_TITULO <- "#D9E1F2" # Color fondo título.
  COLOR_FONDO_FOOTER <- "#F2F2F2" # Color fondo footer.
  add_rect <- function(ppt, left, top, width, height, fill) { # Agregar rectángulo.
    ppt <- ph_with(ppt, value = fpar(ftext(" ", fp_text(font.size = 1))), location = ph_location(left = left, top = top, width = width, height = height, bg = fill)) # Dibujar rectángulo.
    ppt # Retornar ppt.
  } # Fin de add_rect.
  add_banner2 <- function(ppt, title, subtitle = NULL, h = 0.75) { # Agregar banner.
    ppt <- add_rect(ppt, 0, 0, W_SLIDE, h, COLOR_FONDO_TITULO) # Fondo banner.
    ppt <- ph_with(ppt, value = fpar(ftext(title, prop = fp_text(font.size = 20, bold = TRUE, color = "black", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = PAD, top = 0.05, width = W_SLIDE - 2 * PAD, height = h * 0.6)) # Título.
    if (!is.null(subtitle) && nchar(trimws(subtitle)) > 0) { # Subtítulo si aplica.
      ppt <- ph_with(ppt, value = fpar(ftext(subtitle, prop = fp_text(font.size = 12, italic = FALSE, color = "#333333", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = PAD, top = h * 0.55, width = W_SLIDE - 2 * PAD, height = h * 0.4)) # Subtítulo.
    } # Fin de condición.
    ppt # Retornar ppt.
  } # Fin de add_banner2.
  add_footer <- function(ppt, texto) { # Agregar footer.
    ppt <- add_rect(ppt, 0, H_SLIDE - H_FOOTER, W_SLIDE, H_FOOTER, COLOR_FONDO_FOOTER) # Fondo footer.
    ppt <- ph_with(ppt, value = fpar(ftext(texto, prop = fp_text(font.size = 10, italic = TRUE, color = "#555555", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = PAD, top = H_SLIDE - H_FOOTER + 0.1, width = W_SLIDE - 2 * PAD, height = H_FOOTER - 0.1)) # Texto footer.
    ppt # Retornar ppt.
  } # Fin de add_footer.
  add_img_in_body <- function(ppt, img_path, banner_h, footer = TRUE) { # Agregar imagen.
    if (!file.exists(img_path)) return(ppt) # Validar imagen.
    h_body <- H_SLIDE - banner_h - if (footer) H_FOOTER else 0 # Alto del cuerpo.
    ph_with(ppt, value = external_img(img_path), location = ph_location(left = 0, top = banner_h, width = W_SLIDE, height = h_body)) # Insertar imagen.
  } # Fin de add_img_in_body.
  df_to_flextable_img <- function(df, path, width_in, height_in, base_size = 11) { # Convertir dataframe a imagen.
    ft <- flextable(df) %>% theme_vanilla() %>% font(fontname = "Arial", part = "all") %>% fontsize(size = base_size, part = "all") %>% align(align = "center", part = "all") %>% align(j = 1, align = "left", part = "all") %>% autofit() # Construir flextable.
    ft <- ft %>% bg(part = "header", bg = "#1F4E79") %>% color(part = "header", color = "white") %>% bold(part = "header") %>% border_remove() %>% hline_bottom(part = "header", border = fp_border(color = "white", width = 2)) # Estilo header.
    ft <- ft %>% bg(i = seq(1, nrow(df), 2), bg = "#F2F2F2", part = "body") %>% bg(i = seq(2, nrow(df), 2), bg = "white", part = "body") # Zebra.
    idx_total <- which(str_detect(toupper(as.character(df[[1]])), "TOTAL") | str_detect(toupper(as.character(df[[1]])), "GENERAL")) # Detectar total.
    if (length(idx_total) > 0) { # Estilizar total.
      ft <- ft %>% bold(i = idx_total, part = "body") %>% bg(i = idx_total, bg = "#D9E1F2", part = "body") %>% hline(i = idx_total, border = fp_border(color = "#1F4E79", width = 1.5)) # Estilo total.
    } # Fin de condición.
    if (nrow(df) > 25) ft <- fontsize(ft, size = 9, part = "all") # Reducir fuente si es larga.
    if (nrow(df) > 40) ft <- fontsize(ft, size = 8, part = "all") # Reducir más.
    ft <- padding(ft, padding = 3, part = "all") # Ajustar padding.
    save_as_image(ft, path = path, webshot = "webshot2", zoom = 3) # Guardar imagen.
    path # Retornar ruta.
  } # Fin de df_to_flextable_img.
  add_table_slide_pretty <- function(ppt, df, title, subtitle, prefix, banner_h = 0.85) { # Agregar slide con tabla.
    if (is.null(df) || nrow(df) == 0) return(ppt) # Validar data.
    box_h <- H_SLIDE - banner_h # Alto de caja.
    out_png <- file.path(carpeta, sprintf("%s_pretty.png", prefix)) # Ruta imagen.
    tryCatch({ # Intentar generar imagen.
      df_to_flextable_img(df, out_png, width_in = W_SLIDE, height_in = box_h) # Crear imagen.
    }, error = function(e) { # Manejo de error.
      warning("Error generando flextable: ", e$message) # Advertencia.
    }) # Fin de tryCatch.
    ppt <- add_slide(ppt, layout = LAYOUT, master = MASTER) # Agregar slide.
    ppt <- add_banner2(ppt, title, subtitle, h = banner_h) # Agregar banner.
    if (file.exists(out_png)) { # Insertar imagen si existe.
      ppt <- ph_with(ppt, value = external_img(out_png), location = ph_location(left = 0, top = banner_h, width = W_SLIDE, height = box_h)) # Insertar imagen.
    } # Fin de condición.
    ppt # Retornar ppt.
  } # Fin de add_table_slide_pretty.
  get_solid_img <- function(color, width_in, height_in) { # Generar imagen sólida.
    fname <- paste0("solid_", gsub("#", "", color), ".png") # Nombre de archivo.
    fpath <- file.path(tempdir(), fname) # Ruta en temp.
    if (!file.exists(fpath)) { # Crear si no existe.
      png(fpath, width = 100, height = 100) # Abrir dispositivo PNG.
      par(mar = c(0, 0, 0, 0)) # Sin márgenes.
      plot(0, 0, type = "n", xlim = c(0, 1), ylim = c(0, 1), axes = FALSE, xlab = "", ylab = "") # Lienzo.
      rect(0, 0, 1, 1, col = color, border = NA) # Rectángulo sólido.
      dev.off() # Cerrar dispositivo.
    } # Fin de condición.
    fpath # Retornar ruta.
  } # Fin de get_solid_img.
  add_executive_banner <- function(ppt, title) { # Banner ejecutivo.
    bg_img <- get_solid_img("#1F4E79", W_SLIDE, 0.85) # Imagen de fondo.
    ppt <- ph_with(ppt, value = external_img(bg_img), location = ph_location(left = 0, top = 0, width = W_SLIDE, height = 0.85)) # Fondo.
    ppt <- ph_with(ppt, value = fpar(ftext(title, prop = fp_text(font.size = 20, bold = TRUE, color = "white", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = 0.2, top = 0.1, width = W_SLIDE - 0.4, height = 0.7)) # Título.
    ppt # Retornar ppt.
  } # Fin de add_executive_banner.
  add_executive_interpretation <- function(ppt, text) { # Caja interpretación.
    top_pos <- 6.1 # Posición superior.
    height_box <- 1.1 # Alto de caja.
    bg_img <- get_solid_img("#F2F2F2", W_SLIDE - 1.0, height_box) # Fondo gris.
    ppt <- ph_with(ppt, value = external_img(bg_img), location = ph_location(left = 0.5, top = top_pos, width = W_SLIDE - 1.0, height = height_box)) # Fondo.
    ppt <- ph_with(ppt, value = fpar(ftext("INTERPRETACIÓN:", prop = fp_text(font.size = 10, bold = TRUE, color = "#1F4E79", font.family = "Arial")), ftext("\r\n", prop = fp_text(font.size = 5)), ftext(text, prop = fp_text(font.size = 11, color = "#333333", font.family = "Arial"))), location = ph_location(left = 0.6, top = top_pos + 0.1, width = W_SLIDE - 1.2, height = height_box - 0.2)) # Texto.
    ppt # Retornar ppt.
  } # Fin de add_executive_interpretation.
  add_centered_image <- function(ppt, img_path) { # Imagen central.
    if (!file.exists(img_path)) return(ppt) # Validar imagen.
    ph_with(ppt, value = external_img(img_path), location = ph_location(left = 0, top = 0.9, width = W_SLIDE, height = 5.1)) # Insertar imagen.
  } # Fin de add_centered_image.
  add_executive_footer <- function(ppt, text) { # Footer ejecutivo.
    ph_with(ppt, value = fpar(ftext(text, prop = fp_text(font.size = 9, italic = FALSE, color = "#777777", font.family = "Arial")), fp_p = fp_par(text.align = "right")), location = ph_location(left = 0, top = 7.15, width = W_SLIDE - 0.3, height = 0.35)) # Insertar footer.
  } # Fin de add_executive_footer.
  min_se_g1 <- min(sem_plot$se) # Semana mínima.
  max_se_g1 <- max(sem_plot$se) # Semana máxima.
  anio_g1 <- unique(sem_plot$anio)[1] # Año de gráfico 1.
  fecha_corte_raw <- max(dat$fecha_coleccion, na.rm = TRUE) # Fecha corte.
  fecha_corte_txt <- format(fecha_corte_raw, "%d.%m.%Y") # Formato fecha.
  fila_ultima_se <- sem_plot %>% filter(se == max_se_g1) # Última semana.
  ip_actual <- if (nrow(fila_ultima_se) > 0) round(fila_ultima_se$IP, 1) else 0 # IP actual.
  ip_max_rango <- max(sem_plot$IP, na.rm = TRUE) # IP máximo.
  techo_ip <- ceiling(ip_max_rango) # Techo IP.
  titulo_dinamico_g1 <- sprintf("NÚMERO DE MUESTRAS DE DENGUE E INDICE DE POSITIVIDAD (IP) POR SEMANAS EPIDEMIOLÓGICAS. SE %d-%d – SE %d-%d. (CORTE %s)", min_se_g1, anio_g1, max_se_g1, anio_g1, fecha_corte_txt) # Título G1.
  interp_dinamica_g1 <- sprintf("El Índice de Positividad muestra el porcentaje de resultados positivos con respecto al número total de pruebas procesadas. A partir de la SE %d el IP se mantuvo constante menor al %d%%. En la SE %d el IP es %.1f%%.", min_se_g1, techo_ip, max_se_g1, ip_actual) # Interpretación G1.
  titulo_dinamico_g2 <- sprintf("PROCESAMIENTO DE MUESTRAS (NS1 E IGM) POR SEMANAS EPIDEMIOLÓGICAS. SE %d-%d – SE %d-%d. (CORTE %s)", min_se_g1, anio_g1, max_se_g1, anio_g1, fecha_corte_txt) # Título G2.
  total_muestras_g2 <- sum(res_g2$TOTAL, na.rm = TRUE) # Total muestras G2.
  interp_dinamica_g2 <- sprintf("Durante el periodo analizado (SE %d a %d), se procesaron un total de %s muestras. Se observa el comportamiento comparativo entre las pruebas de NS1 e IgM para la vigilancia virológica.", min_se_g1, max_se_g1, scales::comma(total_muestras_g2)) # Interpretación G2.
  txt_footer <- sprintf("Fuente: NETLABv2 | Corte: %s", fecha_corte_txt) # Footer.
  message("Generando diapositivas ejecutivas (con fondo corregido)...") # Mensaje.
  my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER) # Portada.
  my_ppt <- ph_with(my_ppt, value = fpar(ftext(" ", fp_text(font.size = 1))), location = ph_location_fullsize(bg = "white")) # Fondo blanco.
  my_ppt <- ph_with(my_ppt, value = fpar(ftext("VIGILANCIA DE DENGUE", fp_text(font.size = 44, bold = TRUE, color = "#1F4E79", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = 0, top = 2.5, width = W_SLIDE, height = 1.5)) # Título portada.
  subtitulo_portada <- sprintf("SEMANA %02d - %s", max_se_g1, anio_g1) # Subtítulo portada.
  my_ppt <- ph_with(my_ppt, value = fpar(ftext(subtitulo_portada, fp_text(font.size = 24, color = "#666666", font.family = "Arial")), fp_p = fp_par(text.align = "center")), location = ph_location(left = 0, top = 3.8, width = W_SLIDE, height = 1.0)) # Subtítulo.
  my_ppt <- add_executive_footer(my_ppt, txt_footer) # Footer portada.
  if (file.exists(out_g1)) { # Diapo gráfico 1.
    my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER) # Nueva slide.
    my_ppt <- add_executive_banner(my_ppt, titulo_dinamico_g1) # Banner.
    my_ppt <- add_centered_image(my_ppt, out_g1) # Imagen.
    my_ppt <- add_executive_interpretation(my_ppt, interp_dinamica_g1) # Interpretación.
    my_ppt <- add_executive_footer(my_ppt, txt_footer) # Footer.
  } # Fin diapo G1.
  if (file.exists(out_g2)) { # Diapo gráfico 2.
    my_ppt <- add_slide(my_ppt, layout = LAYOUT, master = MASTER) # Nueva slide.
    my_ppt <- add_executive_banner(my_ppt, titulo_dinamico_g2) # Banner.
    my_ppt <- add_centered_image(my_ppt, out_g2) # Imagen.
    my_ppt <- add_executive_interpretation(my_ppt, interp_dinamica_g2) # Interpretación.
    my_ppt <- add_executive_footer(my_ppt, txt_footer) # Footer.
  } # Fin diapo G2.
  if (!is.null(tab_final) && nrow(tab_final) > 0) { # Diapo tabla provincia.
    col_target <- "Total / POSITIVOS" # Columna de referencia.
    interp_prov <- "Distribución de casos positivos." # Texto por defecto.
    if (col_target %in% names(tab_final)) { # Calcular interpretación.
      vals <- as.numeric(tab_final[[col_target]]) # Valores.
      total_pos <- sum(vals, na.rm = TRUE) # Total positivos.
      if (length(vals) > 0) { # Validar longitud.
        top1_n <- vals[1] # Top 1.
        top1_name <- as.character(tab_final[[1]][1]) # Nombre top 1.
        pct_top1 <- if (total_pos > 0) round(100 * top1_n / total_pos, 1) else 0 # Porcentaje.
        interp_prov <- sprintf("La provincia de %s concentra el %.1f%% (%s) de los casos positivos reportados.", top1_name, pct_top1, scales::comma(top1_n)) # Texto.
      } # Fin de condición.
    } # Fin de condición.
    my_ppt <- add_table_slide_pretty(ppt = my_ppt, df = tab_final, title = "POSITIVOS POR PROVINCIA", subtitle = interp_prov, prefix = "03_Tabla_Prov") # Agregar tabla.
    my_ppt <- add_executive_footer(my_ppt, txt_footer) # Footer.
  } # Fin diapo tabla provincia.
  if (!is.null(tabla_se_final) && nrow(tabla_se_final) > 0) { # Diapo tabla microred.
    my_ppt <- add_table_slide_pretty(ppt = my_ppt, df = tabla_se_final, title = "DETALLE POR MICRORED Y EE.SS.", subtitle = sprintf("Desglose detallado SE %02d (Ordenado por mayor carga)", max_se_g1), prefix = "04_Tabla_Microred") # Agregar tabla.
    my_ppt <- add_executive_footer(my_ppt, txt_footer) # Footer.
  } # Fin diapo tabla microred.
  nombre_ppt <- file.path(carpeta, sprintf("Reporte_Dengue_SE%02d_Ejecutivo_V2.pptx", max_se_g1)) # Nombre PPT.
  print(my_ppt, target = nombre_ppt) # Guardar PPT.
  message("PPT generado correctamente: ", nombre_ppt) # Mensaje final.
  nombre_ppt # Retornar ruta PPT.
} # Fin de generate_ppt.

# --------------------------- # Separador de sección.
# 3) EJECUCIÓN PRINCIPAL # Orquestación del flujo.
# --------------------------- # Separador visual.

load_packages(req_pkgs) # Cargar paquetes.
validate_inputs(archivo, hoja, week_system, unidad_global, c("examen", "muestra")) # Validar inputs.

raw <- load_raw_data(archivo, hoja) # Leer datos.
cols <- detect_columns(raw) # Detectar columnas.

# --------------------------- # Separador de sección.
# 4) LIMPIEZA + FILTRO DENGUE # Limpieza de datos.
# --------------------------- # Separador visual.

if (nrow(raw) == 0) stop("El archivo no contiene registros.") # Validación de datos vacíos.

dat <- build_dataset(raw, cols, examenes_permitidos, lab_destino_global) # Construir dataset filtrado.

dat <- add_epi(dat, "fecha_coleccion", week_system = week_system) # Agregar SE y año.
dat <- dat %>% # Agregar SE y año de verificación.
  mutate( # Crear columnas de verificación.
    se_verif = if_else(week_system == "ISO", lubridate::isoweek(fecha_verificacion), lubridate::epiweek(fecha_verificacion)), # SE verificación.
    anio_verif = if_else(week_system == "ISO", lubridate::isoyear(fecha_verificacion), lubridate::epiyear(fecha_verificacion)) # Año verificación.
  ) # Fin de mutate.

context <- setup_report_context(dat, week_system, incluir_anio_en_carpeta) # Configurar carpeta y SE.

se_reporte <- context$se_reporte # Semana de reporte.
anio_rep <- context$anio_rep # Año de reporte.
carpeta <- context$carpeta # Carpeta de salida.

dat <- dat %>% # Filtrar por semana de verificación del reporte.
  filter(se_verif == se_reporte, anio_verif == anio_rep) # Mantener solo verificados en la SE del reporte.

if (nrow(dat) == 0) stop("No hay registros verificados en la SE del reporte (Fecha Verificación).") # Validar datos tras filtro.

se_reporte_coleccion <- max(dat$se, na.rm = TRUE) # Semana máxima según Fecha Colección.
anio_coleccion_rep <- max(dat$anio, na.rm = TRUE) # Año máximo según Fecha Colección.

# Resolver AUTO en configuraciones # Comentario de paso.
g1_anio <- resolve_auto(g1_anio, anio_coleccion_rep) # Resolver año gráfico 1.
g2_anio <- resolve_auto(g2_anio, anio_rep) # Resolver año gráfico 2 (verificación).
tabprov_anio <- resolve_auto(tabprov_anio, anio_coleccion_rep) # Resolver año tabla provincia.
tabse_anio <- resolve_auto(tabse_anio, anio_coleccion_rep) # Resolver año tabla SE.
tabse_se <- resolve_auto(tabse_se, se_reporte_coleccion) # Resolver semana tabla SE.

# --------------------------- # Separador de sección.
# 5) GRÁFICO 1 # Generar gráfico 1.
# --------------------------- # Separador visual.

graph1 <- create_graph1(dat, cols, g1_anio, g1_se_inicio, g1_unidad, se_reporte_coleccion, carpeta, week_system) # Crear gráfico 1.
sem_plot <- graph1$sem_plot # Guardar sem_plot.
out_g1 <- graph1$out_g1 # Ruta gráfico 1.

# --------------------------- # Separador de sección.
# 6) GRÁFICO 2 # Generar gráfico 2.
# --------------------------- # Separador visual.

graph2 <- create_graph2(dat, g2_anio, g2_se_inicio, g2_se_fin, g2_excluir_labs, g2_lab_solo, g2_unidad, carpeta) # Crear gráfico 2.
out_g2 <- graph2$out_g2 # Ruta gráfico 2.
res_g2 <- graph2$res_g2 # Resumen gráfico 2.

# --------------------------- # Separador de sección.
# 7) TABLAS # Generar tablas.
# --------------------------- # Separador visual.

tab_final <- create_table_prov(dat, cols$col_prov, tabprov_anio, tabprov_se_inicio, tabprov_se_fin, tabprov_excluir_labs, tabprov_lab_solo, tabprov_unidad, carpeta) # Tabla provincia.

tabla_se_final <- create_table_se(dat, cols$col_micro, cols$col_estab, tabse_anio, tabse_se, tabse_excluir_labs, tabse_lab_solo, tabse_unidad, carpeta) # Tabla SE.

# --------------------------- # Separador de sección.
# 8) MENSAJES FINALES # Mensajes al usuario.
# --------------------------- # Separador visual.

message("OK. Carpeta del reporte: ", carpeta) # Mensaje carpeta.
message("SE del reporte (según max Fecha Verificación): SE ", sprintf("%02d", se_reporte), " - Año ", anio_rep) # Mensaje SE/año.
message("SE máxima en datos (según Fecha Colección): SE ", sprintf("%02d", se_reporte_coleccion), " - Año ", anio_coleccion_rep) # Mensaje SE colección.
message("Gráfico 1: ", out_g1) # Mensaje gráfico 1.
message("Gráfico 2: ", out_g2) # Mensaje gráfico 2.
message("Listo.") # Mensaje final.

# --------------------------- # Separador de sección.
# 9) PPT # Generar PPT ejecutivo.
# --------------------------- # Separador visual.

generate_ppt(carpeta, out_g1, out_g2, sem_plot, dat, res_g2, tab_final, tabla_se_final) # Generar PPT.
