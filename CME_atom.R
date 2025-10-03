

# ------------ 0) Pacotes e instalação condicional ------------
pkgs <- c("pacman")
if (!all(pkgs %in% rownames(installed.packages()))) install.packages(setdiff(pkgs, rownames(installed.packages())))
suppressWarnings(suppressMessages(library(pacman)))

pacman::p_load(
  tidyverse, scales, officer, rvg, janitor, lubridate
)

# ------------ 1) Parâmetros editáveis ------------
# Template
TEMPLATE_PATH <- "C:\\Users\\Sacha\\OneDrive - Ministério da Saúde\\Documentos\\CME_automatizado\\CME.pptx"


MASTER_NAME <- "CME"
LAYOUT_NAME <- "slide_1"


# Dados (ajuste caminhos conforme sua máquina)
ARQ_DENGUE_2025 <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Dengue atual\\DENGUE.csv"
ARQ_DENGUE_2024 <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Dengue\\DENGUE2024.csv"
ARQ_CHIK_2025   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Chik atual\\CHIKUNGUNYA.csv"
ARQ_CHIK_2024   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Chik\\CHIKUNGUNYA2024.csv"
POP_PATH        <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Populacao\\pop_dc_18_25_uf.csv"  # opcional

# Escopo da apresentação (defina o que gerar)
incluir_brasil   <- TRUE
regioes_alvo     <- c("Nordeste", "Norte")   # ou NA
ufs_alvo         <- c("PE","AL")    # ou NA
POP_YEAR <- 2025  


# Doença(s)
agravos          <- c("dengue","chik")

# Janela epidemiológica
se_inicial  <- 27
se_final    <- NA    # se NA, usar epiweek(Sys.Date())
data_base   <- Sys.Date()

# Saída
pasta_saida <- "C:\\Users\\Sacha\\OneDrive - Ministério da Saúde\\Documentos\\CME_automatizado\\"


ACCURACY_INC <- 0.1   
ACCURACY_PCT <- 0.1   

log_msg <- function(...){
  message(sprintf("[%s] %s", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), paste(..., collapse=" ")))
}

safe_read <- function(path){
  try_order <- list(
    list(fun = readr::read_csv2, args = list(file=path, locale=readr::locale(encoding="UTF-8", decimal_mark=",", grouping_mark="."), guess_max=200000, show_col_types=FALSE, progress=FALSE)),
    list(fun = readr::read_delim, args = list(file=path, delim=";", locale=readr::locale(encoding="UTF-8", decimal_mark=",", grouping_mark="."), guess_max=200000, show_col_types=FALSE, progress=FALSE, trim_ws=TRUE)),
    list(fun = readr::read_csv,  args = list(file=path, locale=readr::locale(encoding="UTF-8", decimal_mark=".", grouping_mark=","), guess_max=200000, show_col_types=FALSE, progress=FALSE)),
    list(fun = readr::read_delim, args = list(file=path, delim="\t", locale=readr::locale(encoding="UTF-8"), guess_max=200000, show_col_types=FALSE, progress=FALSE, trim_ws=TRUE)),
    list(fun = readr::read_delim, args = list(file=path, delim=";", locale=readr::locale(encoding="Latin1", decimal_mark=",", grouping_mark="."), guess_max=200000, show_col_types=FALSE, progress=FALSE, trim_ws=TRUE))
  )
  for (trk in try_order){
    x <- try(do.call(trk$fun, trk$args), silent = TRUE)
    if (!inherits(x, "try-error") && is.data.frame(x) && ncol(x) > 1) return(x)
  }
  stop("Não foi possível ler o arquivo (separador/encoding desconhecido): ", path)
}

# ------------ 4) Mapeamento UF -> Região ------------
uf_to_regiao <- function(uf){
  uf <- toupper(uf)
  ne <- c("BA","SE","AL","PE","PB","RN","CE","PI","MA")
  n  <- c("AC","AP","AM","PA","RO","RR","TO")
  co <- c("GO","MT","MS","DF")
  se <- c("SP","RJ","ES","MG")
  sul<- c("PR","SC","RS")
  case_when(
    uf %in% ne ~ "Nordeste",
    uf %in% n  ~ "Norte",
    uf %in% co ~ "Centro-Oeste",
    uf %in% se ~ "Sudeste",
    uf %in% sul~ "Sul",
    TRUE ~ NA_character_
  )
}

# ------------ 5) Datas/SE (SINAN) ------------
inferir_data_evento <- function(df){
  nms <- names(df)
  prefer <- c(
    "dt_sinpri","dt_sin_pri","dt_primeiros_sintomas","dt_inicio_sintomas","data_inicio_sintomas","dt_sintomas","data_sintomas",
    "dt_notif","dt_notific","data_notificacao",
    "dt_invest","dt_investigacao",
    "dt_encerra","dt_encerramento",
    "dt_obito"
  )
  hit <- intersect(prefer, nms)
  if (length(hit)) return(hit[1])
  dt_like <- grep("^dt_", nms, value = TRUE)
  if (length(dt_like)) return(dt_like[1])
  regex_cands <- nms[grepl("(data|dt).*(sin|sint|notif|invest|encer|obit)", nms)]
  if (length(regex_cands)) return(regex_cands[1])
  return(NA_character_)
}

adicionar_se_ano <- function(df, prefer_cols = c(
  "dt_sinpri","dt_sin_pri","dt_primeiros_sintomas","dt_inicio_sintomas","data_inicio_sintomas","dt_sintomas","data_sintomas",
  "dt_notif","dt_notific","data_notificacao",
  "dt_invest","dt_investigacao",
  "dt_encerra","dt_encerramento",
  "dt_obito"
)){
  cand <- intersect(prefer_cols, names(df))
  nm_data <- if (length(cand)) cand[1] else inferir_data_evento(df)
  if (is.na(nm_data) || is.null(nm_data))
    stop("Não foi possível identificar coluna de data para calcular SE.")
  x <- df[[nm_data]]
  
  to_date <- function(v){
    if (inherits(v, "Date")) return(v)
    if (inherits(v, c("POSIXct","POSIXt"))) return(as.Date(v))
    if (is.factor(v)) v <- as.character(v)
    if (is.character(v)){
      suppressWarnings({
        ymd_ <- lubridate::ymd(v, quiet = TRUE)
        dmy_ <- lubridate::dmy(v, quiet = TRUE)
        mdy_ <- lubridate::mdy(v, quiet = TRUE)
      })
      return(coalesce(ymd_, dmy_, mdy_))
    }
    if (is.numeric(v)){
      out <- rep(as.Date(NA), length(v))
      idx <- is.finite(v) & v > 10000 & v < 90000
      out[idx] <- as.Date(v[idx], origin = "1899-12-30")
      return(out)
    }
    as.Date(NA)
  }
  
  data_col <- to_date(x)
  if (all(is.na(data_col)))
    stop("Falha ao parsear datas em: ", nm_data)
  
  df %>%
    mutate(
      data_evento = as.Date(data_col),
      semana_epi  = lubridate::epiweek(data_evento),
      ano_epi     = suppressWarnings(lubridate::epiyear(data_evento)),
      ano         = year(data_evento),
      se_label    = ifelse(!is.na(semana_epi) & !is.na(ano_epi),
                           sprintf("%d-SE%02d", ano_epi, as.integer(semana_epi)),
                           NA_character_)
    )
}


map_ibge_uf <- c(
  "11"="RO","12"="AC","13"="AM","14"="RR","15"="PA","16"="AP","17"="TO",
  "21"="MA","22"="PI","23"="CE","24"="RN","25"="PB","26"="PE","27"="AL","28"="SE","29"="BA",
  "31"="MG","32"="ES","33"="RJ","35"="SP",
  "41"="PR","42"="SC","43"="RS",
  "50"="MS","51"="MT","52"="GO","53"="DF"
)

normalize_uf <- function(uf, municipio = NULL){
  u <- toupper(stringr::str_trim(as.character(uf)))
  
  is_ibge2 <- !is.na(u) & grepl("^\\d{2}$", u) & u %in% names(map_ibge_uf)
  u[is_ibge2] <- unname(map_ibge_uf[u[is_ibge2]])
  
  if (!is.null(municipio)) {
    m <- as.character(municipio)
    pick <- (is.na(u) | !nzchar(u)) & !is.na(m) & grepl("^\\d{7}$", m)
    if (any(pick)) {
      cod2 <- substr(m[pick], 1, 2)
      u[pick] <- unname(map_ibge_uf[cod2])
    }
  }
  
  u <- toupper(stringr::str_trim(u))
  siglas_validas <- c("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT",
                      "PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO")
  u[!(u %in% siglas_validas)] <- NA_character_
  u
}

# ------------ 6) Leitura/padronização das bases ------------
ler_e_padronizar <- function(path, agravo = c("dengue","chik")){
  agravo <- match.arg(agravo)
  if (!file.exists(path)) stop("Arquivo não encontrado: ", path)
  
  log_msg("Lendo: ", path)
  x <- safe_read(path) %>% janitor::clean_names()
  

  if (!"uf" %in% names(x)){
    if ("coufinf" %in% names(x)) {
      x <- x %>% mutate(uf = toupper(coufinf))
    } else {
      cand_uf <- c("sg_uf","uf","uf_not","sg_uf_not","uf_res","uf_residencia")
      cand_uf <- cand_uf[cand_uf %in% names(x)]
      x$uf <- if (length(cand_uf)) toupper(x[[cand_uf[1]]]) else NA_character_
      if (length(cand_uf)) log_msg("UF via fallback: ", cand_uf[1])
    }
  } else x$uf <- toupper(x$uf)
  
  x$uf <- stringr::str_trim(toupper(x$uf))
  x$uf[!nzchar(x$uf) | nchar(x$uf) != 2] <- NA_character_

  x$regiao <- uf_to_regiao(x$uf)
  
  if (!"municipio" %in% names(x)){
    cand_mun <- c("comuninf","co_municipio","id_municipio","municipio","municipio_residencia","id_mn_resi")
    cand_mun <- cand_mun[cand_mun %in% names(x)]
    x$municipio <- if (length(cand_mun)) x[[cand_mun[1]]] else NA
    if (length(cand_mun)) log_msg("Município via fallback: ", cand_mun[1])
  }
  
  x$uf <- normalize_uf(x$uf, municipio = x$municipio)
  x$regiao <- uf_to_regiao(x$uf)
  
  if (!"casos" %in% names(x)) x$casos <- 1L
  
  if (!"evolucao" %in% names(x)) x$evolucao <- NA_integer_
  
  if (!"sorotipo" %in% names(x)) x$sorotipo <- NA_character_
  
  x <- adicionar_se_ano(x)
  
  x$agravo <- agravo
  x
  
  
  
}

# ------------ 7) População por UF (CSV ou fallback) ------------

carregar_pop <- function(pop_path, pop_year = POP_YEAR) {
  # Fallback embutido
  fallback_vec <- function() {
    log_msg("POP_PATH não encontrado; usando vetor embutido (ajuste quando possível).")
    tibble::tibble(
      uf = c("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"),
      pop = c(906876, 3351543, 4292875, 877613, 15044172, 9240580, 3094325, 4086788, 7113540, 7200181, 21168791, 2778986,
              3526220, 8777124, 4018127, 9674793, 3273227, 11433957, 17264943, 3560903, 1777225, 652713, 11377239, 7164788,
              2318822, 45919049, 1607363)
    )
  }
  
  if (!file.exists(pop_path)) return(fallback_vec())
  
  log_msg("Carregando população UF de: ", pop_path)
  df <- safe_read(pop_path) %>% janitor::clean_names()
  
  uf_cands <- c("uf","sigla_uf","sg_uf","uf_sigla","estado_sigla","estado","nome_uf","nm_uf")
  uf_col <- intersect(uf_cands, names(df))
  if (!length(uf_col)) stop("CSV de população sem coluna de UF. Colunas: ", paste(names(df), collapse=", "))
  
  nms <- names(df)
  year_cols <- nms[grepl("^ano_\\d{4}$|^\\d{4}$", nms)]
  if (!length(year_cols)) {
    if ("pop" %in% nms) {
      return(df %>% transmute(uf = toupper(.data[[uf_col[1]]]),
                              pop = readr::parse_number(as.character(pop),
                                                        locale = readr::locale(decimal_mark=",", grouping_mark="."))) %>%
               filter(!is.na(uf) & !is.na(pop)))
    }
    stop("CSV de população não possui colunas ano_YYYY/YYYY nem 'pop'. Colunas: ", paste(nms, collapse=", "))
  }
  
  target_col <- paste0("ano_", pop_year)
  if (!(target_col %in% year_cols)) {
    if (as.character(pop_year) %in% year_cols) {
      target_col <- as.character(pop_year)
    } else {
      anos_detectados <- as.integer(gsub("^ano_", "", gsub("^(.+)$", "\\1", year_cols)))
      anos_detectados[is.na(anos_detectados)] <- as.integer(year_cols[is.na(anos_detectados)])
      ano_escolhido <- max(anos_detectados, na.rm = TRUE)
      target_col <- if (paste0("ano_", ano_escolhido) %in% year_cols) paste0("ano_", ano_escolhido) else as.character(ano_escolhido)
      log_msg("Aviso: coluna para ", pop_year, " não encontrada. Usando ano disponível: ", target_col)
    }
  }
  
  out <- df %>%
    transmute(
      uf_raw = .data[[uf_col[1]]],
      pop_raw = .data[[target_col]]
    ) %>%
    mutate(
      uf = {
        nome_estado_para_sigla <- function(v) {
          mapa <- c(
            "ACRE"="AC","ALAGOAS"="AL","AMAPA"="AP","AMAPÁ"="AP","AMAZONAS"="AM","BAHIA"="BA",
            "CEARA"="CE","CEARÁ"="CE","DISTRITO FEDERAL"="DF","ESPIRITO SANTO"="ES","ESPÍRITO SANTO"="ES",
            "GOIAS"="GO","GOIÁS"="GO","MARANHAO"="MA","MARANHÃO"="MA","MATO GROSSO"="MT","MATO GROSSO DO SUL"="MS",
            "MINAS GERAIS"="MG","PARA"="PA","PARÁ"="PA","PARAIBA"="PB","PARAÍBA"="PB","PARANA"="PR","PARANÁ"="PR",
            "PERNAMBUCO"="PE","PIAUI"="PI","PIAUÍ"="PI","RIO DE JANEIRO"="RJ","RIO GRANDE DO NORTE"="RN",
            "RIO GRANDE DO SUL"="RS","RONDONIA"="RO","RONDÔNIA"="RO","RORAIMA"="RR","SANTA CATARINA"="SC",
            "SAO PAULO"="SP","SÃO PAULO"="SP","SERGIPE"="SE","TOCANTINS"="TO"
          )
          up <- toupper(stringr::str_trim(v))
          out <- unname(mapa[up])
          ifelse(is.na(out), up, out)
        }
        v <- ifelse(nchar(uf_raw) <= 2 | is.na(uf_raw), uf_raw, nome_estado_para_sigla(uf_raw))
        toupper(v)
      },
      pop = {
        p1 <- suppressWarnings(readr::parse_number(as.character(pop_raw),
                                                   locale = readr::locale(decimal_mark=",", grouping_mark=".")))
        p2 <- suppressWarnings(readr::parse_number(as.character(pop_raw),
                                                   locale = readr::locale(decimal_mark=".", grouping_mark=",")))
        dplyr::coalesce(p1, p2)
      }
    ) %>%
    select(uf, pop) %>%
    filter(!is.na(uf) & !is.na(pop)) %>%
    distinct()
  
  ufs_validas <- c("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO")
  out <- out %>% filter(uf %in% ufs_validas)
  
  if (nrow(out) < length(ufs_validas)) {
    faltam <- setdiff(ufs_validas, out$uf)
    if (length(faltam)) log_msg("Aviso: UFs faltando população: ", paste(faltam, collapse=", "))
  }
  
  out
}


# ------------ 8) KPIs por escopo ------------
kpis_escopo <- function(dados24, dados25, se_ini, se_fim, escopo = c("brasil","regiao","uf"), nome = NULL, pop_uf){
  escopo <- match.arg(escopo)
  SE_TXT  <- sprintf("SE %d – %d", se_ini, se_fim)
  
  # Filtra por escopo
  fesc <- function(d) {
    if (escopo == "brasil") return(d)
    if (escopo == "regiao") return(d %>% filter(regiao == nome))
    if (escopo == "uf")     return(d %>% filter(uf == toupper(nome)))
  }
  
  d24 <- fesc(dados24) %>% filter(between(semana_epi, se_ini, se_fim))
  d25 <- fesc(dados25) %>% filter(between(semana_epi, se_ini, se_fim))
  
  pop_den <- case_when(
    escopo == "brasil" ~ sum(pop_uf$pop, na.rm = TRUE),
    escopo == "regiao" ~ { ufs <- pop_uf %>% mutate(regiao = uf_to_regiao(uf)) %>% filter(regiao == nome); sum(ufs$pop, na.rm=TRUE) },
    escopo == "uf"     ~ { ufs <- pop_uf %>% filter(uf == toupper(nome)); sum(ufs$pop, na.rm=TRUE) }
  )
  
  casos25 <- sum(d25$casos, na.rm = TRUE)
  casos24 <- sum(d24$casos, na.rm = TRUE)
  
  inc25   <- ifelse(pop_den > 0, (casos25 / pop_den) * 1e5, NA_real_)
  
  ob_conf <- sum(ifelse(d25$evolucao == 2, 1L, 0L), na.rm = TRUE)
  ob_inv  <- sum(ifelse(d25$evolucao == 4, 1L, 0L), na.rm = TRUE)
  
  varpct  <- ifelse(casos24 > 0, (casos25/casos24 - 1) * 100, NA_real_)
  seta    <- ifelse(is.na(varpct), "", ifelse(varpct < 0, "▼ ", "▲ "))
  pct_txt <- ifelse(is.na(varpct), "ND", paste0(seta, scales::number(abs(varpct), accuracy = ACCURACY_PCT, decimal.mark = ","), "%"))
  
  agrv <- if (length(unique(d25$agravo)) == 1) unique(d25$agravo) else unique(d24$agravo)
  agrv <- ifelse(is.na(agrv) || length(agrv)==0, "", toupper(agrv[1]))
  esc_txt <- if (escopo=="brasil") "BRASIL" else toupper(nome)
  titulo <- sprintf("SITUAÇÃO EPIDEMIOLÓGICA SE %d – %d – %s %s", se_ini, se_fim, agrv, esc_txt)
  
  list(
    casos_provaveis     = scales::number(casos25, big.mark = ".", decimal.mark = ","),
    incidencia          = scales::number(inc25, accuracy = ACCURACY_INC, decimal.mark = ",", big.mark = "."),
    obitos_investi      = scales::number(ob_inv, big.mark = ".", decimal.mark = ","),
    obitos_confirmados  = scales::number(ob_conf, big.mark = ".", decimal.mark = ","),
    compara_2024        = pct_txt,
    titulo              = titulo
  )
}

# ------------ 9) Preenchimento de placeholders ------------

required_labels <- c("SE","SE_2","data","casos_provaveis","incidencia",
                     "obitos_investi","obitos_confirmados","compara_2024")

check_layout_has_labels <- function(doc, layout, master, req_labels = required_labels){
  lp <- officer::layout_properties(doc, layout = layout, master = master)
  lbls <- lp$ph_label
  types <- lp$type
  
  missing <- setdiff(req_labels, lbls)
  has_title <- any(types == "title") 
  
  if (length(missing) || !has_title) {
    stop(
      "Placeholders não encontrados no layout escolhido.\n",
      "Layout='", layout, "' | Master='", master, "'\n",
      if (length(missing)) paste0("Faltando (labels): ", paste(missing, collapse = ", "), "\n") else "",
      if (!has_title) "Faltando: placeholder de TÍTULO (type = 'title')\n" else "",
      "Disponíveis (ph_label): ", paste(na.omit(unique(lbls)), collapse = ", "), "\n",
      "Tipos disponíveis: ", paste(unique(types), collapse = ", ")
    )
  }
  invisible(TRUE)
}

find_layout_with_labels <- function(doc, req_labels = required_labels){
  ls <- officer::layout_summary(doc)
  for (i in seq_len(nrow(ls))) {
    lay <- ls$layout[i]; mast <- ls$master[i]
    lp <- try(officer::layout_properties(doc, layout = lay, master = mast), silent = TRUE)
    if (inherits(lp, "try-error")) next
    lbls <- lp$ph_label; types <- lp$type
    if (all(req_labels %in% lbls) && any(types == "title"))
      return(list(layout = lay, master = mast))
  }
  stop("Nenhum layout contém todos os placeholders: ",
       paste(req_labels, collapse = ", "),
       " e um placeholder type='title'. Revise o Slide Master.")
}

preencher_slide <- function(doc, valores, SE_NUM, SE2_NUM, DATA_TXT, layout, master){
  check_layout_has_labels(doc, layout, master, required_labels)
  doc <- officer::add_slide(doc, layout = layout, master = master)
  
  doc <- officer::ph_with(doc, sprintf("%d", SE_NUM),  location = officer::ph_location_label("SE"))
  doc <- officer::ph_with(doc, sprintf("%d", SE2_NUM), location = officer::ph_location_label("SE_2"))
  doc <- officer::ph_with(doc, DATA_TXT,                location = officer::ph_location_label("data"))
  
  doc <- officer::ph_with(doc, valores$casos_provaveis,    location = officer::ph_location_label("casos_provaveis"))
  doc <- officer::ph_with(doc, valores$incidencia,         location = officer::ph_location_label("incidencia"))
  doc <- officer::ph_with(doc, valores$obitos_investi,     location = officer::ph_location_label("obitos_investi"))
  doc <- officer::ph_with(doc, valores$obitos_confirmados, location = officer::ph_location_label("obitos_confirmados"))
  doc <- officer::ph_with(doc, valores$compara_2024,       location = officer::ph_location_label("compara_2024"))

  lp <- officer::layout_properties(doc, layout = layout, master = master)
  if ("titulo" %in% lp$ph_label) {
    doc <- officer::ph_with(doc, valores$titulo, location = officer::ph_location_label("titulo"))
  } else {
    doc <- officer::ph_with(doc, valores$titulo, location = officer::ph_location_type(type = "title"))
  }
  doc
}



# ------------ 10) Slide por escopo ------------
montar_slide_por_escopo <- function(doc, agravo, escopo, nome, dados24, dados25, se_ini, se_fim, se_atual, data_base, layout, master, pop_uf){
  # Filtra agravo
  d24 <- dados24 %>% filter(agravo == agravo)
  d25 <- dados25 %>% filter(agravo == agravo)
  
  # Calcula KPIs
  vals <- kpis_escopo(d24, d25, se_ini, se_fim, escopo = escopo, nome = nome, pop_uf = pop_uf)
  
  # Textos SE e Data
  DATA_TXT <- paste0(" ", format(data_base, "%d.%m.%y"))
  
  # Preenche slide
  doc <- preencher_slide(
    doc, valores = vals,
    SE_NUM = se_atual, SE2_NUM = se_atual, DATA_TXT = DATA_TXT,
    layout = layout, master = master
  )
  doc
}

# ============================================================
# 11) EXECUÇÃO
# ============================================================
if (!file.exists(TEMPLATE_PATH)) stop("Template não encontrado: ", TEMPLATE_PATH)
if (!dir.exists(pasta_saida)) dir.create(pasta_saida, recursive = TRUE, showWarnings = FALSE)

dengue_2025 <- ler_e_padronizar(ARQ_DENGUE_2025, "dengue")
dengue_2024 <- ler_e_padronizar(ARQ_DENGUE_2024, "dengue")
chik_2025   <- ler_e_padronizar(ARQ_CHIK_2025,   "chik")
chik_2024   <- ler_e_padronizar(ARQ_CHIK_2024,   "chik")

pop_uf <- carregar_pop(POP_PATH)

SE_ATUAL <- if (is.na(se_final)) lubridate::epiweek(Sys.Date()) else se_final
SE_INI   <- se_inicial
SE_FIM   <- SE_ATUAL

doc_probe <- read_pptx(TEMPLATE_PATH)
if (is.null(LAYOUT_NAME) || is.null(MASTER_NAME)) {
  pick <- find_layout_with_labels(doc_probe, required_labels)
  LAYOUT_NAME <- pick$layout
  MASTER_NAME <- pick$master
  log_msg("Layout detectado: ", LAYOUT_NAME, " | Master: ", MASTER_NAME)
} else {
  check_layout_has_labels(doc_probe, LAYOUT_NAME, MASTER_NAME, required_labels)
}

doc <- read_pptx(TEMPLATE_PATH)

DATA_TXT <- paste0(" ", format(data_base, "%d.%m.%y"))

for (agr in agravos){
  log_msg("Gerando slides para agravo: ", toupper(agr))
  
  # Brasil
  if (isTRUE(incluir_brasil)) {
    doc <- montar_slide_por_escopo(
      doc, agravo = agr, escopo = "brasil", nome = "BRASIL",
      dados24 = if (agr=="dengue") dengue_2024 else chik_2024,
      dados25 = if (agr=="dengue") dengue_2025 else chik_2025,
      se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
      layout = LAYOUT_NAME, master = MASTER_NAME, pop_uf = pop_uf
    )
  }
  
  # Regiões
  if (!is.null(regioes_alvo) && !all(is.na(regioes_alvo)) && length(regioes_alvo)>0){
    for (reg in regioes_alvo){
      doc <- montar_slide_por_escopo(
        doc, agravo = agr, escopo = "regiao", nome = reg,
        dados24 = if (agr=="dengue") dengue_2024 else chik_2024,
        dados25 = if (agr=="dengue") dengue_2025 else chik_2025,
        se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
        layout = LAYOUT_NAME, master = MASTER_NAME, pop_uf = pop_uf
      )
    }
  }
  
  # UFs
  if (!is.null(ufs_alvo) && !all(is.na(ufs_alvo)) && length(ufs_alvo)>0){
    for (ufx in ufs_alvo){
      doc <- montar_slide_por_escopo(
        doc, agravo = agr, escopo = "uf", nome = toupper(ufx),
        dados24 = if (agr=="dengue") dengue_2024 else chik_2024,
        dados25 = if (agr=="dengue") dengue_2025 else chik_2025,
        se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
        layout = LAYOUT_NAME, master = MASTER_NAME, pop_uf = pop_uf
      )
    }
  }
}

# h) Salvar
fname <- file.path(
  pasta_saida,
  sprintf("CME_%s_SE_%02d_%s.pptx", toupper(paste(agravos, collapse="-")), SE_ATUAL, format(data_base, "%Y%m%d"))
)
print(fname)
print(doc, target = fname)
log_msg("Apresentação gerada: ", fname)


# ============================================================
