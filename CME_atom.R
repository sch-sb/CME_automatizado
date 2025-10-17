
pkgs <- c("pacman")
if (!all(pkgs %in% rownames(installed.packages()))) install.packages(setdiff(pkgs, rownames(installed.packages())))
suppressWarnings(suppressMessages(library(pacman)))
pacman::p_load(tidyverse, scales, officer, rvg, janitor, lubridate)

# ============================
# 1) Parâmetros editáveis
# ============================
TEMPLATE_PATH <- "C:\\Users\\Sacha\\OneDrive - Ministério da Saúde\\Documentos\\CME_automatizado\\CME.pptx"
MASTER_NAME   <- "CME"
LAYOUT_NAME_DENGUE <- "slide_1"  # dengue
LAYOUT_NAME_CHIK   <- "slide_2"  # chikungunya
LAYOUT_NAME_ZIKA   <- "slide_3"  # zika

ARQ_DENGUE_2025 <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Dengue atual\\DENGUE.csv"
ARQ_DENGUE_2024 <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Dengue\\DENGUE2024.csv"
ARQ_CHIK_2025   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Chik atual\\CHIKUNGUNYA.csv"
ARQ_CHIK_2024   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Chik\\CHIKUNGUNYA2024.csv"
ARQ_ZIKA_2025   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Zika atual\\ZIKA.csv"
ARQ_ZIKA_2024   <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Zika\\ZIKA2024.csv"

POP_PATH        <- "C:\\Relatorios DC e Nowcasting\\DC e Nowcasting_CSVs\\Populacao\\pop_dc_18_25_uf.csv"

incluir_brasil <- TRUE
regioes_alvo   <- c("Nordeste", "Sudeste")
ufs_alvo       <- c("AL", "PE", "RN", "ES")
POP_YEAR       <- 2025

agravos        <- c("dengue","chik","zika")

se_inicial  <- 27
se_final    <- 41
data_base   <- Sys.Date()

pasta_saida  <- "C:\\Users\\Sacha\\OneDrive - Ministério da Saúde\\Documentos\\CME_automatizado\\"
ACCURACY_INC <- 0.1
ACCURACY_PCT <- 0.1

# Intervalo de datas "razoável" para filtrar (pode ajustar aqui)
DATE_MIN_YEAR <- 2010
DATE_MAX_YEAR <- as.integer(format(Sys.Date(), "%Y")) + 1

# ============================
# 1A) Mapeamento de classi_fin
# ============================
DENGUE_CODES <- c(10, 11, 12)  # Mantido apenas como referência (não é mais determinante)
CHIK_CODES   <- 13             # 13 = Chikungunya
ZIKA_CODES   <- c(1, 2)        # 1/2 = Zika (conforme base)
DESC_CODE    <- 5              # 5 = Descartado

is_prob_dengue <- function(cf){
  cf <- suppressWarnings(as.integer(cf))
  is.na(cf) || !(cf %in% c(DESC_CODE, CHIK_CODES))
}

log_msg <- function(...) message(sprintf("[%s] %s", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), paste(..., collapse=" ")))

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
  stop("Não foi possível ler o arquivo: ", path)
}

uf_to_regiao <- function(uf){
  uf <- toupper(uf)
  ne <- c("BA","SE","AL","PE","PB","RN","CE","PI","MA")
  n  <- c("AC","AP","AM","PA","RO","RR","TO")
  co <- c("GO","MT","MS","DF")
  se <- c("SP","RJ","ES","MG")
  sul<- c("PR","SC","RS")
  dplyr::case_when(
    uf %in% ne ~ "Nordeste",
    uf %in% n  ~ "Norte",
    uf %in% co ~ "Centro-Oeste",
    uf %in% se ~ "Sudeste",
    uf %in% sul~ "Sul",
    TRUE ~ NA_character_
  )
}

inferir_data_evento <- function(df){
  prefer <- c(
    "dt_sinpri","dt_sin_pri","dt_primeiros_sintomas","dt_inicio_sintomas","data_inicio_sintomas","dt_sintomas","data_sintomas",
    "dt_notif","dt_notific","data_notificacao",
    "dt_invest","dt_investigacao",
    "dt_encerra","dt_encerramento",
    "dt_obito"
  )
  nms <- names(df)
  hit <- intersect(prefer, nms)
  if (length(hit)) return(hit[1])
  dt_like <- grep("^dt_", nms, value = TRUE)
  if (length(dt_like)) return(dt_like[1])
  regex_cands <- nms[grepl("(data|dt).*(sin|sint|notif|invest|encer|obit)", nms)]
  if (length(regex_cands)) return(regex_cands[1])
  NA_character_
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
  if (is.na(nm_data)) stop("Não foi possível identificar coluna de data para calcular SE.")
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
  if (all(is.na(data_col))) stop("Falha ao parsear datas em: ", nm_data)
  df %>%
    mutate(
      data_evento = as.Date(data_col),
      semana_epi  = lubridate::epiweek(data_evento),
      ano_epi     = suppressWarnings(lubridate::epiyear(data_evento)),
      ano         = year(data_evento),
      se_label    = ifelse(!is.na(semana_epi) & !is.na(ano_epi),
                           sprintf("%d-SE%02d", ano_epi, as.integer(semana_epi)), NA_character_)
    )
}

na_if_blank <- function(v){
  v <- as.character(v)
  v <- stringr::str_trim(v)
  v[v %in% c("", "NA", "NaN", "null", "NULL")] <- NA_character_
  v
}

safe_coalesce_vec <- function(df, cols) {
  cols <- intersect(cols, names(df))
  if (length(cols) == 0) return(NULL)
  vecs <- lapply(cols, function(nm) na_if_blank(df[[nm]]))
  vecs <- lapply(vecs, as.character)
  dplyr::coalesce(!!!vecs)
}

mun_to_7 <- function(m){
  if (is.null(m)) return(NA_character_)
  m <- na_if_blank(m)
  m_num <- suppressWarnings(readr::parse_number(m))
  m_str <- ifelse(is.na(m_num), NA_character_, as.character(as.integer(m_num)))
  idx_lt7 <- !is.na(m_str) & nchar(m_str) < 7
  if (any(idx_lt7)) m_str[idx_lt7] <- stringr::str_pad(m_str[idx_lt7], 7, pad = "0")
  idx_gt7 <- !is.na(m_str) & nchar(m_str) > 7
  if (any(idx_gt7)) m_str[idx_gt7] <- substr(m_str[idx_gt7], 1, 7)
  bad <- !is.na(m_str) & !grepl("^\\d{7}$", m_str)
  if (any(bad)) m_str[bad] <- NA_character_
  m_str
}

map_ibge_uf <- c(
  "11"="RO","12"="AC","13"="AM","14"="RR","15"="PA","16"="AP","17"="TO",
  "21"="MA","22"="PI","23"="CE","24"="RN","25"="PB","26"="PE","27"="AL","28"="SE","29"="BA",
  "31"="MG","32"="ES","33"="RJ","35"="SP",
  "41"="PR","42"="SC","43"="RS","50"="MS","51"="MT","52"="GO","53"="DF"
)

normalize_uf <- function(uf, municipio=NULL){
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
  siglas_validas <- c("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT",
                      "PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO")
  u[!(u %in% siglas_validas)] <- NA_character_
  u
}

dedup_mais_recente <- function(df){
  id_cols <- intersect(c("nu_notific","nu_notif","id_sinan","id","id_aih"), names(df))
  if (!length(id_cols)) return(df)
  df %>%
    arrange(desc(data_evento)) %>%
    distinct(across(all_of(id_cols)), .keep_all = TRUE)
}

# ============================
# 2A) Detecção de agravo por linha (alinhada à regra de provável dengue)
# ============================
detectar_agravo <- function(df) {
  n <- nrow(df)
  out <- rep(NA_character_, n)
  cf  <- suppressWarnings(as.integer(df$classi_fin))
  

  if ("agravo" %in% names(df)) {
    idx_dengue_base <- df$agravo == "dengue"
    out[idx_dengue_base & is_prob_dengue(cf)] <- "dengue"
    out[idx_dengue_base & !is_prob_dengue(cf)] <- "outro"
  }
  

  idx_chik_code <- !is.na(cf) & cf %in% CHIK_CODES
  out[idx_chik_code] <- "chik"
  

  idx_zika_code <- !is.na(cf) & cf %in% ZIKA_CODES
  out[idx_zika_code] <- "zika"
  

  if (!is.null(df$agravo)) {
    out[is.na(out) & df$agravo == "zika"] <- "zika"
    out[is.na(out) & df$agravo == "chik" & idx_chik_code] <- "chik"
  }
  
  out[is.na(out)] <- "outro"
  out
}

# ============================
# 3) Leitura/padronização das bases
# ============================
ler_e_padronizar <- function(path, agravo = c("dengue","chik","zika")){
  agravo <- match.arg(agravo)
  if (!file.exists(path)) stop("Arquivo não encontrado: ", path)
  log_msg("Lendo: ", path)
  x <- safe_read(path) %>% janitor::clean_names()
  
  cand_mun <- c("municipio","id_mn_resi","co_municipio","id_municipio","id_mn_inf","co_mun_inf","id_municip")
  mun_ibge <- safe_coalesce_vec(x, cand_mun)
  mun_ibge <- mun_to_7(mun_ibge)
  
  cand_uf <- c(
    "sg_uf","sg_uf_not",
    "uf_res","sg_uf_res","uf_residencia",
    "uf_inf","sg_uf_inf","coufinf",
    "uf_ocor","sg_uf_ocor",
    "uf"
  )
  cand_uf <- unique(c(cand_uf, grep("(^|_)uf($|_\\b)", names(x), value = TRUE, ignore.case = TRUE)))
  uf_raw <- safe_coalesce_vec(x, cand_uf)
  if (is.null(uf_raw)) uf_raw <- NA_character_
  
  x$uf     <- normalize_uf(uf_raw, municipio = mun_ibge)
  x$regiao <- uf_to_regiao(x$uf)
  
  if (!"casos"    %in% names(x)) x$casos    <- 1L
  if (!"evolucao" %in% names(x)) x$evolucao <- NA_integer_
  if (!"sorotipo" %in% names(x)) x$sorotipo <- NA_character_
  
  x <- adicionar_se_ano(x)
  
  # agravo da origem + agravo detectado
  x$agravo     <- agravo
  x$agravo_det <- detectar_agravo(x)
  
  x
}
filtrar_datas_raz <- function(df, ano_min = DATE_MIN_YEAR, ano_max = DATE_MAX_YEAR){
  df %>% dplyr::filter(is.na(data_evento) | dplyr::between(lubridate::year(data_evento), ano_min, ano_max))
}

log_range_datas <- function(df, label=""){
  rng <- range(df$data_evento, na.rm = TRUE)
  cat(sprintf(
    "\n%s | range data_evento: %s -> %s | n=%s\n",
    label,
    as.character(rng[1]), as.character(rng[2]),
    scales::number(nrow(df), big.mark = ".", decimal.mark = ",")
  ))
}

# ============================
# 4) População por UF
# ============================
carregar_pop <- function(pop_path, pop_year = POP_YEAR) {
  fallback_vec <- function() {
    log_msg("POP_PATH não encontrado; usando vetor embutido.")
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
  if (!length(uf_col)) stop("CSV de população sem coluna de UF.")
  
  nms <- names(df)
  year_cols <- nms[grepl("^ano_\\d{4}$|^\\d{4}$", nms)]
  if (!length(year_cols) && !"pop" %in% nms) stop("CSV população sem ano_YYYY/YYYY nem 'pop'.")
  
  target_col <- paste0("ano_", pop_year)
  if (!(target_col %in% year_cols)) {
    if (as.character(pop_year) %in% year_cols) {
      target_col <- as.character(pop_year)
    } else {
      anos_detectados <- suppressWarnings(as.integer(gsub("^ano_", "", year_cols)))
      anos_detectados[is.na(anos_detectados)] <- suppressWarnings(as.integer(year_cols[is.na(anos_detectados)]))
      ano_escolhido <- max(anos_detectados, na.rm = TRUE)
      target_col <- if (paste0("ano_", ano_escolhido) %in% year_cols) paste0("ano_", ano_escolhido) else as.character(ano_escolhido)
      log_msg("Aviso: usando população de ", target_col)
    }
  }
  
  out <- if (length(year_cols)) {
    df %>%
      transmute(uf_raw = .data[[uf_col[1]]], pop_raw = .data[[target_col]]) %>%
      mutate(
        uf = {
          mapa <- c(
            "ACRE"="AC","ALAGOAS"="AL","AMAPA"="AP","AMAPÁ"="AP","AMAZONAS"="AM","BAHIA"="BA",
            "CEARA"="CE","CEARÁ"="CE","DISTRITO FEDERAL"="DF","ESPIRITO SANTO"="ES","ESPÍRITO SANTO"="ES",
            "GOIAS"="GO","GOIÁS"="GO","MARANHAO"="MA","MARANHÃO"="MA","MATO GROSSO"="MT","MATO GROSSO DO SUL"="MS",
            "MINAS GERAIS"="MG","PARA"="PA","PARÁ"="PA","PARAIBA"="PB","PARAÍBA"="PB","PARANA"="PR","PARANÁ"="PR",
            "PERNAMBUCO"="PE","PIAUI"="PI","PIAUÍ"="PI","RIO DE JANEIRO"="RJ","RIO GRANDE DO NORTE"="RN",
            "RIO GRANDE DO SUL"="RS","RONDONIA"="RO","RONDÔNIA"="RO","RORAIMA"="RR","SANTA CATARINA"="SC",
            "SAO PAULO"="SP","SÃO PAULO"="SP","SERGIPE"="SE","TOCANTINS"="TO"
          )
          up <- toupper(stringr::str_trim(uf_raw))
          out <- unname(mapa[up]); ifelse(is.na(out), up, out)
        },
        pop = coalesce(
          suppressWarnings(readr::parse_number(as.character(pop_raw), locale = readr::locale(decimal_mark=",", grouping_mark="."))),
          suppressWarnings(readr::parse_number(as.character(pop_raw), locale = readr::locale(decimal_mark=".", grouping_mark=",")))
        )
      ) %>% select(uf, pop)
  } else {
    df %>% transmute(uf = toupper(.data[[uf_col[1]]]),
                     pop = readr::parse_number(as.character(pop), locale = readr::locale(decimal_mark=",", grouping_mark=".")))
  }
  
  ufs_validas <- c("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO")
  out <- out %>% filter(uf %in% ufs_validas, !is.na(pop)) %>% distinct()
  if (nrow(out) < length(ufs_validas)) log_msg("Aviso: UFs faltando população: ", paste(setdiff(ufs_validas, out$uf), collapse=", "))
  out
}

# ============================
# 5) KPIs
# ============================
kpis_escopo <- function(dados24, dados25, se_ini, se_fim,
                        escopo = c("brasil","regiao","uf"),
                        nome = NULL, pop_uf){
  
  escopo <- match.arg(escopo)
  
  filtra_escopo_e_se <- function(d){
    d <- switch(escopo,
                "brasil" = d,
                "regiao" = d %>% dplyr::filter(regiao == nome),
                "uf"     = d %>% dplyr::filter(uf == toupper(nome)))
    d %>% dplyr::filter(dplyr::between(semana_epi, se_ini, se_fim))
  }
  

  d24 <- filtra_escopo_e_se(dados24)
  d25 <- filtra_escopo_e_se(dados25)
  

  if ("classi_fin" %in% names(d24)) d24 <- d24 %>% dplyr::filter(is.na(classi_fin) | classi_fin != DESC_CODE)
  if ("classi_fin" %in% names(d25)) d25 <- d25 %>% dplyr::filter(is.na(classi_fin) | classi_fin != DESC_CODE)
  

  alvo24 <- unique(dados24$agravo)[1]
  alvo25 <- unique(dados25$agravo)[1]
  
  if ("classi_fin" %in% names(d24)) {
    d24 <- switch(alvo24,
                  "dengue" = d24 %>% dplyr::filter(is.na(classi_fin) | !(classi_fin %in% CHIK_CODES)),
                  "chik"   = d24 %>% dplyr::filter(is.na(classi_fin) | classi_fin %in% CHIK_CODES),
                  "zika"   = if (length(ZIKA_CODES) > 0) d24 %>% dplyr::filter(is.na(classi_fin) | classi_fin %in% ZIKA_CODES) else d24,
                  d24
    )
  }
  if ("classi_fin" %in% names(d25)) {
    d25 <- switch(alvo25,
                  "dengue" = d25 %>% dplyr::filter(is.na(classi_fin) | !(classi_fin %in% CHIK_CODES)),
                  "chik"   = d25 %>% dplyr::filter(is.na(classi_fin) | classi_fin %in% CHIK_CODES),
                  "zika"   = if (length(ZIKA_CODES) > 0) d25 %>% dplyr::filter(is.na(classi_fin) | classi_fin %in% ZIKA_CODES) else d25,
                  d25
    )
  }
  
  if ("agravo_det" %in% names(d24)) {
    d24_sub <- d24 %>% dplyr::filter(agravo_det == alvo24)
    if (nrow(d24_sub) > 0) d24 <- d24_sub
  }
  if ("agravo_det" %in% names(d25)) {
    d25_sub <- d25 %>% dplyr::filter(agravo_det == alvo25)
    if (nrow(d25_sub) > 0) d25 <- d25_sub
  }
  
  d24 <- dedup_mais_recente(d24)
  d25 <- dedup_mais_recente(d25)
  
  pop_den <- dplyr::case_when(
    escopo == "brasil" ~ sum(pop_uf$pop, na.rm = TRUE),
    escopo == "regiao" ~ { ufs <- pop_uf %>% dplyr::mutate(regiao = uf_to_regiao(uf)) %>% dplyr::filter(regiao == nome); sum(ufs$pop, na.rm=TRUE) },
    escopo == "uf"     ~ { ufs <- pop_uf %>% dplyr::filter(uf == toupper(nome)); sum(ufs$pop, na.rm=TRUE) }
  )
  
  casos25 <- nrow(d25)
  casos24 <- nrow(d24)
  
  inc25   <- ifelse(pop_den > 0, (casos25 / pop_den) * 1e5, NA_real_)
  ob_conf <- sum(ifelse(d25$evolucao == 2, 1L, 0L), na.rm = TRUE)
  ob_inv  <- sum(ifelse(d25$evolucao == 4, 1L, 0L), na.rm = TRUE)
  
  varpct  <- ifelse(casos24 > 0, (casos25/casos24 - 1) * 100, NA_real_)
  seta    <- ifelse(is.na(varpct), "", ifelse(varpct < 0, "▼ ", "▲ "))
  pct_txt <- ifelse(is.na(varpct), "ND", paste0(seta, scales::number(abs(varpct), accuracy = ACCURACY_PCT, decimal.mark = ","), "%"))
  
  agrv <- if (length(unique(d25$agravo)) == 1) unique(d25$agravo) else unique(d24$agravo)
  agrv <- ifelse(is.na(agrv) || length(agrv)==0, "", toupper(agrv[1]))
  esc_txt <- if (escopo=="brasil") "BRASIL" else toupper(nome)
  
  list(
    casos_provaveis     = scales::number(casos25, big.mark = ".", decimal.mark = ","),
    incidencia          = scales::number(inc25, accuracy = ACCURACY_INC, decimal.mark = ",", big.mark = "."),
    obitos_investi      = scales::number(ob_inv, big.mark = ".", decimal.mark = ","),
    obitos_confirmados  = scales::number(ob_conf, big.mark = ".", decimal.mark = ","),
    compara_2024        = pct_txt,
    titulo              = sprintf("SITUAÇÃO EPIDEMIOLÓGICA SE %d – %d – %s %s", se_ini, se_fim, agrv, esc_txt)
  )
}

# ============================
# 6) Helpers PPT
# ============================
required_labels <- c("SE","SE_2","data","casos_provaveis","incidencia","obitos_investi","obitos_confirmados","compara_2024")

check_layout_has_labels <- function(doc, layout, master, req_labels = required_labels){
  lp <- officer::layout_properties(doc, layout = layout, master = master)
  lbls <- lp$ph_label; types <- lp$type
  missing <- setdiff(req_labels, lbls); has_title <- any(types == "title")
  if (length(missing) || !has_title) {
    stop(
      "Placeholders não encontrados no layout.\n",
      "Layout='", layout, "' | Master='", master, "'\n",
      if (length(missing)) paste0("Faltando: ", paste(missing, collapse = ", "), "\n") else "",
      if (!has_title) "Faltando placeholder de TÍTULO (type='title')\n" else "",
      "Disponíveis: ", paste(na.omit(unique(lbls)), collapse = ", "), "\n",
      "Tipos: ", paste(unique(types), collapse = ", ")
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
    if (all(req_labels %in% lp$ph_label) && any(lp$type == "title"))
      return(list(layout = lay, master = mast))
  }
  stop("Nenhum layout contém todos os placeholders requeridos + 'title'.")
}

preencher_slide <- function(doc, valores, SE_NUM, SE2_NUM, DATA_TXT, layout, master){
  check_layout_has_labels(doc, layout, master, required_labels)
  doc <- officer::add_slide(doc, layout = layout, master = master)
  doc <- officer::ph_with(doc, sprintf("%d", SE_NUM),  location = officer::ph_location_label("SE"))
  doc <- officer::ph_with(doc, sprintf("%d", SE2_NUM), location = officer::ph_location_label("SE_2"))
  doc <- officer::ph_with(doc, DATA_TXT,               location = officer::ph_location_label("data"))
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

montar_slide_por_escopo <- function(doc, agravo, escopo, nome, dados24, dados25, se_ini, se_fim, se_atual, data_base, layout, master, pop_uf){
  d24  <- dados24 %>% dplyr::filter(agravo == agravo)
  d25  <- dados25 %>% dplyr::filter(agravo == agravo)
  vals <- kpis_escopo(d24, d25, se_ini, se_fim, escopo = escopo, nome = nome, pop_uf = pop_uf)
  DATA_TXT <- paste0(" ", format(data_base, "%d.%m.%y"))
  preencher_slide(doc, valores = vals, SE_NUM = se_atual, SE2_NUM = se_atual, DATA_TXT = DATA_TXT, layout = layout, master = master)
}

# ============================
# 7) Execução
# ============================
if (!file.exists(TEMPLATE_PATH)) stop("Template não encontrado: ", TEMPLATE_PATH)
if (!dir.exists(pasta_saida)) dir.create(pasta_saida, recursive = TRUE, showWarnings = FALSE)

dengue_2025 <- ler_e_padronizar(ARQ_DENGUE_2025, "dengue")
dengue_2024 <- ler_e_padronizar(ARQ_DENGUE_2024, "dengue")
chik_2025   <- ler_e_padronizar(ARQ_CHIK_2025,   "chik")
chik_2024   <- ler_e_padronizar(ARQ_CHIK_2024,   "chik")
zika_2025   <- ler_e_padronizar(ARQ_ZIKA_2025,   "zika")
zika_2024   <- ler_e_padronizar(ARQ_ZIKA_2024,   "zika")

cat("\n# Checagem de outliers de data (antes do filtro)\n")
log_range_datas(dengue_2024 %>% dplyr::filter(agravo=="dengue"), "DENGUE 2024")
log_range_datas(dengue_2025 %>% dplyr::filter(agravo=="dengue"), "DENGUE 2025")
log_range_datas(chik_2024   %>% dplyr::filter(agravo=="chik"),   "CHIK 2024")
log_range_datas(chik_2025   %>% dplyr::filter(agravo=="chik"),   "CHIK 2025")
log_range_datas(zika_2024   %>% dplyr::filter(agravo=="zika"),   "ZIKA 2024")
log_range_datas(zika_2025   %>% dplyr::filter(agravo=="zika"),   "ZIKA 2025")

dengue_2024 <- filtrar_datas_raz(dengue_2024)
dengue_2025 <- filtrar_datas_raz(dengue_2025)
chik_2024   <- filtrar_datas_raz(chik_2024)
chik_2025   <- filtrar_datas_raz(chik_2025)
zika_2024   <- filtrar_datas_raz(zika_2024)
zika_2025   <- filtrar_datas_raz(zika_2025)

cat("\n# Checagem de outliers de data (depois do filtro)\n")
log_range_datas(dengue_2024 %>% dplyr::filter(agravo=="dengue"), "DENGUE 2024")
log_range_datas(dengue_2025 %>% dplyr::filter(agravo=="dengue"), "DENGUE 2025")
log_range_datas(chik_2024   %>% dplyr::filter(agravo=="chik"),   "CHIK 2024")
log_range_datas(chik_2025   %>% dplyr::filter(agravo=="chik"),   "CHIK 2025")
log_range_datas(zika_2024   %>% dplyr::filter(agravo=="zika"),   "ZIKA 2024")
log_range_datas(zika_2025   %>% dplyr::filter(agravo=="zika"),   "ZIKA 2025")

pop_uf <- carregar_pop(POP_PATH)

REF_DATE <- Sys.Date() - 7
SE_ATUAL <- if (is.na(se_final)) lubridate::epiweek(REF_DATE) else se_final
SE_INI   <- se_inicial
SE_FIM   <- SE_ATUAL

doc_probe <- read_pptx(TEMPLATE_PATH)
check_layout_has_labels(doc_probe, LAYOUT_NAME_DENGUE, MASTER_NAME, required_labels)
check_layout_has_labels(doc_probe, LAYOUT_NAME_CHIK,   MASTER_NAME, required_labels)
check_layout_has_labels(doc_probe, LAYOUT_NAME_ZIKA,   MASTER_NAME, required_labels)

doc <- read_pptx(TEMPLATE_PATH)

layout_for <- function(agr) {
  a <- tolower(agr)
  if (a == "chik") return(LAYOUT_NAME_CHIK)
  if (a == "zika") return(LAYOUT_NAME_ZIKA)
  LAYOUT_NAME_DENGUE
}

get_data <- function(agr, ano = 2025){
  a <- tolower(agr)
  if (ano == 2025) {
    if (a == "dengue") return(dengue_2025)
    if (a == "chik")   return(chik_2025)
    if (a == "zika")   return(zika_2025)
  } else {
    if (a == "dengue") return(dengue_2024)
    if (a == "chik")   return(chik_2024)
    if (a == "zika")   return(zika_2024)
  }
  stop("Agravo/ano não reconhecido em get_data().")
}

for (agr in agravos){
  log_msg("Gerando slides para agravo: ", toupper(agr))
  lay <- layout_for(agr)
  dados24 <- get_data(agr, 2024)
  dados25 <- get_data(agr, 2025)
  
  if (isTRUE(incluir_brasil)) {
    doc <- montar_slide_por_escopo(
      doc, agravo = agr, escopo = "brasil", nome = "BRASIL",
      dados24 = dados24, dados25 = dados25,
      se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
      layout = lay, master = MASTER_NAME, pop_uf = pop_uf
    )
  }
  
  if (!is.null(regioes_alvo) && !all(is.na(regioes_alvo)) && length(regioes_alvo)>0){
    for (reg in regioes_alvo){
      doc <- montar_slide_por_escopo(
        doc, agravo = agr, escopo = "regiao", nome = reg,
        dados24 = dados24, dados25 = dados25,
        se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
        layout = lay, master = MASTER_NAME, pop_uf = pop_uf
      )
    }
  }
  
  if (!is.null(ufs_alvo) && !all(is.na(ufs_alvo)) && length(ufs_alvo)>0){
    for (ufx in ufs_alvo){
      doc <- montar_slide_por_escopo(
        doc, agravo = agr, escopo = "uf", nome = toupper(ufx),
        dados24 = dados24, dados25 = dados25,
        se_ini = SE_INI, se_fim = SE_FIM, se_atual = SE_ATUAL, data_base = data_base,
        layout = lay, master = MASTER_NAME, pop_uf = pop_uf
      )
    }
  }
}

fname <- file.path(
  pasta_saida,
  sprintf("CME_%s_SE_%02d_%s.pptx", toupper(paste(agravos, collapse="-")), SE_ATUAL, format(data_base, "%Y%m%d"))
)
print(fname)
print(doc, target = fname)
log_msg("Apresentação gerada: ", fname)
