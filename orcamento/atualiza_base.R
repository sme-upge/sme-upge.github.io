# --- Script de Atualização de Orçamento SME ---
library(readxl)
library(dplyr)
library(readr)
library(lubridate)

# 1. GERAR LINK DINÂMICO
get_url <- function(data) {
  ano <- format(data, "%Y")
  mes_ano <- format(data, "%m%y")
  return(paste0("https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/", 
                ano, "/basedadosexecucaoconsolidados_", mes_ano, ".xlsx"))
}

url_base <- get_url(Sys.Date())
arquivo_temp <- tempfile(fileext = ".xlsx")

try_download <- try(download.file(url_base, destfile = arquivo_temp, mode = "wb"), silent = TRUE)

if(inherits(try_download, "try-error")) {
  message("Mês atual não disponível, tentando mês anterior...")
  url_base <- get_url(Sys.Date() %m-% months(1))
  download.file(url_base, destfile = arquivo_temp, mode = "wb")
}

# 2. PROCESSAMENTO
base_completa <- read_excel(arquivo_temp)

mapa_unidade <- c(
  "Gabinete do Secretário" = "Gabinete do Secretário",
  "Diretoria Regional de Educação Ipiranga" = "Ipiranga",
  "Diretoria Regional de Educação - Ipiranga" = "Ipiranga",
  "Diretoria Regional de Educação Jaçanã/Tremembé" = "Jaçanã/Tremembé",
  "Diretoria Regional de Educação - Jaçanã/Tremembé" = "Jaçanã/Tremembé",
  "Diretoria Regional de Educação Freguesia/Brasilândia" = "Freguesia/Brasilândia",
  "Diretoria Regional de Educação - Freguesia/Brasilândia" = "Freguesia/Brasilândia",
  "Diretoria Regional de Educação Pirituba" = "Pirituba/Jaraguá",
  "Diretoria Regional de Educação - Pirituba" = "Pirituba/Jaraguá",
  "Diretoria Regional de Educação Campo Limpo" = "Campo Limpo",
  "Diretoria Regional de Educação - Campo Limpo" = "Campo Limpo",
  "Diretoria Regional de Educação Capela do Socorro" = "Capela do Socorro",
  "Diretoria Regional de Educação  Capela do Socorro" = "Capela do Socorro",
  "Diretoria Regional de Educação - Capela do Socorro" = "Capela do Socorro",
  "Diretoria Regional de Educação Penha" = "Penha",
  "Diretoria Regional de Educação - Penha" = "Penha",
