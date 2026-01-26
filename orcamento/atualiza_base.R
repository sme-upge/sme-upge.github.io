# Instalar pacotes se não houver (o GitHub fará isso)
if(!require(readxl)) install.packages("readxl")
if(!require(dplyr)) install.packages("dplyr")
if(!require(openxlsx)) install.packages("openxlsx")
if(!require(lubridate)) install.packages("lubridate")

library(readxl)
library(dplyr)
library(openxlsx)
library(lubridate)

# 1. GERAR LINK DINÂMICO (Resolve o problema de mudar todo mês)
get_url <- function(data) {
  ano <- format(data, "%Y")
  mes_ano <- format(data, "%m%y")
  return(paste0("https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/", 
                ano, "/basedadosexecucaoconsolidados_", mes_ano, ".xlsx"))
}

# Tenta o mês atual, se falhar tenta o anterior
url_base <- get_url(Sys.Date())
arquivo_temp <- tempfile(fileext = ".xlsx")

try_download <- try(download.file(url_base, destfile = arquivo_temp, mode = "wb"), silent = TRUE)

if(inherits(try_download, "try-error")) {
  url_base <- get_url(Sys.Date() %m-% months(1))
  download.file(url_base, destfile = arquivo_temp, mode = "wb")
}

# 2. PROCESSAMENTO (Seu código original com ajustes)
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
  "Diretoria Regional de Educação Santo Amaro" = "Santo Amaro",
  "Diretoria Regional de Educação - Santo Amaro" = "Santo Amaro",
  "Diretoria Regional de Educação Itaquera" = "Itaquera",
  "Diretoria Regional de Educação - Itaquera" = "Itaquera",
  "Diretoria Regional de Educação São Miguel" = "São Miguel",
  "Diretoria Regional de Educação - São Miguel" = "São Miguel",
  "Diretoria Regional de Educação Guaianases" = "Guaianases",
  "Diretoria Regional de Educação - Guaianases" = "Guaianases",
  "Diretoria Regional de Educação Butantã" = "Butantã",
  "Diretoria Regional de Educação - Butantã" = "Butantã",
  "Diretoria Regional de Educação São Mateus" = "São Mateus",
  "Diretoria Regional de Educação - São Mateus" = "São Mateus",
  "Coordenadoria de Alimentação Escolar" = "Coordenadoria de Alimentação Escolar",
  "Departamento da Merenda Escolar" = "Coordenadoria de Alimentação Escolar",
  "Departamento de Alimentação Escolar" = "Coordenadoria de Alimentação Escolar"
)

base_filtrada <- base_completa %>%
  mutate(
    Cd_AnoExecucao = as.numeric(Cd_AnoExecucao),
    Ds_Orgao = trimws(Ds_Orgao)
  ) %>%
  filter(Cd_AnoExecucao >= 2010, Sigla_Orgao == "SME")

base_filtrada$Ds_Unidade <- recode(base_filtrada$Ds_Unidade, !!!mapa_unidade)

# 3. SALVAR O RESULTADO
# No GitHub, salvamos no diretório atual
write.xlsx(base_filtrada, "orcamento/Execucao_Orcamentaria_Atualizada.xlsx", overwrite = TRUE)
