# --- Script de Atualização de Orçamento SME ---
library(readxl)
library(dplyr)
library(readr)
library(lubridate)
options(scipen = 999) # Garante que números grandes não virem '1e+10'

# 1. LINK DA SEPLAN (O link atualizado que você encontrou)
# Este link é o oficial para o exercício de 2026
url_base <- "https://prefeitura.sp.gov.br/cidade/secretarias/upload/seplan/arquivos/Exercicio_2026/basedadosexecucaoconsolidados_2026.xlsx"

arquivo_temp <- tempfile(fileext = ".xlsx")

message("Iniciando download do arquivo da SEPLAN...")

# Tenta baixar o arquivo
try_download <- try(download.file(url_base, destfile = arquivo_temp, mode = "wb"), silent = TRUE)

# Se o link da SEPLAN falhar por algum motivo, ele tenta o link antigo da Fazenda como backup
if(inherits(try_download, "try-error")) {
  message("Link da SEPLAN falhou, tentando link reserva da Fazenda...")
  url_base <- paste0("https://orcamento.sf.prefeitura.sp.gov.br/orcamento/uploads/", 
                     format(Sys.Date(), "%Y"), "/basedadosexecucaoconsolidados_", 
                     format(Sys.Date() %m-% months(1), "%m%y"), ".xlsx")
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

# Filtro e Limpeza
base_filtrada <- base_completa %>%
  mutate(
    Cd_AnoExecucao = as.numeric(Cd_AnoExecucao),
    Ds_Orgao = trimws(Ds_Orgao),
    # Converte datas para o formato limpo (remove horas 23:59:59)
    DataInicial = as.Date(DataInicial),
    DataFinal = as.Date(DataFinal),
    DataExtracao = Sys.time()
  ) %>%
  filter(Cd_AnoExecucao >= 2010, Sigla_Orgao == "SME")

# Aplicar o mapa de unidades
base_filtrada$Ds_Unidade <- recode(base_filtrada$Ds_Unidade, !!!mapa_unidade)

# 3. SALVAR RESULTADO
# Criar pasta orcamento se não existir
if(!dir.exists("orcamento")) { dir.create("orcamento") }

# Gravar o CSV (padrão brasileiro: ; e ,)
write_excel_csv2(base_filtrada, "orcamento/Execucao_Orcamentaria_Atualizada.csv")

message("Arquivo CSV gerado com sucesso!")
