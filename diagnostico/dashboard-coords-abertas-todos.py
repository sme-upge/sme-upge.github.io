import pandas as pd
import numpy as np
import os
import json
import random

# --- 1. Configurações ---
base_path = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Questionário"

respostas_file_name = "respostas-diagnostico-anon.xlsx"
participacao_file_name = "emails-servidores.xlsx"
# NOVO ARQUIVO DE SUGESTÕES
sugestoes_file_name = "respostas - sugestões.xlsx"

respostas_file_path = os.path.join(base_path, respostas_file_name)
participacao_file_path = os.path.join(base_path, participacao_file_name)
# NOVO CAMINHO DO ARQUIVO DE SUGESTÕES
sugestoes_file_path = os.path.join(base_path, sugestoes_file_name)

output_dir = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Relatório"
output_html_filename = "index.html"
output_html_path = os.path.join(output_dir, output_html_filename)

os.makedirs(output_dir, exist_ok=True)

label_mapping = {
    "Diretoria Regional de Educação Butantã": "DRE-BT",
    "Diretoria Regional de Educação Campo Limpo": "DRE-CL",
    "Diretoria Regional de Educação Capela do Socorro": "DRE-CS",
    "Diretoria Regional de Educação Freguesia/Brasilândia": "DRE-FB",
    "Diretoria Regional de Educação Guaianases": "DRE-G",
    "Diretoria Regional de Educação Ipiranga": "DRE-IP",
    "Diretoria Regional de Educação Itaquera": "DRE-IQ",
    "Diretoria Regional de Educação Jaçanã/Tremembé": "DRE-JT",
    "Diretoria Regional de Educação Penha": "DRE-PE",
    "Diretoria Regional de Educação Pirituba/Jaraguá": "DRE-PJ",
    "Diretoria Regional de Educação Santo Amaro": "DRE-SA",
    "Diretoria Regional de Educação São Mateus": "DRE-SM",
    "Diretoria Regional de Educação São Miguel": "DRE-MP",
    "Secretaria Municipal de Educação (órgão central)": "SME"
}

fixed_location_color_map = {
    "SME": "#D32F2F", "DRE-BT": "#1976D2", "DRE-CL": "#388E3C",
    "DRE-CS": "#F57C00", "DRE-FB": "#7B1FA2", "DRE-G": "#00796B",
    "DRE-IP": "#E64A19", "DRE-IQ": "#5D4037", "DRE-JT": "#C2185B",
    "DRE-PE": "#0288D1", "DRE-PJ": "#AFB42B", "DRE-SA": "#FFA000",
    "DRE-SM": "#616161", "DRE-MP": "#455A64"
}

location_col_letter = 'G'
sme_unit_col_letter = 'X'
question_col_letters = [chr(i) for i in range(ord('I'), ord('V') + 1)]

location_col_index = ord(location_col_letter) - ord('A')
sme_unit_col_index = ord(sme_unit_col_letter) - ord('A')
question_col_indices = [ord(letter) - ord('A') for letter in question_col_letters]
use_cols_indices = [location_col_index, sme_unit_col_index] + question_col_indices


# --- 2. Funções Auxiliares ---
def aggregate_low_counts(item_counts_series, threshold=5, other_label="Outras respostas"):
    low_count_items = item_counts_series[item_counts_series < threshold]
    high_count_items = item_counts_series[item_counts_series >= threshold]
    aggregated_counts = high_count_items.copy()
    if not low_count_items.empty:
        other_sum = low_count_items.sum()
        if other_sum > 0:
            aggregated_counts[other_label] = aggregated_counts.get(other_label, 0) + other_sum
    return aggregated_counts.sort_values(ascending=False)

def generate_random_color():
    return f"#{random.randint(0, 0xFFFFFF):06x}"

# --- 3. Processamento de Dados de Participação ---
print("Iniciando processamento dos dados de participação...")
try:
    emails_df = pd.read_excel(participacao_file_path, sheet_name="emails-servidores")

    known_dre_codes = ["PE", "SA", "FB", "BT", "IQ", "G", "MP", "CL", "PJ", "JT", "SM", "CS", "IP"]

    def transform_dre_value(dre):
        if isinstance(dre, str):
            if dre == "GA": return "SME/GAB"
            if dre in known_dre_codes: return f"DRE-{dre}"
            return f"SME/{dre}"
        return dre

    if 'DRE' in emails_df.columns:
        emails_df['DRE'] = emails_df['DRE'].apply(transform_dre_value)

    emails_filtrados = emails_df[(emails_df['RESP'].astype(str) != "0") & (emails_df['RESP'].astype(str) != "1")]
    contagem_resp_nao_zero_por_unidade = emails_filtrados.groupby('DRE').size().reset_index(name='Respostas')
    contagem_resp_nao_zero_por_unidade = contagem_resp_nao_zero_por_unidade.rename(columns={'DRE': 'Unidade'})

    total_linhas_por_unidade = emails_df.groupby('DRE').size().reset_index(name='Total esperado')
    total_linhas_por_unidade = total_linhas_por_unidade.rename(columns={'DRE': 'Unidade'})

    analise_por_unidade = pd.merge(contagem_resp_nao_zero_por_unidade, total_linhas_por_unidade, on='Unidade', how='left')

    analise_por_unidade['Porcentagem_Num'] = 0.0
    analise_por_unidade.loc[analise_por_unidade['Total esperado'].notna() & (analise_por_unidade['Total esperado'] > 0), 'Porcentagem_Num'] = \
        (analise_por_unidade['Respostas'] / analise_por_unidade['Total esperado']) * 100

    analise_por_unidade_sorted = analise_por_unidade.sort_values(by='Porcentagem_Num', ascending=False).reset_index(drop=True)
    analise_por_unidade_sorted['Participação'] = analise_por_unidade_sorted['Porcentagem_Num'].map('{:.2f}%'.format)
    analise_por_unidade_sorted = analise_por_unidade_sorted.drop(columns=['Porcentagem_Num'])

    total_geral_quantidade_resp_nao_zero = emails_filtrados.shape[0]
    total_geral_linhas = emails_df.shape[0]
    overall_percentage = (total_geral_quantidade_resp_nao_zero / total_geral_linhas) * 100 if total_geral_linhas > 0 else 0

    total_row_dict = {
        'Unidade': 'TOTAL GERAL',
        'Respostas': total_geral_quantidade_resp_nao_zero,
        'Total esperado': total_geral_linhas,
        'Participação': f"{overall_percentage:.2f}%"
    }
    
    participation_table_data = analise_por_unidade_sorted.to_dict('records')
    
    print("Dados de participação processados com sucesso.")

except FileNotFoundError:
    print(f"Erro: O arquivo de participação '{participacao_file_path}' não foi encontrado. A tabela não será gerada.")
    participation_table_data = []
    total_row_dict = {}
except Exception as e:
    print(f"Erro crítico ao processar o arquivo de participação: {e}")
    participation_table_data = []
    total_row_dict = {}


# --- 4. Ler e Preparar os Dados para os Gráficos ---
print("Iniciando leitura e processamento dos dados do questionário para os gráficos...")
try:
    header_df = pd.read_excel(respostas_file_path, usecols=use_cols_indices, nrows=1, header=None)
    
    sme_unit_col_name = 'Unidade SME'
    column_titles = {
        location_col_index: 'Local de trabalho',
        sme_unit_col_index: sme_unit_col_name
    }
    question_titles_map = {}
    for i, col_index in enumerate(question_col_indices):
        title = header_df.iloc[0, 1 + i]
        column_titles[col_index] = title
        question_titles_map[question_col_letters[i]] = title

    df = pd.read_excel(respostas_file_path, usecols=use_cols_indices, header=None, skiprows=1)
    df.rename(columns=column_titles, inplace=True)
    
    print("Arquivo de respostas lido com sucesso.")

    df.dropna(subset=['Local de trabalho'], inplace=True)
    df = df[df['Local de trabalho'].astype(str).str.strip() != '']
    df['Local de trabalho'] = df['Local de trabalho'].map(label_mapping).fillna(df['Local de trabalho'])

    df[sme_unit_col_name] = df[sme_unit_col_name].astype(str).str.strip().replace('nan', '')
    
    sme_sub_units_df = df[df[sme_unit_col_name].str.startswith('SME/', na=False)].copy()
    
    if df.empty and sme_sub_units_df.empty:
        raise ValueError("Nenhum dado válido de 'Local de trabalho' ou 'Unidade SME' encontrado.")

    dre_sme_counts = df['Local de trabalho'].value_counts()
    sme_sub_counts = sme_sub_units_df[sme_unit_col_name].value_counts()
    all_location_counts = pd.concat([dre_sme_counts, sme_sub_counts])

    locations_with_enough_responses = all_location_counts[all_location_counts >= 10].index.tolist()
    
    # *** INÍCIO DA MODIFICAÇÃO ***
    # Obter listas de locais elegíveis (sem ordenação inicial)
    dre_sme_locations = [loc for loc in df['Local de trabalho'].unique() if loc in locations_with_enough_responses]
    sme_sub_locations = [loc for loc in sme_sub_units_df[sme_unit_col_name].unique() if loc in locations_with_enough_responses]
    
    # Combinar todas as localizações e remover duplicatas
    other_locations = list(set(dre_sme_locations + sme_sub_locations))

    # Ordenar alfabeticamente, mas colocar 'SME' no topo se existir
    if 'SME' in other_locations:
        other_locations.remove('SME')
        other_locations = ['SME'] + sorted(other_locations)
    else:
        other_locations = sorted(other_locations)

    # Adicionar "Rede" como a primeira opção na lista final para o dropdown
    all_comparable_locations = ['Rede'] + other_locations
    # *** FIM DA MODIFICAÇÃO ***


except Exception as e:
    print(f"Erro crítico ao processar o arquivo de respostas: {e}")
    # Definir valores padrão para evitar falha total do script
    all_comparable_locations = []
    question_titles_map = {}


# --- 5. LER E PREPARAR DADOS DE SUGESTÕES (NOVO BLOCO) ---
print("Iniciando leitura e processamento dos dados de sugestões...")
try:
    sugestoes_df = pd.read_excel(sugestoes_file_path, sheet_name="Principal", usecols="B,E,D", header=0)
    sugestoes_df.columns = ['local_original', 'eixo', 'sugestao']

    # Remover linhas onde a sugestão está vazia
    sugestoes_df.dropna(subset=['sugestao'], inplace=True)
    sugestoes_df = sugestoes_df[sugestoes_df['sugestao'].astype(str).str.strip() != '']

    # Aplicar o mesmo mapeamento de nomes de local
    sugestoes_df['local'] = sugestoes_df['local_original'].map(label_mapping).fillna(sugestoes_df['local_original'])
    
    # Limpar e preparar os dados para JSON
    sugestoes_df.dropna(subset=['local', 'eixo'], inplace=True)
    sugestoes_df['local'] = sugestoes_df['local'].astype(str).str.strip()
    sugestoes_df['eixo'] = sugestoes_df['eixo'].astype(str).str.strip()

    # Criar listas para os filtros dropdown
    suggestion_locations = sorted(sugestoes_df['local'].unique().tolist())
    suggestion_axes = sorted(sugestoes_df['eixo'].unique().tolist())
    
    # Converter para lista de dicionários
    suggestions_data = sugestoes_df[['local', 'eixo', 'sugestao']].to_dict('records')
    print("Dados de sugestões processados com sucesso.")

except FileNotFoundError:
    print(f"Aviso: O arquivo de sugestões '{sugestoes_file_path}' não foi encontrado. A tabela de sugestões não será gerada.")
    suggestions_data = []
    suggestion_locations = []
    suggestion_axes = []
except Exception as e:
    print(f"Erro crítico ao processar o arquivo de sugestões: {e}")
    suggestions_data = []
    suggestion_locations = []
    suggestion_axes = []


# --- 6. Estruturar os Dados para o Dashboard ---
print("Estruturando os dados para o dashboard HTML...")

all_colors = fixed_location_color_map.copy()
# É necessário separar os locais para a busca de dados, pois 'Rede' não existe como um local no dataframe
locations_for_data_processing = [loc for loc in all_comparable_locations if loc != 'Rede']

for unit in sme_sub_locations:
    if unit not in all_colors:
        all_colors[unit] = generate_random_color()

dashboard_data = {
    "locationColors": all_colors,
    "responsesByLocation": {},
    "allComparableLocations": all_comparable_locations,
    "questions": {},
    "answers": {},
    "participationData": participation_table_data,
    "participationTotals": total_row_dict,
    "suggestionsData": suggestions_data,
    "suggestionLocations": suggestion_locations,
    "suggestionAxes": suggestion_axes
}

if 'df' in locals():
    location_counts = df['Local de trabalho'].value_counts()
    dashboard_data["responsesByLocation"] = location_counts.to_dict()

    for col_letter, question_title in question_titles_map.items():
        print(f"Processando Eixo: {question_title}")
        
        dashboard_data["questions"][col_letter] = question_title
        dashboard_data["answers"][col_letter] = {}

        items_series = df[question_title].astype(str).str.strip().replace('', np.nan).dropna()
        if items_series.empty: continue

        total_respondents_for_question = len(items_series)
        
        items_exploded = (
            items_series.str.split(';')
            .explode()
            .str.strip()
            .replace('', np.nan)
            .dropna()
            .replace("Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família", "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias")
        )
        
        if items_exploded.empty: continue
            
        global_counts = aggregate_low_counts(items_exploded.value_counts(), threshold=5, other_label="Outras respostas")
        dashboard_data["answers"][col_letter]["Global"] = {
            "totalRespondents": total_respondents_for_question,
            "items": global_counts.head(15).to_dict()
        }

        # Iterar sobre os locais de dados, que não incluem 'Rede'
        for location in locations_for_data_processing:
            if location in dre_sme_locations:
                df_location = df[df['Local de trabalho'] == location]
            else: # SME Sub-units
                df_location = df[df[sme_unit_col_name] == location]

            local_items_series = df_location[question_title].astype(str).str.strip().replace('', np.nan).dropna()
            if local_items_series.empty: continue

            total_local_respondents = len(local_items_series)
            local_items_exploded = (
                local_items_series.str.split(';')
                .explode()
                .str.strip()
                .replace('', np.nan)
                .dropna()
                .replace("Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família", "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias")
            )

            if local_items_exploded.empty: continue

            local_counts = aggregate_low_counts(local_items_exploded.value_counts(), threshold=5, other_label="Outras respostas")
            
            show_warning = False
            other_responses_count = local_counts.get("Outras respostas", 0)
            if other_responses_count > total_local_respondents:
                show_warning = True
            
            dashboard_data["answers"][col_letter][location] = {
                "totalRespondents": total_local_respondents,
                "items": local_counts.head(15).to_dict(),
                "showOtherResponsesWarning": show_warning
            }


# --- 7. Gerar o Arquivo HTML ---
print("Gerando o arquivo HTML do dashboard...")

data_as_json_string = json.dumps(dashboard_data, indent=4, ensure_ascii=False, default=lambda x: int(x) if isinstance(x, (np.int64, np.int32)) else x)


html_template = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Diagnóstico da Educação Paulistana</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f0f2f5;
        }}
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
        .chart-legend-item {{
            display: flex;
            align-items: center;
            font-size: 0.875rem;
        }}
        .chart-legend-color-box {{
            width: 1rem;
            height: 1rem;
            margin-right: 0.5rem;
            border-radius: 0.25rem;
        }}
        .sortable-th {{
            cursor: pointer;
            position: sticky;
            top: 0;
            background-color: #f8fafc; /* Tailwind's gray-50 */
            z-index: 10;
        }}
        .sortable-th:hover {{
            background-color: #e2e8f0; /* Tailwind's gray-200 */
        }}
        .sortable-th .sort-asc::after {{
            content: ' ▲';
            font-size: 0.8em;
        }}
        .sortable-th .sort-desc::after {{
            content: ' ▼';
            font-size: 0.8em;
        }}
    </style>
</head>
<body class="text-gray-800">

    <div class="container mx-auto p-4 md:p-8">
        <header class="mb-8 flex flex-col md:flex-row justify-between items-center">
            <div>
                <h1 class="text-3xl md:text-4xl font-bold text-gray-800">Diagnóstico da Educação Paulistana</h1>
                <p class="text-md text-gray-600 mt-2">Resultados preliminares por DRE e coordenadoria da SME do questionário de diagnóstico</p>
            </div>
            <img src="logo-pe.png" alt="Logotipo do Planejamento Estratégico 2025-2028" class="mt-4 md:mt-0" style="max-width: 200px; height: auto;">
        </header>

        <section class="bg-white p-6 rounded-2xl shadow-lg mb-8">
            <h2 class="text-2xl font-semibold mb-4">Análise Comparativa por Eixo</h2>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                <div>
                    <label for="question-select" class="block text-sm font-medium text-gray-700 mb-1">Selecione o Eixo (Pergunta):</label>
                    <select id="question-select" class="w-full p-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"></select>
                </div>
                <div>
                    <label for="location-select" class="block text-sm font-medium text-gray-700 mb-1">Selecione um Local para análise:</label>
                    <select id="location-select" class="w-full p-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"></select>
                </div>
            </div>

            <div class="grid grid-cols-1 mt-8">
                <div class="bg-gray-50 p-4 rounded-xl">
                    <h3 id="local-chart-title" class="text-lg font-semibold text-center mb-1"></h3>
                    <p id="local-chart-subtitle" class="text-sm text-center text-red-600 italic font-bold mb-2"></p>
                    <div style="height: 975px;" class="w-full">
                        <canvas id="localComparisonChart"></canvas>
                    </div>
                </div>
            </div>
            
            <div id="comparison-legend" class="mt-6 flex justify-center items-center space-x-6 text-xs text-gray-600">
                <div class="chart-legend-item">
                    <div class="chart-legend-color-box" style="background-color: #d32f2f;"></div>
                    <span>Item 20% ou mais ACIMA da média da Rede</span>
                </div>
                <div class="chart-legend-item">
                    <div class="chart-legend-color-box" style="background-color: #388e3c;"></div>
                    <span>Item 20% ou mais ABAIXO da média da Rede</span>
                </div>
            </div>
        </section>
        
        <section id="suggestions-section" class="bg-white p-6 rounded-2xl shadow-lg mb-8">
            <h2 class="text-2xl font-semibold mb-4">Sugestões abertas</h2>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                <div>
                    <label for="suggestion-location-select" class="block text-sm font-medium text-gray-700 mb-1">Filtrar por Local:</label>
                    <select id="suggestion-location-select" class="w-full p-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"></select>
                </div>
                <div>
                    <label for="suggestion-eixo-select" class="block text-sm font-medium text-gray-700 mb-1">Filtrar por Eixo:</label>
                    <select id="suggestion-eixo-select" class="w-full p-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"></select>
                </div>
            </div>
            
            <div class="overflow-y-auto" style="max-height: 500px;">
                <table id="suggestionsTable" class="w-full text-sm text-left text-gray-500">
                    <thead class="text-xs text-gray-700 uppercase">
                        <tr>
                            <th scope="col" class="py-3 px-6 sortable-th" data-column-key="local" style="width: 10%;">Local <span class="sort-icon"></span></th>
                            <th scope="col" class="py-3 px-6 sortable-th" data-column-key="eixo" style="width: 30%;">Eixo <span class="sort-icon"></span></th>
                            <th scope="col" class="py-3 px-6 sortable-th" data-column-key="sugestao" style="width: 60%;">Sugestão <span class="sort-icon"></span></th>
                        </tr>
                    </thead>
                    <tbody id="suggestions-table-body">
                    </tbody>
                </table>
            </div>
        </section>
        
        <section class="bg-white p-6 rounded-2xl shadow-lg mb-8">
            <h2 class="text-2xl font-semibold mb-4">Respostas por Unidade</h2>
            <div style="height: 450px;" class="w-full">
                <canvas id="locationResponsesChart"></canvas>
            </div>
        </section>
        
        <section id="participation-section" class="bg-white p-6 rounded-2xl shadow-lg">
            <h2 class="text-2xl font-semibold mb-4">Engajamento por Unidade</h2>
            <div class="overflow-y-auto" style="max-height: 600px;">
                <table id="participationTable" class="w-full text-sm text-left text-gray-500">
                    <thead class="text-xs text-gray-700 uppercase">
                        <tr>
                            <th scope="col" class="py-3 px-6 sortable-th" data-column-index="0" data-column-type="string">Unidade <span class="sort-icon"></span></th>
                            <th scope="col" class="py-3 px-6 text-right sortable-th" data-column-index="1" data-column-type="number">Respostas <span class="sort-icon"></span></th>
                            <th scope="col" class="py-3 px-6 text-right sortable-th" data-column-index="2" data-column-type="number">Total esperado <span class="sort-icon"></span></th>
                            <th scope="col" class="py-3 px-6 text-right sortable-th" data-column-index="3" data-column-type="percentage">Participação <span class="sort-icon"></span></th>
                        </tr>
                    </thead>
                    <tbody>
                    </tbody>
                    <tfoot class="font-bold text-gray-800 bg-gray-100 sticky bottom-0">
                    </tfoot>
                </table>
            </div>
             <p class="text-xs text-gray-500 mt-4">
                A determinação da unidade foi realizada pelo cruzamento do e-mail usado para responder o questionário com os dados cadastrais do servidor no EOL em 15/05/2025. Aproximadamente 15% das pessoas não tiveram sua unidade determinada.
            </p>
        </section>
        
        <footer class="text-center mt-12 text-gray-500 text-sm">
            <p>Secretaria Municipal de Educação / Planejamento Estratégico 2025-2028</p>
        </footer>
    </div>

<script>
const surveyData = {data_as_json_string};

Chart.register(ChartDataLabels);

function wrapText(text, maxWidth) {{
    if (typeof text !== 'string') return '';
    if (text.length <= maxWidth) return text;
    
    const words = text.split(' ');
    const lines = [];
    let currentLine = '';

    for (const word of words) {{
        if ((currentLine + ' ' + word).length > maxWidth && currentLine.length > 0) {{
            lines.push(currentLine.trim());
            currentLine = '';
        }}
        currentLine += word + ' ';
    }}
    if (currentLine) lines.push(currentLine.trim());
    return lines;
}}

document.addEventListener('DOMContentLoaded', () => {{
    let locationChart, localChart;

    const questionSelect = document.getElementById('question-select');
    const locationSelect = document.getElementById('location-select');
    const participationTable = document.getElementById('participationTable');
    const suggestionsSection = document.getElementById('suggestions-section');
    const suggestionLocationSelect = document.getElementById('suggestion-location-select');
    const suggestionEixoSelect = document.getElementById('suggestion-eixo-select');
    const suggestionsTable = document.getElementById('suggestionsTable');
    const suggestionsTableBody = document.getElementById('suggestions-table-body');
    
    let participationSort = {{ column: 3, order: 'desc' }};
    let suggestionSort = {{ key: 'local', order: 'asc' }};

    function renderParticipationTable() {{
        const tableBody = participationTable.querySelector('tbody');
        const tableFoot = participationTable.querySelector('tfoot');
        tableBody.innerHTML = '';
        tableFoot.innerHTML = '';

        if (!surveyData.participationData || surveyData.participationData.length === 0) {{
            document.getElementById('participation-section').style.display = 'none';
            return;
        }}
        
        const dataToSort = [...surveyData.participationData];
        dataToSort.sort((a, b) => {{
            const key = Object.keys(a)[participationSort.column];
            const valA = a[key];
            const valB = b[key];
            const columnType = participationTable.querySelector(`th[data-column-index='${{participationSort.column}}']`).dataset.columnType;
            let compareA, compareB;

            if (columnType === 'percentage') {{
                compareA = parseFloat(valA.replace('%', '').replace(',', '.'));
                compareB = parseFloat(valB.replace('%', '').replace(',', '.'));
            }} else if (columnType === 'number') {{
                compareA = Number(valA);
                compareB = Number(valB);
            }} else {{
                compareA = String(valA).toLowerCase();
                compareB = String(valB).toLowerCase();
            }}

            if (compareA < compareB) return participationSort.order === 'asc' ? -1 : 1;
            if (compareA > compareB) return participationSort.order === 'asc' ? 1 : -1;
            return 0;
        }});

        dataToSort.forEach(row => {{
            const tr = `
                <tr class="bg-white border-b hover:bg-gray-50">
                    <td class="py-4 px-6 font-medium text-gray-900">${{row['Unidade']}}</td>
                    <td class="py-4 px-6 text-right">${{Number(row['Respostas']).toLocaleString('pt-BR')}}</td>
                    <td class="py-4 px-6 text-right">${{Number(row['Total esperado']).toLocaleString('pt-BR')}}</td>
                    <td class="py-4 px-6 text-right">${{row['Participação']}}</td>
                </tr>
            `;
            tableBody.innerHTML += tr;
        }});

        const totals = surveyData.participationTotals;
        if (totals && totals['Unidade']) {{
            const totalRow = `
                <tr class="border-t-2 border-gray-300">
                    <td class="py-4 px-6">${{totals['Unidade']}}</td>
                    <td class="py-4 px-6 text-right">${{Number(totals['Respostas']).toLocaleString('pt-BR')}}</td>
                    <td class="py-4 px-6 text-right">${{Number(totals['Total esperado']).toLocaleString('pt-BR')}}</td>
                    <td class="py-4 px-6 text-right">${{totals['Participação']}}</td>
                </tr>
            `;
            tableFoot.innerHTML = totalRow;
        }}
        updateSortIcons(participationTable, participationSort.column, participationSort.order, 'data-column-index');
    }}
    
    function populateSuggestionFilters() {{
        suggestionLocationSelect.innerHTML = '<option value="Todos">Todos</option>';
        surveyData.suggestionLocations.forEach(location => {{
            const option = document.createElement('option');
            option.value = location;
            option.textContent = location;
            suggestionLocationSelect.appendChild(option);
        }});
        
        suggestionEixoSelect.innerHTML = '<option value="Todos">Todos</option>';
        surveyData.suggestionAxes.forEach(eixo => {{
            const option = document.createElement('option');
            option.value = eixo;
            option.textContent = eixo;
            suggestionEixoSelect.appendChild(option);
        }});
    }}
    
    function renderSuggestionsTable() {{
        const selectedLocation = suggestionLocationSelect.value;
        const selectedEixo = suggestionEixoSelect.value;
        
        const filteredData = surveyData.suggestionsData.filter(item => {{
            const locationMatch = (selectedLocation === 'Todos' || item.local === selectedLocation);
            const eixoMatch = (selectedEixo === 'Todos' || item.eixo === selectedEixo);
            return locationMatch && eixoMatch;
        }});
        
        filteredData.sort((a, b) => {{
            const key = suggestionSort.key;
            const valA = String(a[key] || '').toLowerCase();
            const valB = String(b[key] || '').toLowerCase();
            
            if (valA < valB) return suggestionSort.order === 'asc' ? -1 : 1;
            if (valA > valB) return suggestionSort.order === 'asc' ? 1 : -1;
            return 0;
        }});
        
        suggestionsTableBody.innerHTML = ''; 
        
        if (filteredData.length === 0) {{
            suggestionsTableBody.innerHTML = `<tr><td colspan="3" class="text-center py-4 px-6 text-gray-500">Nenhuma sugestão encontrada para os filtros selecionados.</td></tr>`;
        }} else {{
            filteredData.forEach(item => {{
                const row = document.createElement('tr');
                row.className = 'bg-white border-b hover:bg-gray-50';
                row.innerHTML = `
                    <td class="py-4 px-6 align-top">${{item.local}}</td>
                    <td class="py-4 px-6 align-top">${{item.eixo}}</td>
                    <td class="py-4 px-6 align-top">${{item.sugestao}}</td>
                `;
                suggestionsTableBody.appendChild(row);
            }});
        }}
        updateSortIcons(suggestionsTable, suggestionSort.key, suggestionSort.order, 'data-column-key');
    }}

    function renderLocationResponsesChart() {{
        const ctx = document.getElementById('locationResponsesChart').getContext('2d');
        const data = surveyData.responsesByLocation;
        const labels = Object.keys(data).sort((a, b) => data[b] - data[a]);
        const values = labels.map(label => data[label]);
        const backgroundColors = labels.map(label => surveyData.locationColors[label] || '#cccccc');

        if (locationChart) locationChart.destroy();
        locationChart = new Chart(ctx, {{
            type: 'bar',
            data: {{ labels, datasets: [{{ label: 'Respostas', data: values, backgroundColor: backgroundColors }}] }},
            options: {{
                indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    datalabels: {{ display: true, anchor: 'end', align: 'end', color: '#333', font: {{ weight: 'bold', size: 12 }}, formatter: (v) => v.toLocaleString('pt-BR') }}
                }},
                scales: {{
                    x: {{ type: 'logarithmic', title: {{ display: true, text: 'Quantidade de Respostas (Escala Log)' }} }},
                    y: {{ ticks: {{ autoSkip: false }} }}
                }}
            }}
        }});
    }}
    
    // *** INÍCIO DA MODIFICAÇÃO DO GRÁFICO PRINCIPAL ***
    function renderComparisonChart(questionId, locationId) {{
        const localCtx = document.getElementById('localComparisonChart').getContext('2d');
        const globalQuestionData = surveyData.answers[questionId]?.['Global'];
        const subtitleElement = document.getElementById('local-chart-subtitle');
        const titleElement = document.getElementById('local-chart-title');
        const comparisonLegend = document.getElementById('comparison-legend');
        subtitleElement.textContent = '';

        if (localChart) localChart.destroy();

        // CASO 1: Usuário selecionou "Rede" para ver apenas os dados gerais
        if (locationId === 'Rede') {{
            comparisonLegend.style.display = 'none'; // Esconde a legenda de comparação
            if (globalQuestionData) {{
                const {{ totalRespondents: totalGlobal, items: globalItems }} = globalQuestionData;
                titleElement.textContent = `Resultados Gerais da Rede (${{totalGlobal.toLocaleString('pt-BR')}} resp.)`;

                const sortedGlobalItems = Object.entries(globalItems).sort(([, a], [, b]) => b - a);
                const labels = sortedGlobalItems.map(([label]) => label);
                const globalAbsoluteCounts = sortedGlobalItems.map(([, value]) => value);
                const globalPercentages = globalAbsoluteCounts.map(count => totalGlobal > 0 ? (count / totalGlobal) * 100 : 0);

                const maxPercentage = Math.max(0, ...globalPercentages);
                const suggestedMax = Math.ceil(maxPercentage / 10) * 10 + 5;

                localChart = new Chart(localCtx, {{
                    type: 'bar',
                    data: {{
                        labels,
                        datasets: [{{
                            label: 'Rede',
                            data: globalPercentages,
                            backgroundColor: '#a0aec0', // Cor neutra para a rede
                            absoluteCounts: globalAbsoluteCounts
                        }}]
                    }},
                    options: getChartOptions('y', '% de Respostas', (ctx) => `${{ctx.raw.toFixed(1)}}%`, {{
                        formatter: (value, context) => {{
                            const count = context.chart.data.datasets[0].absoluteCounts[context.dataIndex];
                            return `${{count}} (${{value.toFixed(1)}}%)`;
                        }}
                    }}, '#666', suggestedMax) // Cor padrão para os ticks do eixo Y
                }});
            }} else {{
                titleElement.textContent = 'Resultados Gerais da Rede';
                localCtx.clearRect(0, 0, localCtx.canvas.width, localCtx.canvas.height);
                localCtx.font = "16px Inter";
                localCtx.fillStyle = "#888";
                localCtx.textAlign = "center";
                localCtx.fillText(`Sem dados para este eixo.`, localCtx.canvas.width / 2, localCtx.canvas.height / 2);
            }}
        // CASO 2: Usuário selecionou um local específico para comparar com a rede (lógica original)
        }} else {{
            comparisonLegend.style.display = 'flex'; // Mostra a legenda de comparação
            const localQuestionData = surveyData.answers[questionId]?.[locationId];

            if (localQuestionData && globalQuestionData) {{
                const {{ totalRespondents: totalLocal, items: localItems, showOtherResponsesWarning }} = localQuestionData;
                const {{ totalRespondents: totalGlobal, items: globalItems }} = globalQuestionData;

                titleElement.textContent = `Comparativo: ${{locationId}} (${{totalLocal.toLocaleString('pt-BR')}} resp.) vs. Rede (${{totalGlobal.toLocaleString('pt-BR')}} resp.)`;
                if (showOtherResponsesWarning) {{
                    subtitleElement.textContent = `Como cada respondente podia selecionar até 3 itens por eixo e escrever sua própria sugestão, a porcentagem de "Outras respostas" pode ultrapassar 100%`;
                }}

                const sortedLocalItems = Object.entries(localItems).sort(([, a], [, b]) => b - a);
                const labels = sortedLocalItems.map(([label]) => label);
                const localAbsoluteCounts = sortedLocalItems.map(([, value]) => value);
                const tickColors = [], localPercentages = [], globalPercentagesForCompare = [];

                labels.forEach((label, i) => {{
                    const local_perc = totalLocal > 0 ? (localAbsoluteCounts[i] / totalLocal) * 100 : 0;
                    const global_perc = totalGlobal > 0 ? ((globalItems[label] || 0) / totalGlobal) * 100 : 0;
                    localPercentages.push(local_perc);
                    globalPercentagesForCompare.push(global_perc);

                    if (global_perc > 0) {{
                        const diff = (local_perc - global_perc) / global_perc;
                        if (diff >= 0.20) tickColors.push('#d32f2f');
                        else if (diff <= -0.20) tickColors.push('#388e3c');
                        else tickColors.push('#666');
                    }} else {{
                        tickColors.push('#666');
                    }}
                }});
                
                const maxPercentage = Math.max(0, ...localPercentages, ...globalPercentagesForCompare);
                const suggestedMax = Math.ceil(maxPercentage / 10) * 10 + 5;

                localChart = new Chart(localCtx, {{
                    type: 'bar',
                    data: {{ labels, datasets: [
                        {{ label: locationId, data: localPercentages, backgroundColor: surveyData.locationColors[locationId] || '#333', absoluteCounts: localAbsoluteCounts }},
                        {{ label: 'Rede', data: globalPercentagesForCompare, backgroundColor: '#a0aec0' }}
                    ]}},
                    options: getChartOptions('y',`% de Respostas (${{locationId}} vs. Rede)`,
                        (ctx) => `${{ctx.dataset.label}}: ${{ctx.raw.toFixed(1)}}%`,
                        {{ formatter: (value, context) => {{
                            const dataset = context.chart.data.datasets[context.datasetIndex];
                            if(dataset.label !== 'Rede') {{ const count = dataset.absoluteCounts[context.dataIndex]; return `${{count}} (${{value.toFixed(1)}}%)`; }}
                            return `${{value.toFixed(1)}}%`;
                        }}}}, tickColors, suggestedMax)
                }});
            }} else {{
                titleElement.textContent = `Comparativo Local vs. Rede`;
                localCtx.clearRect(0, 0, localCtx.canvas.width, localCtx.canvas.height);
                localCtx.font = "16px Inter";
                localCtx.fillStyle = "#888";
                localCtx.textAlign = "center";
                localCtx.fillText(`Sem dados para '${{locationId}}' neste eixo.`, localCtx.canvas.width / 2, localCtx.canvas.height / 2);
            }}
        }}
    }}
    // *** FIM DA MODIFICAÇÃO ***
    
    function getChartOptions(axis, xAxisTitle, tooltipLabelCallback, datalabelsConfig, yAxisTickColors = '#666', suggestedMax = 100) {{
        return {{
            indexAxis: axis, responsive: true, maintainAspectRatio: false, layout: {{ padding: {{ left: 0 }} }},
            plugins: {{
                legend: {{ position: 'top' }},
                tooltip: {{ callbacks: {{ label: tooltipLabelCallback }} }},
                datalabels: {{ display: true, anchor: 'end', align: 'end', color: '#333', font: {{ weight: 'bold', size: 12 }}, ...datalabelsConfig }}
            }},
            scales: {{
                x: {{ suggestedMax, title: {{ display: true, text: xAxisTitle }}, ticks: {{ callback: value => value.toFixed(0) + "%" }} }},
                y: {{ ticks: {{ color: yAxisTickColors, callback: function(value) {{ return wrapText(this.getLabelForValue(value), 70); }} }} }}
            }}
        }};
    }}

    // *** INÍCIO DA MODIFICAÇÃO DOS FILTROS ***
    function populateComparisonSelects() {{
        Object.entries(surveyData.questions).forEach(([id, title]) => {{
            const option = document.createElement('option');
            option.value = id; option.textContent = title;
            questionSelect.appendChild(option);
        }});
        // A lista de locais agora vem pré-ordenada do Python
        const locations = surveyData.allComparableLocations;
        locations.forEach(location => {{
            const option = document.createElement('option');
            option.value = location; option.textContent = location;
            locationSelect.appendChild(option);
        }});
    }}
    // *** FIM DA MODIFICAÇÃO ***
    
    function updateSortIcons(table, activeColumn, order, keyAttribute) {{
        table.querySelectorAll('th .sort-icon').forEach(icon => icon.textContent = '');
        const activeTh = table.querySelector(`th[${{keyAttribute}}='${{activeColumn}}'] .sort-icon`);
        if (activeTh) activeTh.textContent = order === 'asc' ? ' ▲' : ' ▼';
    }}

    function init() {{
        populateComparisonSelects();
        questionSelect.addEventListener('change', () => renderComparisonChart(questionSelect.value, locationSelect.value));
        locationSelect.addEventListener('change', () => renderComparisonChart(questionSelect.value, locationSelect.value));
        
        renderLocationResponsesChart();
        // A chamada inicial agora renderizará o gráfico da "Rede" por padrão
        renderComparisonChart(questionSelect.value, locationSelect.value);

        if (surveyData.participationData && surveyData.participationData.length > 0) {{
            renderParticipationTable();
            participationTable.querySelectorAll('th.sortable-th').forEach(th => {{
                th.addEventListener('click', () => {{
                    const columnIndex = parseInt(th.dataset.columnIndex, 10);
                    if (participationSort.column === columnIndex) {{
                        participationSort.order = participationSort.order === 'asc' ? 'desc' : 'asc';
                    }} else {{
                        participationSort.column = columnIndex;
                        participationSort.order = 'desc';
                    }}
                    renderParticipationTable();
                }});
            }});
        }} else {{
             document.getElementById('participation-section').style.display = 'none';
        }}
        
        if (surveyData.suggestionsData && surveyData.suggestionsData.length > 0) {{
            populateSuggestionFilters();
            renderSuggestionsTable();
            suggestionLocationSelect.addEventListener('change', renderSuggestionsTable);
            suggestionEixoSelect.addEventListener('change', renderSuggestionsTable);
            
            suggestionsTable.querySelectorAll('th.sortable-th').forEach(th => {{
                th.addEventListener('click', () => {{
                    const key = th.dataset.columnKey;
                    if (suggestionSort.key === key) {{
                        suggestionSort.order = suggestionSort.order === 'asc' ? 'desc' : 'asc';
                    }} else {{
                        suggestionSort.key = key;
                        suggestionSort.order = 'asc';
                    }}
                    renderSuggestionsTable();
                }});
            }});
        }} else {{
            suggestionsSection.style.display = 'none';
        }}
    }}

    init();
}});
</script>

</body>
</html>
"""

try:
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_template)
    print(f"\nDashboard gerado com sucesso!")
    print(f"Arquivo salvo em: {output_html_path}")
except Exception as e:
    print(f"Erro ao salvar o arquivo HTML: {e}")