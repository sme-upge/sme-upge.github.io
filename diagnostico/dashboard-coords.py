import pandas as pd
import numpy as np
import os
import json
import random

# --- 1. Configurações ---
base_path = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Questionário"

respostas_file_name = "respostas-diagnostico-anon.xlsx"
participacao_file_name = "emails-servidores.xlsx"
respostas_file_path = os.path.join(base_path, respostas_file_name)
participacao_file_path = os.path.join(base_path, participacao_file_name)

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
    
    dre_sme_locations = sorted([loc for loc in df['Local de trabalho'].unique() if loc in locations_with_enough_responses])
    sme_sub_locations = sorted([loc for loc in sme_sub_units_df[sme_unit_col_name].unique() if loc in locations_with_enough_responses])
    all_comparable_locations = dre_sme_locations + sme_sub_locations

except Exception as e:
    print(f"Erro crítico ao processar o arquivo de respostas: {e}")
    dashboard_data = {
        "locationColors": {},
        "responsesByLocation": {},
        "allComparableLocations": [],
        "questions": {},
        "answers": {},
        "participationData": [],
        "participationTotals": {}
    }
    data_as_json_string = json.dumps(dashboard_data, indent=4, ensure_ascii=False)
    # The script will continue and generate a minimal HTML file.
    # To stop execution here, you could use exit()


# --- 5. Estruturar os Dados para o Dashboard ---
print("Estruturando os dados para o dashboard HTML...")

all_colors = fixed_location_color_map.copy()
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
    "participationTotals": total_row_dict
}

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

    for location in all_comparable_locations:
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
        
        # *** INÍCIO DA ALTERAÇÃO ***
        # Verifica se a contagem de "Outras respostas" é maior que o número de respondentes
        show_warning = False
        other_responses_count = local_counts.get("Outras respostas", 0)
        if other_responses_count > total_local_respondents:
            show_warning = True
        
        dashboard_data["answers"][col_letter][location] = {
            "totalRespondents": total_local_respondents,
            "items": local_counts.head(15).to_dict(),
            "showOtherResponsesWarning": show_warning  # Adiciona a flag
        }
        # *** FIM DA ALTERAÇÃO ***

# --- 6. Gerar o Arquivo HTML ---
print("Gerando o arquivo HTML do dashboard...")

# Custom JSON encoder for numpy integers
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
        #participationTable th {{
            cursor: pointer;
            position: sticky;
            top: 0;
            background-color: #f8fafc;
            z-index: 10;
        }}
        #participationTable th:hover {{
            background-color: #e2e8f0;
        }}
        #participationTable .sort-asc::after {{
            content: ' ▲';
            font-size: 0.8em;
        }}
        #participationTable .sort-desc::after {{
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
                    <label for="location-select" class="block text-sm font-medium text-gray-700 mb-1">Selecione um Local para comparar:</label>
                    <select id="location-select" class="w-full p-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"></select>
                </div>
            </div>

            <div class="grid grid-cols-1 mt-8">
                <div class="bg-gray-50 p-4 rounded-xl">
                    <h3 id="local-chart-title" class="text-lg font-semibold text-center mb-1"></h3>
                    <!-- *** INÍCIO DA ALTERAÇÃO HTML *** -->
                    <p id="local-chart-subtitle" class="text-sm text-center text-red-600 italic font-bold mb-2"></p>
                    <!-- *** FIM DA ALTERAÇÃO HTML *** -->
                    <div style="height: 975px;" class="w-full">
                        <canvas id="localComparisonChart"></canvas>
                    </div>
                </div>
            </div>
            
            <div class="mt-6 flex justify-center items-center space-x-6 text-xs text-gray-600">
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
                            <th scope="col" class="py-3 px-6" data-column-index="0" data-column-type="string">Unidade</th>
                            <th scope="col" class="py-3 px-6 text-right" data-column-index="1" data-column-type="number">Respostas</th>
                            <th scope="col" class="py-3 px-6 text-right" data-column-index="2" data-column-type="number">Total esperado</th>
                            <th scope="col" class="py-3 px-6 text-right" data-column-index="3" data-column-type="percentage">Participação</th>
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
    let currentSort = {{ column: 3, order: 'desc' }};

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
            const key = Object.keys(a)[currentSort.column];
            const valA = a[key];
            const valB = b[key];
            const columnType = participationTable.querySelector(`th[data-column-index='${{currentSort.column}}']`).dataset.columnType;
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

            if (compareA < compareB) return currentSort.order === 'asc' ? -1 : 1;
            if (compareA > compareB) return currentSort.order === 'asc' ? 1 : -1;
            return 0;
        }});

        dataToSort.forEach(row => {{
            const tr = `
                <tr class="bg-white border-b hover:bg-gray-50">
                    <td class="py-4 px-6 font-medium text-gray-900 whitespace-nowrap">${{row['Unidade']}}</td>
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

        participationTable.querySelectorAll('th').forEach(th => {{
            th.classList.remove('sort-asc', 'sort-desc');
            const colIndex = parseInt(th.dataset.columnIndex, 10);
            if (colIndex === currentSort.column) {{
                th.classList.add(currentSort.order === 'asc' ? 'sort-asc' : 'sort-desc');
            }}
        }});
    }}
    
    participationTable.querySelectorAll('th').forEach(th => {{
        th.addEventListener('click', () => {{
            const columnIndex = parseInt(th.dataset.columnIndex, 10);
            if (currentSort.column === columnIndex) {{
                currentSort.order = currentSort.order === 'asc' ? 'desc' : 'asc';
            }} else {{
                currentSort.column = columnIndex;
                currentSort.order = 'desc';
            }}
            renderParticipationTable();
        }});
    }});
    
    function renderLocationResponsesChart() {{
        const ctx = document.getElementById('locationResponsesChart').getContext('2d');
        const data = surveyData.responsesByLocation;
        const labels = Object.keys(data).sort((a, b) => data[b] - data[a]);
        const values = labels.map(label => data[label]);
        const backgroundColors = labels.map(label => surveyData.locationColors[label] || '#cccccc');

        if (locationChart) locationChart.destroy();
        locationChart = new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: labels,
                datasets: [{{ label: 'Respostas', data: values, backgroundColor: backgroundColors }}]
            }},
            options: {{
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    datalabels: {{
                        display: true,
                        anchor: 'end',
                        align: 'end',
                        color: '#333',
                        font: {{ weight: 'bold', size: 12 }},
                        formatter: (value) => value.toLocaleString('pt-BR')
                    }}
                }},
                scales: {{
                    x: {{
                        type: 'logarithmic',
                        title: {{ display: true, text: 'Quantidade de Respostas (Escala Log)' }},
                    }},
                    y: {{ ticks: {{ autoSkip: false }} }}
                }}
            }}
        }});
    }}

    function renderComparisonChart(questionId, locationId) {{
        const localCtx = document.getElementById('localComparisonChart').getContext('2d');
        const localQuestionData = surveyData.answers[questionId]?.[locationId];
        const globalQuestionData = surveyData.answers[questionId]?.['Global'];
        
        // *** INÍCIO DA ALTERAÇÃO JS ***
        const subtitleElement = document.getElementById('local-chart-subtitle');
        subtitleElement.textContent = ''; // Limpa a mensagem por padrão
        // *** FIM DA ALTERAÇÃO JS ***

        if (localChart) localChart.destroy();

        if (localQuestionData && globalQuestionData) {{
            // *** INÍCIO DA ALTERAÇÃO JS ***
            // Pega a nova flag 'showOtherResponsesWarning' dos dados
            const {{ totalRespondents: totalLocal, items: localItems, showOtherResponsesWarning }} = localQuestionData;
            // *** FIM DA ALTERAÇÃO JS ***
            
            const {{ totalRespondents: totalGlobal, items: globalItems }} = globalQuestionData;

            document.getElementById('local-chart-title').textContent =
                `Comparativo: ${{locationId}} (${{totalLocal.toLocaleString('pt-BR')}} resp.) vs. Rede (${{totalGlobal.toLocaleString('pt-BR')}} resp.)`;

            // *** INÍCIO DA ALTERAÇÃO JS ***
            // Mostra a mensagem se a flag for verdadeira
            if (showOtherResponsesWarning) {{
                subtitleElement.textContent = `Como cada respondente podia selecionar até 3 itens por eixo e escrever sua própria sugestão, a porcentagem de "Outras respostas" pode ultrapassar 100%`;
            }}
            // *** FIM DA ALTERAÇÃO JS ***

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
                data: {{
                    labels: labels,
                    datasets: [
                        {{
                            label: locationId,
                            data: localPercentages,
                            backgroundColor: surveyData.locationColors[locationId] || '#333',
                            absoluteCounts: localAbsoluteCounts
                        }},
                        {{
                            label: 'Rede',
                            data: globalPercentagesForCompare,
                            backgroundColor: '#a0aec0',
                        }}
                    ]
                }},
                options: getChartOptions('y',`% de Respostas (${{locationId}} vs. Rede)`,
                    (ctx) => `${{ctx.dataset.label}}: ${{ctx.raw.toFixed(1)}}%`,
                    {{
                        formatter: (value, context) => {{
                            const dataset = context.chart.data.datasets[context.datasetIndex];
                            if(dataset.label !== 'Rede') {{
                                const count = dataset.absoluteCounts[context.dataIndex];
                                return `${{count}} (${{value.toFixed(1)}}%)`;
                            }}
                            return `${{value.toFixed(1)}}%`;
                        }},
                    }}, tickColors, suggestedMax
                )
            }});
        }} else {{
            document.getElementById('local-chart-title').textContent = `Comparativo Local vs. Rede`;
            localCtx.clearRect(0, 0, localCtx.canvas.width, localCtx.canvas.height);
            localCtx.font = "16px Inter";
            localCtx.fillStyle = "#888";
            localCtx.textAlign = "center";
            localCtx.fillText(`Sem dados para '${{locationId}}' neste eixo.`, localCtx.canvas.width / 2, localCtx.canvas.height / 2);
        }}
    }}
    
    function getChartOptions(axis, xAxisTitle, tooltipLabelCallback, datalabelsConfig, yAxisTickColors = '#666', suggestedMax = 100) {{
        return {{
            indexAxis: axis,
            responsive: true,
            maintainAspectRatio: false,
            layout: {{ padding: {{ left: 0 }} }},
            plugins: {{
                legend: {{ position: 'top' }},
                tooltip: {{ callbacks: {{ label: tooltipLabelCallback }} }},
                datalabels: {{ display: true, anchor: 'end', align: 'end', color: '#333', font: {{ weight: 'bold', size: 12 }}, ...datalabelsConfig }}
            }},
            scales: {{
                x: {{
                    suggestedMax: suggestedMax,
                    title: {{ display: true, text: xAxisTitle }},
                    ticks: {{ callback: value => value.toFixed(0) + "%" }}
                }},
                y: {{
                    ticks: {{
                        color: yAxisTickColors,
                        callback: function(value) {{ return wrapText(this.getLabelForValue(value), 70); }}
                    }}
                }}
            }}
        }};
    }}

    function populateSelects() {{
        Object.entries(surveyData.questions).forEach(([id, title]) => {{
            const option = document.createElement('option');
            option.value = id;
            option.textContent = title;
            questionSelect.appendChild(option);
        }});

        const locations = surveyData.allComparableLocations;
        const smeDefault = 'SME';
        let finalLocations;

        if (locations.includes(smeDefault)) {{
            const otherLocations = locations.filter(loc => loc !== smeDefault).sort((a, b) => a.localeCompare(b));
            finalLocations = [smeDefault, ...otherLocations];
        }} else {{
            finalLocations = locations.sort((a, b) => a.localeCompare(b));
        }}
        
        finalLocations.forEach(location => {{
            const option = document.createElement('option');
            option.value = location;
            option.textContent = location;
            locationSelect.appendChild(option);
        }});
    }}

    function updateCharts() {{
        const selectedQuestion = questionSelect.value;
        const selectedLocation = locationSelect.value;
        renderComparisonChart(selectedQuestion, selectedLocation);
    }}
    
    function init() {{
        populateSelects();
        questionSelect.addEventListener('change', updateCharts);
        locationSelect.addEventListener('change', updateCharts);
        renderLocationResponsesChart();
        renderParticipationTable();
        updateCharts();
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
