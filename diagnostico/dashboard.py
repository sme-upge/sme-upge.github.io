# Importar bibliotecas essenciais
import pandas as pd
import numpy as np
import os
import json

# --- 1. Configurações ---
# Define o caminho base e o arquivo de entrada
base_path = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Questionário"
file_name = "respostas-diagnostico-anon.xlsx"
file_path = os.path.join(base_path, file_name)

# Define o diretório de saída e o nome do arquivo HTML final
output_dir = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Relatório"
output_html_filename = "dashboard.html"
output_html_path = os.path.join(output_dir, output_html_filename)

# Cria o diretório de saída se ele não existir
os.makedirs(output_dir, exist_ok=True)

# Dicionários de mapeamento e cores
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

# Configuração das colunas
location_col_letter = 'G'
question_col_letters = [chr(i) for i in range(ord('I'), ord('V') + 1)]

location_col_index = ord(location_col_letter) - ord('A')
question_col_indices = [ord(letter) - ord('A') for letter in question_col_letters]
use_cols_indices = [location_col_index] + question_col_indices

# --- 2. Função Auxiliar para Agregação ---
def aggregate_low_counts(item_counts_series, threshold=5, other_label="Outras respostas"):
    low_count_items = item_counts_series[item_counts_series < threshold]
    high_count_items = item_counts_series[item_counts_series >= threshold]
    aggregated_counts = high_count_items.copy()
    if not low_count_items.empty:
        other_sum = low_count_items.sum()
        if other_sum > 0:
            aggregated_counts[other_label] = aggregated_counts.get(other_label, 0) + other_sum
    return aggregated_counts.sort_values(ascending=False)

# --- 3. Ler e Preparar os Dados ---
print("Iniciando leitura e processamento dos dados do Excel...")
try:
    header_df = pd.read_excel(file_path, usecols=use_cols_indices, nrows=1, header=None)
    
    column_titles = {location_col_index: 'Local de trabalho'}
    question_titles_map = {}
    for i, col_index in enumerate(question_col_indices):
        title = header_df.iloc[0, 1 + i]
        column_titles[col_index] = title
        question_titles_map[question_col_letters[i]] = title

    df = pd.read_excel(file_path, usecols=use_cols_indices, header=None, skiprows=1)
    df.rename(columns=column_titles, inplace=True)
    
    print("Arquivo Excel lido com sucesso.")

    df.dropna(subset=['Local de trabalho'], inplace=True)
    df = df[df['Local de trabalho'].astype(str).str.strip() != '']
    df['Local de trabalho'] = df['Local de trabalho'].map(label_mapping).fillna(df['Local de trabalho'])

    if df.empty:
        raise ValueError("Nenhum dado válido de 'Local de trabalho' encontrado após a limpeza.")

    unique_locations = sorted(list(df['Local de trabalho'].unique()))

except Exception as e:
    print(f"Erro crítico ao processar o arquivo Excel: {e}")
    exit()

# --- 4. Estruturar os Dados para o Dashboard ---
print("Estruturando os dados para o dashboard HTML...")
dashboard_data = {
    "locationColors": fixed_location_color_map,
    "responsesByLocation": {},
    "questions": {},
    "answers": {}
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

    for location in unique_locations:
        df_location = df[df['Local de trabalho'] == location]
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
        dashboard_data["answers"][col_letter][location] = {
            "totalRespondents": total_local_respondents,
            "items": local_counts.head(15).to_dict()
        }

# --- 5. Gerar o Arquivo HTML ---
print("Gerando o arquivo HTML do dashboard...")

data_as_json_string = json.dumps(dashboard_data, indent=4, ensure_ascii=False)

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
    </style>
</head>
<body class="text-gray-800">

    <div class="container mx-auto p-4 md:p-8">
        <header class="mb-8 flex justify-between items-center">
            <div>
                <h1 class="text-3xl md:text-4xl font-bold text-gray-800">Diagnóstico da Educação Paulistana</h1>
                <p class="text-md text-gray-600 mt-2">Análise interativa das respostas do questionário.</p>
            </div>
            <img src="logo-pe.png" alt="Logotipo do Planejamento Estratégico 2025-2028" style="max-width: 200px; height: auto;">
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
                    <h3 id="local-chart-title" class="text-lg font-semibold text-center mb-2"></h3>
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
        
        <section class="bg-white p-6 rounded-2xl shadow-lg">
            <h2 class="text-2xl font-semibold mb-4">Respostas por Unidade</h2>
            <div style="height: 450px;" class="w-full">
                <canvas id="locationResponsesChart"></canvas>
            </div>
        </section>
        
        <footer class="text-center mt-12 text-gray-500 text-sm">
            <p>Secretaria Municipal de Educação / Planejamento Estratégico 2025-2028</p>
        </footer>
    </div>

<script>
// --- INÍCIO DOS DADOS EMBUTIDOS ---
const surveyData = {data_as_json_string};
// --- FIM DOS DADOS EMBUTIDOS ---

Chart.register(ChartDataLabels);

function wrapText(text, maxWidth) {{
    if (typeof text !== 'string') {{
        return '';
    }}
    if (text.length <= maxWidth) {{
        return text;
    }}
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
    if (currentLine) {{
        lines.push(currentLine.trim());
    }}
    return lines;
}}

document.addEventListener('DOMContentLoaded', () => {{
    let locationChart, localChart;

    const questionSelect = document.getElementById('question-select');
    const locationSelect = document.getElementById('location-select');

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
                datasets: [{{
                    label: 'Respostas',
                    data: values,
                    backgroundColor: backgroundColors
                }}]
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
                }}
            }}
        }});
    }}

    function renderComparisonChart(questionId, locationId) {{
        const localCtx = document.getElementById('localComparisonChart').getContext('2d');
        const localQuestionData = surveyData.answers[questionId]?.[locationId];
        const globalQuestionData = surveyData.answers[questionId]?.['Global'];

        if (localChart) localChart.destroy();

        if (localQuestionData && globalQuestionData) {{
            const {{ totalRespondents: totalLocal, items: localItems }} = localQuestionData;
            const {{ totalRespondents: totalGlobal, items: globalItems }} = globalQuestionData;

            document.getElementById('local-chart-title').textContent = 
                `Comparativo: ${{locationId}} (${{totalLocal.toLocaleString('pt-BR')}} resp.) vs. Rede (${{totalGlobal.toLocaleString('pt-BR')}} resp.)`;

            const sortedLocalItems = Object.entries(localItems).sort(([, a], [, b]) => b - a);
            const labels = sortedLocalItems.map(([label]) => label);
            const localAbsoluteCounts = sortedLocalItems.map(([, value]) => value);

            const tickColors = [];
            const localPercentages = [];
            const globalPercentagesForCompare = [];

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
            const suggestedMax = Math.ceil(maxPercentage) + 10;

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
                options: getChartOptions(
                    'y',
                    `% de Respostas (${{locationId}} vs. Rede)`,
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
                     }},
                    tickColors,
                    suggestedMax
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
            layout: {{
                padding: {{
                    left: 0
                }}
            }},
            plugins: {{
                legend: {{ position: 'top' }},
                tooltip: {{ callbacks: {{ label: tooltipLabelCallback }} }},
                datalabels: {{
                    display: true,
                    anchor: 'end',
                    align: 'end',
                    color: '#333',
                    font: {{ weight: 'bold', size: 12 }},
                    ...datalabelsConfig
                }}
            }},
            scales: {{
                x: {{
                    max: suggestedMax,
                    title: {{ display: true, text: xAxisTitle }},
                    ticks: {{ callback: value => value.toFixed(0) + "%" }}
                }},
                y: {{
                    ticks: {{
                        color: yAxisTickColors,
                        callback: function(value) {{
                            const label = this.getLabelForValue(value);
                            return wrapText(label, 70);
                        }}
                    }}
                }}
            }}
        }};
    }}

    function populateSelects() {{
        // ALTERAÇÃO: Mantém a ordem original das perguntas da planilha, sem ordenar alfabeticamente.
        const questionsInOrder = Object.entries(surveyData.questions);
        questionsInOrder.forEach(([id, title]) => {{
            const option = document.createElement('option');
            option.value = id;
            option.textContent = title;
            questionSelect.appendChild(option);
        }});

        // ALTERAÇÃO: Coloca "SME" como a primeira opção padrão.
        const locations = Object.keys(surveyData.responsesByLocation);
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
        updateCharts();
    }}

    init();
}});
</script>

</body>
</html>
"""

# Escreve o conteúdo final no arquivo HTML
try:
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_template)
    print(f"\nDashboard gerado com sucesso!")
    print(f"Arquivo salvo em: {output_html_path}")
except Exception as e:
    print(f"Erro ao salvar o arquivo HTML: {e}")
