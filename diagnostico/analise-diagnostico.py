# Importar bibliotecas essenciais
import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import seaborn as sns
import textwrap # Importar a biblioteca textwrap

# --- Configurações ---
# Define o caminho base para os arquivos do projeto
base_path = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Questionário"

file_name = "respostas-diagnostico.xlsx"
file_path = os.path.join(base_path, file_name)

# Define o diretório de saída para os gráficos
output_chart_dir = r"C:\Users\PAULOSEIKISHIHIGA\OneDrive - Secretaria Municipal de São Paulo\UPGE\1.2 Planejamento\1.2.1 Planejamento Estratégico\Diagnóstico\Relatório\Gráficos"
# Cria o diretório de saída se ele não existir
os.makedirs(output_chart_dir, exist_ok=True)

# Dicionário para mapear os nomes completos dos locais para rótulos curtos
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

# Dicionário de cores fixas para cada local de trabalho (rótulos curtos)
# Cores escolhidas para serem vivas e terem contraste razoável com texto branco.
fixed_location_color_map = {
    "SME": "#D32F2F",    # Vermelho Escuro (Material Design Red 700)
    "DRE-BT": "#1976D2", # Azul Escuro (Material Design Blue 700)
    "DRE-CL": "#388E3C", # Verde Escuro (Material Design Green 700)
    "DRE-CS": "#F57C00", # Laranja Escuro (Material Design Orange 700)
    "DRE-FB": "#7B1FA2", # Roxo (Material Design Purple 700)
    "DRE-G": "#00796B",  # Teal Escuro (Material Design Teal 700)
    "DRE-IP": "#E64A19", # Laranja Avermelhado Profundo (Material Design Deep Orange 700)
    "DRE-IQ": "#5D4037", # Marrom Escuro (Material Design Brown 700)
    "DRE-JT": "#C2185B", # Rosa Escuro (Material Design Pink 700)
    "DRE-PE": "#0288D1", # Azul Claro Escuro (Material Design Light Blue 700)
    "DRE-PJ": "#AFB42B", # Lima Escuro (Material Design Lime 700)
    "DRE-SA": "#FFA000", # Âmbar Escuro (Material Design Amber 700)
    "DRE-SM": "#616161", # Cinza Escuro (Material Design Grey 700)
    "DRE-MP": "#455A64"  # Azul Acinzentado Escuro (Material Design Blue Grey 700)
}

# Coluna de Local de Trabalho (G) e colunas das perguntas (I a W)
location_col_letter = 'G'
# Gera uma lista de letras de 'I' a 'W'
question_col_letters = [chr(i) for i in range(ord('I'), ord('V') + 1)] # Alterado para terminar em 'V'

# Mapeamento de letras de coluna para índices baseados em 0 para uso no pandas read_excel
# A=0, B=1, ..., G=6, I=8, ..., W=22
location_col_index = ord(location_col_letter) - ord('A')
question_col_indices = [ord(letter) - ord('A') for letter in question_col_letters]

# Lista de índices das colunas a serem lidas
use_cols_indices = [location_col_index] + question_col_indices

# --- 1. Ler e Preparar os Dados ---
# Primeiro, ler a primeira linha para obter os títulos das perguntas
try:
    # Ler apenas a primeira linha para obter os títulos das colunas
    header_df = pd.read_excel(file_path,
                              usecols=use_cols_indices,
                              nrows=1,
                              header=None)

    # Criar um dicionário mapeando os índices lidos para os títulos das perguntas
    # O título da coluna de local de trabalho será o nome que definimos ('Local de trabalho')
    # Os títulos das perguntas virão da primeira linha do Excel
    column_titles = {location_col_index: 'Local de trabalho'}
    # O header_df tem apenas as colunas especificadas em use_cols_indices
    # A primeira coluna em header_df corresponde a location_col_index
    # As colunas seguintes correspondem a question_col_indices na mesma ordem
    for i, col_index in enumerate(question_col_indices):
         # O índice da coluna no header_df será 1 + i (0 é a coluna de local de trabalho)
         column_titles[col_index] = header_df.iloc[0, 1 + i] # Pega o valor da primeira linha (índice 0) e coluna correspondente

    # Ler apenas as colunas necessárias, pulando a primeira linha (cabeçalho original)
    df = pd.read_excel(file_path,
                       usecols=use_cols_indices,
                       header=None, # Indica que não há cabeçalho nos dados que estamos lendo (pulamos a primeira linha)
                       skiprows=1) # Pula a primeira linha (que é o cabeçalho original do Excel)

    # Renomear as colunas usando os títulos obtidos
    # Cria um dicionário mapeando os índices lidos (que são os nomes das colunas no df lido) para os títulos desejados
    df.rename(columns=column_titles, inplace=True)

    print(f"Arquivo Excel '{file_name}' lido com sucesso do caminho base. Títulos das perguntas lidos.")
    print(f"Colunas lidas: 'Local de trabalho' (original '{location_col_letter}') e perguntas (originais '{question_col_letters[0]}' a '{question_col_letters[-1]}').")
    print("Primeiras linhas do DataFrame lido:")
    print(df.head())

except FileNotFoundError:
    print(f"Erro: O arquivo Excel '{file_path}' não foi encontrado.")
    print("Verifique se o caminho base e o nome do arquivo estão corretos.")
    exit()
except ValueError as e:
    print(f"Erro ao ler as colunas do Excel: {e}")
    print(f"Verifique se o arquivo Excel contém as colunas {location_col_letter} e {question_col_letters[0]} a {question_col_letters[-1]} e se 'skiprows=1' está correto para pular o cabeçalho.")
    exit()
except Exception as e:
    print(f"Ocorreu um erro inesperado ao ler o arquivo Excel: {e}")
    exit()

# Limpeza básica: remover linhas onde o Local de trabalho é nulo ou vazio
df.dropna(subset=['Local de trabalho'], inplace=True)
df = df[df['Local de trabalho'].astype(str).str.strip() != '']

# Aplicar o mapeamento de rótulos aos locais de trabalho
# Usa .map() e .fillna() para manter o nome original caso não esteja no dicionário
# (embora com a lista fornecida, todos os nomes esperados devem estar lá)
df['Local de trabalho'] = df['Local de trabalho'].map(label_mapping).fillna(df['Local de trabalho'])


if df.empty:
    print("Não foram encontrados dados válidos de 'Local de trabalho' após a limpeza.")
    exit()

# Obter a lista de locais de trabalho únicos após a limpeza
unique_locations = df['Local de trabalho'].unique()
print(f"\nLocais de trabalho únicos encontrados ({len(unique_locations)}):")
print(unique_locations)

# --- Função Auxiliar para Agregação ---
def aggregate_low_counts(item_counts_series, threshold=5, other_label="Outras respostas"):
    """
    Agrega itens em uma Series (value_counts) com contagens abaixo de um limiar
    em uma única categoria 'other_label'.
    """
    # Itens com contagem abaixo do limiar
    low_count_items = item_counts_series[item_counts_series < threshold]
    # Itens com contagem igual ou acima do limiar
    high_count_items = item_counts_series[item_counts_series >= threshold]

    aggregated_counts = high_count_items.copy()

    if not low_count_items.empty:
        other_sum = low_count_items.sum()
        if other_sum > 0:
            if other_label in aggregated_counts: # Caso 'Outras respostas' já exista (improvável com threshold > 1)
                aggregated_counts[other_label] += other_sum
            else:
                aggregated_counts[other_label] = other_sum
    
    return aggregated_counts.sort_values(ascending=False) # Ordena para que o top N funcione corretamente

# --- 2. Gerar Gráfico de Barras para Quantidade de Respostas por Local de Trabalho ---
print("\nGerando gráfico de respostas por Local de Trabalho...")
# Contar a frequência de cada local de trabalho e ordenar pelo nome do local
location_counts = df['Local de trabalho'].value_counts().sort_index()
location_counts = location_counts[location_counts > 0] # Remover locais com 0 respostas após limpeza
total_responses_location_chart = location_counts.sum() # Calcular a soma total para o xlabel

if location_counts.empty:
    print("Nenhum local de trabalho com respostas encontrado após a limpeza. Pulando gráfico de locais.")
else:
    median_val = location_counts.median()
    outlier_cutoff = 1.5 * median_val

    actual_outliers = location_counts[location_counts > outlier_cutoff]
    non_outliers = location_counts[location_counts <= outlier_cutoff]

    should_break = False
    xlim1_end, xlim2_start = 0, 0 # Initialize

    if not actual_outliers.empty and not non_outliers.empty:
        xlim1_end_candidate = non_outliers.max() * 1.2 
        # Ensure the end of the first plot is at least a bit beyond the general outlier cutoff
        xlim1_end_candidate = max(xlim1_end_candidate, outlier_cutoff * 1.1) 
        xlim2_start_candidate = actual_outliers.min() * 0.95

        # Condition for a meaningful break:
        # 1. The calculated end of the first plot is before the start of the second.
        # 2. There's a significant visual gap between the largest non-outlier and smallest outlier.
        if xlim1_end_candidate < xlim2_start_candidate and actual_outliers.min() > (non_outliers.max() * 1.5):
            should_break = True
            xlim1_end = xlim1_end_candidate
            xlim2_start = xlim2_start_candidate

    if should_break:
        print("Detectado outlier significativo. Gerando gráfico com eixo quebrado.")
        # Define width ratios for the two subplots
        # These ratios can be adjusted for better visual balance
        width_ratios = [0.7, 0.3] 
        fig, (ax1, ax2) = plt.subplots(1, 2, sharey=True, 
                                       figsize=(15, max(5, len(location_counts) * 0.55)), 
                                       gridspec_kw={'width_ratios': width_ratios, 'wspace': 0.05})

        # Plot on ax1 (left part)
        sns.barplot(ax=ax1, y=location_counts.index, x=location_counts.values, hue=location_counts.index, palette=fixed_location_color_map, orient='h', legend=False)
        ax1.set_xlim(0, xlim1_end)
        ax1.spines['right'].set_visible(False)
        ax1.grid(axis='x', linestyle='--', alpha=0.7)
        ax1.set_ylabel('Local de Trabalho', fontsize=12)
        ax1.tick_params(axis='y', labelsize=10) # Tamanho da fonte alterado para 10

        # Plot on ax2 (right part, for outliers)
        sns.barplot(ax=ax2, y=location_counts.index, x=location_counts.values, hue=location_counts.index, palette=fixed_location_color_map, orient='h', legend=False)
        ax2.set_xlim(xlim2_start, location_counts.max() * 1.05) # Add padding to max
        ax2.spines['left'].set_visible(False)
        ax2.tick_params(axis='y', which='both', left=False, labelleft=False)
        ax2.set_ylabel('') # No y-label for the right plot
        ax2.grid(axis='x', linestyle='--', alpha=0.7)

        # Add break marks (//)
        # d is the proportion of vertical to horizontal extent of the slanted line for the marker
        d_marker = .5 
        kwargs_marker = dict(marker=[(-1, -d_marker), (1, d_marker)], markersize=12, linestyle="none", color='k', mec='k', mew=1, clip_on=False)
        ax1.plot([1, 1], [0,1],transform=ax1.transAxes, **kwargs_marker) # On right spine of ax1
        ax2.plot([0, 0], [0,1],transform=ax2.transAxes, **kwargs_marker) # On left spine of ax2

        # Bar labels - label based on original value
        for i in range(len(location_counts)):
            count = location_counts.values[i]
            # Label on ax1 if the bar is primarily in ax1 or fully contained
            if count < xlim2_start * 0.98 : # If it's not a "truly broken" bar needing ax2 for its end
                if ax1.containers[i].patches and ax1.containers[i].patches[0].get_width() > 0.01 : # Check if drawn
                    ax1.bar_label(ax1.containers[i], labels=[f'{int(count)}'], padding=3, fontsize=9)
            # Label on ax2 if the bar is an outlier that extends into ax2
            else: 
                if ax2.containers[i].patches and ax2.containers[i].patches[0].get_width() > 0.01: # Check if drawn
                    ax2.bar_label(ax2.containers[i], labels=[f'{int(count)}'], padding=3, fontsize=9)
        
        fig.suptitle('Quantidade de Respostas por Local de Trabalho', fontsize=14, y=0.98) # Removido "(com quebra de eixo)"
        fig.text(0.5, 0.02, f'Quantidade de Respostas (Total: {total_responses_location_chart})', ha='center', va='center', fontsize=12) # Adicionado total
        fig.tight_layout(rect=[0, 0.03, 1, 0.95]) # Adjust rect for suptitle and fig.text

    else: # Original plotting logic if no break is needed or not effective
        if not actual_outliers.empty:
            print("Outliers detectados, mas a quebra de eixo não foi considerada ideal/necessária.")
        fig = plt.figure(figsize=(12, max(5, len(location_counts) * 0.5)))
        ax = sns.barplot(
            y=location_counts.index,
            x=location_counts.values,
            hue=location_counts.index, 
            palette=fixed_location_color_map, # Usar cores fixas
            orient='h',
            legend=False) 
        plt.suptitle('Quantidade de Respostas por Local de Trabalho', fontsize=14) 
        plt.xlabel(f'Quantidade de Respostas (Total: {total_responses_location_chart})', fontsize=12) # Adicionado total
        plt.ylabel('Local de Trabalho', fontsize=12)
        plt.yticks(fontsize=10) # Tamanho da fonte alterado para 10
        plt.grid(axis='x', linestyle='--', alpha=0.7) 
        
        # Adicionar rótulos de contagem nas barras
        for container in ax.containers:
            if container.patches: 
                ax.bar_label(container, fmt='%d', padding=3, fontsize=9)
        plt.tight_layout(rect=[0, 0.03, 1, 0.94]) 

    output_location_chart_path = os.path.join(output_chart_dir, "grafico_respostas_por_local_trabalho.png")
    fig.savefig(output_location_chart_path) # Use fig.savefig()
    print(f"Gráfico salvo como '{output_location_chart_path}'")
    plt.close(fig) # Fecha a figura para liberar memória


# --- 3. Gerar Gráficos de Barras para Contagem de Itens por Pergunta (Global) ---
print("\nGerando gráficos de contagem de itens por pergunta (Global)...")

for col_letter in question_col_letters:
    # Obter o título da pergunta usando a letra da coluna
    # Usamos .get() com um fallback caso a letra da coluna não esteja no dicionário (improvável aqui)
    question_title = column_titles.get(ord(col_letter) - ord('A'), f"Pergunta {col_letter}")

    print(f"\nProcessando coluna '{col_letter}' ('{question_title}') para contagem global...")

    # Processar a coluna: converter para string, remover espaços, dividir por ';', e "explodir"
    # Substituir NaN por string vazia antes de processar
    items = df[question_title].astype(str).str.strip().replace('', np.nan).dropna() # Remove NaNs e strings vazias

    if items.empty:
        print(f"Nenhuma resposta válida encontrada na coluna '{col_letter}' ('{question_title}') após limpeza. Pulando contagem global.")
        continue

    # Número total de respondentes para esta pergunta (antes de explodir as respostas múltiplas)
    total_respondents_for_question = len(items)

    # Agora, para os valores não nulos, podemos aplicar split e explode
    # Tratar múltiplos separadores (ex: "a;;b") - manter espaços
    items = items.str.split(';').explode()

    # Remover quaisquer strings vazias que possam ter resultado do split (ex: "a;;b")
    items = items[items != '']
    items = items.dropna() # Remover NaNs que possam ter surgido após explode/filtragem

    # Correção manual de item específico
    items = items.replace(
        "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família",
        "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias"
    )
    if items.empty:
        print(f"Nenhum item válido encontrado na coluna '{col_letter}' ('{question_title}') após processamento. Pulando contagem global.")
        continue

    # Contar a frequência de cada item
    raw_item_counts = items.value_counts()

    # Agregar itens com menos de 5 respostas em "Outras respostas"
    # A função aggregate_low_counts já ordena por padrão
    aggregated_item_counts = aggregate_low_counts(raw_item_counts, threshold=5, other_label="Outras respostas")

    # Selecionar os top 10 itens (ou menos, se houver menos de 10 categorias após agregação)
    top_10_items = aggregated_item_counts.head(10)

    if top_10_items.empty:
        print(f"Nenhum item top 10 encontrado na coluna '{col_letter}' ('{question_title}') após contagem. Pulando contagem global.")
        continue

    # Definir um tamanho fixo para a figura
    plt.figure(figsize=(14, 8)) # Largura de 14 polegadas, altura de 8 polegadas

    # Address FutureWarning: Assign y to hue
    ax = sns.barplot(
        y=top_10_items.index,
        x=top_10_items.values,
        hue=top_10_items.index, # Assign y to hue
        palette='viridis',
        orient='h',
        legend=False # Disable legend as hue is same as y
    )
    plt.suptitle(f'{question_title} (Rede)', fontsize=14) # Substituído "(Top...Global)" por "Rede"
    plt.xlabel('Quantidade de Respostas', fontsize=12)
    plt.ylabel('Item de Resposta', fontsize=10) # Alterado para uma única linha
    plt.grid(axis='x', linestyle='--', alpha=0.7) # Adiciona grade no eixo X

    # Quebrar rótulos longos do eixo Y (yticks)
    wrap_width = 85 # Retornado para 85

    # Address UserWarning for set_yticklabels by setting fixed ticks
    # The y-values for barplot are top_10_items.index. These are the labels.
    # The tick positions are typically 0, 1, ..., len(top_10_items)-1
    tick_positions = np.arange(len(top_10_items.index))
    wrapped_yticklabels = [textwrap.fill(str(label), width=wrap_width, break_long_words=False, replace_whitespace=False) 
                           for label in top_10_items.index]
    
    ax.set_yticks(tick_positions)
    ax.set_yticklabels(wrapped_yticklabels, fontsize=10) # Changed to 10
    
    # Ajustar o layout para que os rótulos dos itens (eixo Y) ocupem a metade esquerda
    # e as barras ocupem a metade direita.
    # Reduzir 'top' para dar espaço ao suptitle.
    # top=0.93: A área de plotagem termina em 93% da altura da figura.
    plt.subplots_adjust(left=0.5, right=0.98, top=0.93, bottom=0.1)
    # Adicionar rótulos de contagem nas barras
    # Modificado para incluir a porcentagem em relação ao total de respondentes da pergunta
    # Prepare custom_labels once
    custom_labels = []
    for count_value in top_10_items.values: # Estes são os valores das barras
        percentage = (count_value / total_respondents_for_question) * 100 if total_respondents_for_question > 0 else 0
        custom_labels.append(f'{int(count_value)} ({percentage:.1f}%)')

    # If hue is used, ax.containers is a list of containers, each with one bar.
    # Iterate through containers and provide the specific label for that bar.
    for i, container_for_single_bar in enumerate(ax.containers):
        # Each container_for_single_bar should have exactly one patch (bar)
        if container_for_single_bar.patches:
            patch = container_for_single_bar.patches[0] # Get the single patch
            count_value = patch.get_width() # The width of the bar is the count
            label_text = custom_labels[i] # Get the corresponding label text

            # Conditional placement and color
            if count_value < 300: # Changed threshold from 200 to 300
                # Place label outside the bar (to the right)
                ax.text(patch.get_width() + 5, # Add a small positive padding
                        patch.get_y() + patch.get_height() / 2.0, # Vertical center
                        label_text,
                        ha='left', va='center', # Align text to the left of the position
                        color='black', fontsize=14, fontweight='bold')
            else:
                # Place label inside the bar
                ax.text(patch.get_width() - 50, # Negative padding to move inside
                        patch.get_y() + patch.get_height() / 2.0, # Vertical center
                        label_text,
                        ha='right', va='center', # Align text to the right of the position
                        color='white', fontsize=14, fontweight='bold')

    # Salvar no novo diretório de gráficos
    output_item_count_chart_path = os.path.join(output_chart_dir, f"grafico_pergunta_{col_letter}_contagem_global.png")
    plt.savefig(output_item_count_chart_path)
    print(f"Gráfico salvo como '{output_item_count_chart_path}'")
    plt.close() # Fecha a figura

# --- 4. Gerar Gráficos de Contagem de Itens por Pergunta e Local de Trabalho ---
print("\nGerando gráficos de contagem de itens por pergunta e por local de trabalho...")
output_local_chart_dir = os.path.join(output_chart_dir, "Por_Local")
os.makedirs(output_local_chart_dir, exist_ok=True)

# Pre-calculate global percentages for all items for each question
global_data_for_questions = {}
print("\nPré-calculando dados globais de referência para gráficos por local...")
for col_letter_global_calc in question_col_letters:
    question_title_global_calc = column_titles.get(ord(col_letter_global_calc) - ord('A'), f"Pergunta {col_letter_global_calc}")
    
    global_items_series_calc = df[question_title_global_calc].astype(str).str.strip().replace('', np.nan).dropna()
    if global_items_series_calc.empty:
        global_data_for_questions[question_title_global_calc] = {'percentages': pd.Series(dtype=float), 'total_respondents': 0}
        continue
        
    total_global_respondents_calc = len(global_items_series_calc)
    
    global_items_processed_calc = global_items_series_calc.str.split(';').explode()
    global_items_processed_calc = global_items_processed_calc[global_items_processed_calc != ''].str.strip().dropna()
    
    # Correção manual de item específico para dados globais de referência
    global_items_processed_calc = global_items_processed_calc.replace(
        "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família",
        "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias"
    )

    if global_items_processed_calc.empty:
        global_data_for_questions[question_title_global_calc] = {'percentages': pd.Series(dtype=float), 'total_respondents': total_global_respondents_calc}
        continue
        
    raw_global_item_counts_calc = global_items_processed_calc.value_counts()
    # Aplicar agregação aos dados globais de referência, usando o mesmo threshold e rótulo
    # que é usado para a agregação local e para os gráficos globais principais.
    aggregated_global_item_counts_for_ref = aggregate_low_counts(raw_global_item_counts_calc, threshold=5, other_label="Outras respostas")
    
    # Calcular percentuais com base nas contagens globais agregadas
    global_item_percentages_calc = (aggregated_global_item_counts_for_ref / total_global_respondents_calc) * 100 if total_global_respondents_calc > 0 else pd.Series(dtype=float)
    global_data_for_questions[question_title_global_calc] = {'percentages': global_item_percentages_calc, 'total_respondents': total_global_respondents_calc}

# Usar o mapa de cores fixas definido no início do script para os gráficos por local
location_color_map = fixed_location_color_map

for col_letter in question_col_letters: # Defined as 'I' to 'V'
    question_title = column_titles.get(ord(col_letter) - ord('A'), f"Pergunta {col_letter}")

    for location_label in unique_locations: # unique_locations contains mapped short labels
        print(f"\nProcessando coluna '{col_letter}' ('{question_title}') para o local '{location_label}'...")

        df_location = df[df['Local de trabalho'] == location_label].copy()

        if df_location.empty:
            print(f"Nenhum dado encontrado para o local '{location_label}' na pergunta '{question_title}'. Pulando.")
            continue

        # Processar a coluna para este local
        local_items_series = df_location[question_title].astype(str).str.strip().replace('', np.nan).dropna()

        if local_items_series.empty:
            print(f"Nenhuma resposta válida na coluna '{question_title}' para o local '{location_label}'. Pulando.")
            continue

        total_respondents_for_question_at_location = len(local_items_series)

        local_items_processed = local_items_series.str.split(';').explode()
        local_items_processed = local_items_processed[local_items_processed != ''].str.strip() # Remove empty strings and strip
        local_items_processed = local_items_processed.dropna()

        # Correção manual de item específico para dados locais
        local_items_processed = local_items_processed.replace(
            "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família",
            "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias"
        )

        if local_items_processed.empty:
            print(f"Nenhum item válido na coluna '{question_title}' para o local '{location_label}' após processamento. Pulando.")
            continue

        raw_local_item_counts = local_items_processed.value_counts()
        # Aggregate low counts (e.g., items with < 5 responses for this specific location/question)
        aggregated_local_item_counts = aggregate_low_counts(raw_local_item_counts, threshold=5, other_label="Outras respostas")
        
        top_10_local_items = aggregated_local_item_counts.head(10)

        if top_10_local_items.empty:
            print(f"Nenhum item top 10 encontrado na coluna '{question_title}' para o local '{location_label}'. Pulando.")
            continue

        # --- Plotting ---
        fig_local, ax_local = plt.subplots(figsize=(14, 9)) # Adjusted height slightly for two bars

        num_items = len(top_10_local_items)
        y_positions = np.arange(num_items)
        
        bar_height_main = 0.35
        bar_height_comparison = 0.18 # Reduced height for "Rede" bar
        gap_between_bars = 0.0 # Remove gap to make bars touch

        local_percentages_values = []
        global_percentages_values_for_plot = []
        local_counts_values = top_10_local_items.values

        current_global_data = global_data_for_questions.get(question_title, {'percentages': pd.Series(dtype=float)})
        global_item_percentages_for_this_question = current_global_data['percentages']

        for item_name, local_count_val in top_10_local_items.items():
            local_perc = (local_count_val / total_respondents_for_question_at_location) * 100 if total_respondents_for_question_at_location > 0 else 0
            local_percentages_values.append(local_perc)
            
            global_perc = global_item_percentages_for_this_question.get(item_name, 0.0)
            global_percentages_values_for_plot.append(global_perc)

        # Define colors
        color_local = location_color_map.get(location_label, 'grey') # Get unique color for the location
        color_global_avg = 'silver' # Cor para a barra "Rede", para destacar as cores dos locais

        # Plot local bars (main - top bar for each item)
        bars_local = ax_local.barh(y_positions + bar_height_comparison/2 + gap_between_bars, local_percentages_values, 
                                   height=bar_height_main, color=color_local, label=location_label) # Changed color and label
        # Plot global comparison bars (comparison - bottom bar for each item)
        bars_global_avg = ax_local.barh(y_positions - bar_height_main/2 - gap_between_bars, global_percentages_values_for_plot, 
                                        height=bar_height_comparison, color=color_global_avg, label='Rede')

        ax_local.set_yticks(y_positions)
        wrap_width_local = 85
        
        wrapped_yticklabels_local = []
        any_item_highlighted = False
        tick_color_map = {} # Para armazenar a cor de cada ytick destacado {index: 'red'/'green'}
        for i, item_name_tick in enumerate(top_10_local_items.index):
            local_p = local_percentages_values[i]
            global_p = global_percentages_values_for_plot[i]
            prefix = ""
            color_for_tick = None

            if item_name_tick != "Outras respostas": # Não aplicar destaque para "Outras respostas"
                # Critério de destaque: diferença relativa de 20%
                # Lidar com global_p == 0 para evitar divisão por zero
                if global_p == 0:
                    if local_p > 0: # Se global é 0 e local não, considera-se uma grande diferença relativa
                        prefix = "* "
                        any_item_highlighted = True
                        color_for_tick = 'red' # Local é maior
                else: # global_p > 0
                    relative_diff = (local_p - global_p) / global_p
                    if relative_diff >= 0.20: # Local é 20% ou mais MAIOR
                        prefix = "* "
                        any_item_highlighted = True
                        color_for_tick = 'red'
                    elif relative_diff <= -0.20: # Local é 20% ou mais MENOR
                        prefix = "* "
                        any_item_highlighted = True
                        color_for_tick = 'green'
            
            if color_for_tick:
                tick_color_map[i] = color_for_tick
            wrapped_yticklabels_local.append(prefix + textwrap.fill(str(item_name_tick), width=wrap_width_local, break_long_words=False, replace_whitespace=False))

        tick_positions_local = np.arange(len(top_10_local_items.index)) # Mantém a mesma lógica para posições
        ax_local.set_yticks(tick_positions_local)
        ax_local.set_yticklabels(wrapped_yticklabels_local, fontsize=10) # Changed to 10

        # Colorir os yticks destacados em vermelho
        # É importante fazer isso APÓS set_yticklabels
        for i, tick_label in enumerate(ax_local.get_yticklabels()):
            if i in tick_color_map:
                tick_label.set_color(tick_color_map[i])

        # Set title for the entire figure, centered
        fig_local.suptitle(f'{question_title} ({location_label})', fontsize=14, y=0.96) # y adjusts vertical position
        ax_local.set_xlabel('Porcentagem de Respostas (%)', fontsize=12)
        ax_local.set_ylabel('Item de Resposta', fontsize=10)
        ax_local.grid(axis='x', linestyle='--', alpha=0.7)
        ax_local.legend(loc='best') # Adjust legend location if needed
        ax_local.invert_yaxis() # Ensure most frequent items are at the top
        
        # Adjust bottom margin if a note about highlighted items is added
        bottom_margin = 0.15 if any_item_highlighted else 0.1
        fig_local.subplots_adjust(left=0.5, right=0.98, top=0.90, bottom=bottom_margin) # Ensure space for suptitle and note


        # Bar labels
        label_fontsize = 9
        local_perc_threshold_for_inner_label = 20 # Percentage threshold

        # Label local bars (count + local percentage)
        for i, bar in enumerate(bars_local):
            local_perc_val = bar.get_width()
            local_count_val = local_counts_values[i]
            label_text = f'{int(local_count_val)} ({local_perc_val:.1f}%)'
            
            if local_perc_val < local_perc_threshold_for_inner_label or len(label_text) > 5 and local_perc_val < local_perc_threshold_for_inner_label + 10 :
                ax_local.text(local_perc_val + 0.5, bar.get_y() + bar.get_height()/2, label_text,
                              ha='left', va='center', color='black', fontsize=label_fontsize, fontweight='bold')
            else:
                ax_local.text(local_perc_val - 0.5, bar.get_y() + bar.get_height()/2, label_text,
                              ha='right', va='center', color='white', fontsize=label_fontsize, fontweight='bold')

        # Label global average bars (global percentage)
        for bar in bars_global_avg:
            global_perc_val = bar.get_width()
            if global_perc_val > 0.05: # Only label if there's a visible bar
                label_text = f'{global_perc_val:.1f}%'
                ax_local.text(global_perc_val + 0.5, bar.get_y() + bar.get_height()/2, label_text,
                              ha='left', va='center', color='dimgray', fontsize=label_fontsize - 1)
        
        if any_item_highlighted:
            y_pos = 0.03
            fontsize = 10 # Increased font size
            renderer = fig_local.canvas.get_renderer()

            text_parts_data = [
                ("* Diferença relativa >= 20% em relação à Rede (", "dimgray"),
                ("vermelho", "red"),
                (f": {location_label} > Rede; ", "dimgray"),
                ("verde", "green"),
                (f": {location_label} < Rede)", "dimgray")
            ]

            # Calculate total width in figure coordinates to center the block
            total_width_fig_coords = 0
            for text_segment, _ in text_parts_data:
                temp_text_obj = fig_local.text(0, 0, text_segment, fontsize=fontsize, visible=False)
                bbox = temp_text_obj.get_window_extent(renderer=renderer)
                segment_width_fig_coords = bbox.width / (fig_local.get_figwidth() * fig_local.dpi)
                total_width_fig_coords += segment_width_fig_coords
                temp_text_obj.remove()

            current_x = 0.5 - (total_width_fig_coords / 2) # Starting x to center the whole block

            for text_segment, color in text_parts_data:
                text_obj = fig_local.text(current_x, y_pos, text_segment,
                                          ha="left", va="bottom", fontsize=fontsize, color=color)
                
                # Update current_x for the next segment based on the width of the segment just plotted
                bbox = text_obj.get_window_extent(renderer=renderer)
                # Ensure bbox width is not None or zero before division
                segment_width_display_coords = bbox.width if bbox.width else 0
                segment_width_fig_coords = segment_width_display_coords / (fig_local.get_figwidth() * fig_local.dpi)
                current_x += segment_width_fig_coords

        output_filename = f"grafico_pergunta_{col_letter}_local_{location_label}.png"
        output_local_chart_path = os.path.join(output_local_chart_dir, output_filename)
        fig_local.savefig(output_local_chart_path)
        print(f"Gráfico salvo como '{output_local_chart_path}'")
        plt.close(fig_local) # Close figure

print("\nProcessamento de gráficos concluído.")

# --- 5. Gerar Planilha Excel com Itens Destacados ---
print("\nGerando planilha Excel com itens destacados (diferença absoluta >= 10 p.p. da Rede)...")
highlighted_items_data = []

for col_letter_excel in question_col_letters:
    question_title_excel = column_titles.get(ord(col_letter_excel) - ord('A'), f"Pergunta {col_letter_excel}")
    current_global_data_excel = global_data_for_questions.get(question_title_excel, {'percentages': pd.Series(dtype=float), 'total_respondents': 0})
    global_item_percentages_excel = current_global_data_excel['percentages']

    # Se não há dados globais válidos para esta pergunta, não podemos comparar.
    if global_item_percentages_excel.empty and current_global_data_excel['total_respondents'] == 0:
        print(f"Sem dados globais para a pergunta '{question_title_excel}'. Pulando para a planilha.")
        continue

    for location_label_excel in unique_locations:
        df_location_excel = df[df['Local de trabalho'] == location_label_excel].copy()
        if df_location_excel.empty:
            continue # Pula se não houver dados para este local

        local_items_series_excel = df_location_excel[question_title_excel].astype(str).str.strip().replace('', np.nan).dropna()
        if local_items_series_excel.empty:
            continue # Pula se não houver respostas para esta pergunta neste local
        
        total_respondents_local_excel = len(local_items_series_excel)

        local_items_processed_excel = local_items_series_excel.str.split(';').explode()
        local_items_processed_excel = local_items_processed_excel[local_items_processed_excel != ''].str.strip().dropna()
        
        # Aplicar correção manual de item específico
        local_items_processed_excel = local_items_processed_excel.replace(
            "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas família",
            "Ampliação de matrículas em unidades mais próximas dos endereços de referência fornecidos pelas famílias"
        )

        if local_items_processed_excel.empty:
            continue
            
        raw_local_item_counts_excel = local_items_processed_excel.value_counts()
        # Usar a mesma agregação que para os gráficos
        aggregated_local_item_counts_excel = aggregate_low_counts(raw_local_item_counts_excel, threshold=5, other_label="Outras respostas")

        for item_name_excel, local_count_excel in aggregated_local_item_counts_excel.items():
            if item_name_excel == "Outras respostas": # Não incluir "Outras respostas" na análise de destaque
                continue

            local_perc_excel = (local_count_excel / total_respondents_local_excel) * 100 if total_respondents_local_excel > 0 else 0
            global_perc_excel = global_item_percentages_excel.get(item_name_excel, 0.0) # Pega a % global para este item
            
            percentage_point_diff = local_perc_excel - global_perc_excel
            
            # Novo critério: diferença absoluta de 8 pontos percentuais ou mais
            is_highlighted = abs(percentage_point_diff) >= 8
            
            if is_highlighted:
                highlighted_items_data.append({
                    'Local': location_label_excel,
                    'Eixo': question_title_excel,
                    'Item': item_name_excel,
                    '% Local': round(local_perc_excel, 2),
                    '% Rede': round(global_perc_excel, 2),
                    'Diferença (p.p)': round(percentage_point_diff, 2)
                })

# Criar DataFrame e salvar em Excel
if highlighted_items_data:
    highlighted_df = pd.DataFrame(highlighted_items_data)
    excel_output_path = os.path.join(output_chart_dir, "itens_destacados_diagnostico.xlsx") # Salva no mesmo dir dos gráficos
    
    try:
        highlighted_df.to_excel(excel_output_path, index=False, sheet_name="Itens Destacados")
        print(f"Planilha Excel com itens destacados salva como '{excel_output_path}'")
    except Exception as e:
        print(f"Erro ao salvar a planilha Excel: {e}")
else:
    print("Nenhum item destacado encontrado para incluir na planilha Excel.")
