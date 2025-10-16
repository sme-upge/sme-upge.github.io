import csv
import io

# Este script lê o arquivo CSV não padronizado e o salva em um formato limpo.
# Ele lida com quebras de linha dentro das células e usa um delimitador padrão.
try:
    with io.open('arvore.csv', 'r', encoding='utf-8') as infile, \
         io.open('arvore_limpa.csv', 'w', encoding='utf-8', newline='') as outfile:
        
        # O leitor está configurado para o delimitador de ponto e vírgula e aspas.
        reader = csv.reader(infile, delimiter=';', quotechar='"')
        
        # O gravador usará vírgula como delimitador e colocará aspas em todos os campos
        # para garantir a compatibilidade com o interpretador do navegador.
        writer = csv.writer(outfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        
        # Lê cada linha do arquivo de entrada e a escreve no arquivo de saída.
        for row in reader:
            writer.writerow(row)
            
    print("Arquivo 'arvore_limpa.csv' criado com sucesso.")

except FileNotFoundError:
    print("Erro: O arquivo 'arvore.csv' não foi encontrado.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")

