import pandas as pd
import os
import matplotlib.pyplot as plt

# Função para salvar os dados como PDF
def salvar_pdf(df, nome_arquivo, movimento, tecnico):
    plt.figure(figsize=(10, 6))
    plt.axis('off')  # Remove os eixos
    plt.title(f"Relatório de {movimento} - Técnico: {tecnico}", fontsize=14, weight='bold')

    # Cria uma tabela com os dados
    tabela = plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center', fontsize=12)
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(8)

    # Salva como PDF
    plt.savefig(nome_arquivo, bbox_inches='tight', pad_inches=0.5)
    plt.close()

# Caminho do arquivo Excel
arquivo_excel = 'marcelo-entrega-devolucao.xlsx'

# Colunas selecionadas
colunas_selecionadas = ['sku', 'qtd', 'descricao', 'atlas', 'tecnico', 'data', 'movimento']

# Carregar dados da planilha Excel
df = pd.read_excel(arquivo_excel, usecols=colunas_selecionadas)

# Cria pastas de saída
output_folder = 'pdfs_movimentos'
os.makedirs(output_folder, exist_ok=True)

# Converte a coluna 'data' para o formato de data (caso não esteja)
df['data'] = pd.to_datetime(df['data'], errors='coerce').dt.date

# Processa os movimentos "ENTREGA" e "DEVOLUCAO" separadamente
for movimento in ['ENTREGA', 'DEVOLUCAO']:
    df_movimento = df[df['movimento'] == movimento]
    
    # Filtra os dados por técnico
    for tecnico, grupo in df_movimento.groupby('tecnico'):
        # Converte a coluna 'data' para datetime e formata como DD-MM-AAAA
        grupo['data'] = pd.to_datetime(grupo['data'], errors='coerce').dt.strftime('%d-%m-%Y')
        
        # Pega a data do primeiro registro formatada como DD-MM-AAAA
        data = grupo['data'].iloc[0]
        
        # Define o nome do arquivo com base no movimento, técnico e data
        nome_arquivo = os.path.join(output_folder, f"{movimento.lower()}-{tecnico}-{data}.pdf")
        
        salvar_pdf(grupo, nome_arquivo, movimento, tecnico)
        print(f"PDF gerado: {nome_arquivo}")

print("Processamento concluído. PDFs salvos na pasta:", output_folder)