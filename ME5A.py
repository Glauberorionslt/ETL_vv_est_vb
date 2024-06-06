import pandas as pd
import os

def processar_arquivo_me5a(caminho_arquivo_entrada, caminho_arquivo_saida):
    try:
        # Carregar o arquivo ME5A.xlsx em um DataFrame
        df = pd.read_excel(caminho_arquivo_entrada)

        # Selecionar apenas as colunas desejadas
        df = df[['Requisição de compra', 'Nome de fornecedor']]

        # Remover valores duplicados da coluna 'Requisição de compra'
        df = df.drop_duplicates(subset=['Requisição de compra'])

        # Exibir informações sobre o DataFrame resultante
        print("DataFrame processado com sucesso:")

        # Verificar se o arquivo de saída já existe
        if os.path.exists(caminho_arquivo_saida):
            print(f"Arquivo '{caminho_arquivo_saida}' já existe. Sobrepondo o arquivo existente...")
            os.remove(caminho_arquivo_saida)  # Remover o arquivo existente

        # Salvar o DataFrame com valores distintos em um arquivo Excel
        df.to_excel(caminho_arquivo_saida, index=False)
        print(f"DataFrame salvo com sucesso em '{caminho_arquivo_saida}'")

    except Exception as e:
        print(f"Erro ao processar o arquivo ME5A.xlsx: {str(e)}")


# Caminho completo para o arquivo ME5A.xlsx
caminho_arquivo_entrada = r"Z:\#BASES DE APOIO - GERENCIAL\01 - BASES\PROJETOS\Controle Contabil\ME5A\EXPORT.xlsx"

# Caminho completo para o arquivo de saída f_ME5A.xlsx
caminho_arquivo_saida = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_ME5A.xlsx"

# Chamar a função para processar o arquivo ME5A.xlsx e salvar o resultado
processar_arquivo_me5a(caminho_arquivo_entrada, caminho_arquivo_saida)
