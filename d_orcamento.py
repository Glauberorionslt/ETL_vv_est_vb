import pandas as pd
import os

def transformar_tabela(caminho_arquivo_entrada, caminho_arquivo_saida):
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(caminho_arquivo_entrada, sheet_name="ORÇAMENTO CLONE", skiprows=23)

        # Converter todas as células em valores de texto (string)
        df = df.astype(str)

        # Transformar as colunas de meses em linhas correspondentes
        df_transformado = pd.melt(df, id_vars=["Empresa", "Pacote", "Sub-pacote", "Conta Contábil", "Centro de Custo", "Cedente", "Status", "Diretor", "Regional", "Coordenador", "ORÇADO Total/22"],var_name='Data',value_name="Total"
        )
        
        # Converter tipo da coluna "Data" para datetime
        df_transformado["Data"] = pd.to_datetime(df_transformado["Data"], format="%d/%m/%Y")

        colunas_desejadas = ['Conta Contábil','Cedente','Data','Total']

        df_transformado = df_transformado[colunas_desejadas]

        #verificar se arquivo existe no diretorio

        # Verificar se o arquivo de saída já existe
        if os.path.exists(caminho_arquivo_saida):
            print(f"Arquivo '{caminho_arquivo_saida}' já existe. Sobrepondo o arquivo existente...")
            os.remove(caminho_arquivo_saida)  # Remove o arquivo existente

        # Salvar o DataFrame resultante em um novo arquivo Excel
        df_transformado.to_excel(caminho_arquivo_saida, index=False)

        print(f"Arquivo transformado salvo com sucesso em: {caminho_arquivo_saida}")

    except Exception as e:
        print(f"Erro ao transformar o arquivo: {str(e)}")

# Caminhos dos arquivos de entrada e saída
caminho_arquivo_entrada = r"Z:\#BASES DE APOIO - GERENCIAL\01 - BASES\PROJETOS\Orçamento -Centro de Custo\f_verba_ajust_REGIONAL(2024).xlsx"
caminho_arquivo_saida = r"C:\Users\2160011883\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\d_orcamento.xlsx"

# Chamar a função para transformar a tabela e salvar o arquivo
transformar_tabela(caminho_arquivo_entrada, caminho_arquivo_saida)



