import pandas as pd
import os

def concatenar_arquivos_excel(caminho_arquivo1, caminho_arquivo2):
    try:
        # Carregar ambos os arquivos Excel em DataFrames
        df1 = pd.read_excel(caminho_arquivo1)
        df2 = pd.read_excel(caminho_arquivo2)

        df1 = df1.rename(columns={
            "Centro localização": "Centro de manutenção",
            "Centro trab.respons.": "CenTrabalho princ.",
            "Grp.plnj.PM": "Grupo planej."            
        })

        # Concatenar os DataFrames verticalmente
        df_concatenado = pd.concat([df1, df2], ignore_index=True)

        # Exibir informações sobre o DataFrame concatenado
        print("DataFrames concatenados com sucesso:")
        print(df_concatenado.head())

        return df_concatenado

    except Exception as e:
        print(f"Erro ao carregar ou concatenar os arquivos Excel: {str(e)}")
        return None

def salvar_dataframe_excel(dataframe, caminho_saida):
    try:
        # Extrair o diretório pai do caminho de saída
        diretorio_saida = os.path.dirname(caminho_saida)

        # Verificar se o diretório de saída existe; se não existir, criar
        if not os.path.exists(diretorio_saida):
            os.makedirs(diretorio_saida)
            print(f"Diretório '{diretorio_saida}' criado com sucesso.")

        # Salvar o DataFrame em um arquivo Excel no caminho de saída especificado
        dataframe.to_excel(caminho_saida, index=False)
        print(f"DataFrame salvo com sucesso em '{caminho_saida}'")

    except Exception as e:
        print(f"Erro ao salvar o DataFrame em arquivo Excel: {str(e)}")

# Caminhos completos para os arquivos Excel a serem concatenados
caminho_arquivo1 = r"Z:\#BASES DE APOIO - GERENCIAL\01 - BASES\PROJETOS\Controle Contabil\SAP\IW39\IW39.xlsx"
caminho_arquivo2 = r"Z:\#BASES DE APOIO - GERENCIAL\01 - BASES\PROJETOS\Controle Contabil\SAP\IW39\EXPORT.xlsx"  # Substitua com o caminho do segundo arquivo

# Chamar a função para concatenar os arquivos Excel
df_concatenado = concatenar_arquivos_excel(caminho_arquivo1, caminho_arquivo2)

# Se o DataFrame concatenado for válido, salvar em um arquivo Excel no diretório específico
if df_concatenado is not None:
    # Caminho completo para o arquivo Excel de saída
    caminho_saida = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_IW39.xlsx"

    # Chamar a função para salvar o DataFrame em um arquivo Excel no caminho de saída específico
    salvar_dataframe_excel(df_concatenado, caminho_saida)
