import pandas as pd
import os

# Carregar os arquivos
caminho_iw39 = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_IW39.xlsx"
caminho_me5k = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_ME5K.xlsx"
caminho_IW29 = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_IW29.xlsx"
caminho_me5k = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_ME5K.xlsx"
caminho_IW29 = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_IW29.xlsx"
caminho_ME5A = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_ME5A.xlsx"
caminho_SN = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\GLAUBER S-ONE DRIVE\PROJETOS\PYTHON\Analise de dados\ETL\Pbi\Estudo de verbas\f_Servicenow_combinado.xlsx"



# Colunas a serem selecionadas de cada arquivo
colunas_iw39 = ['Ordem','Texto breve', 'Data de entrada', 'Custos estimados','Centro de manutenção','Status do sistema']
colunas_me5k = ['Ordem','Requisição de compra','Pedido']
colunas_iw29 =['Ordem','Texto de grupo de codificação','Texto code para codificação']
colunas_SN =['Número OM','Item','Prioridade']
colunas_ME5A =['Requisição de compra','Nome de fornecedor']

# Carregar os dados dos arquivos
df_iw39 = pd.read_excel(caminho_iw39,usecols=colunas_iw39)
df_me5k = pd.read_excel(caminho_me5k, usecols=colunas_me5k)
df_iw29 = pd.read_excel(caminho_IW29, usecols=colunas_iw29)
df_SN = pd.read_excel(caminho_SN, usecols=colunas_SN)
df_SN =df_SN.rename(columns={'Número OM':'Ordem'})
df_ME5A = pd.read_excel(caminho_ME5A, usecols=colunas_ME5A)


# Combinação dos dados com base na coluna "Ordem"
df_combinado = pd.merge(df_iw39, df_me5k, on='Ordem', how='left')
df_combinado1 = pd.merge(df_combinado, df_iw29, on='Ordem', how='left')
df_combinado2 = pd.merge(df_combinado1, df_SN, on='Ordem', how='left')
df_combinado3 = pd.merge(df_combinado2, df_ME5A, on='Requisição de compra', how='left')

cond_RITM = df_combinado3["Texto breve"].str.startswith('RITM')
df_combinado3.loc[cond_RITM,"RITM do Texto Breve"] = df_combinado3.loc[cond_RITM,"Texto breve"].str.slice(0,11)

df_combinado3['RITM condicional'] = df_combinado3.apply(lambda row: row['Item'] if pd.notnull(row['Item']) else row['RITM do Texto Breve'], axis=1)


# Definir as condições e suas respectivas classificações
condicoes = [
    (df_combinado3['Texto breve'].str.contains('ZM', case=False, na=False), 'MANUTENÇÃO'),
    (df_combinado3['Texto breve'].str.contains('ZP', case=False, na=False), 'PLANEJAMENTO'),
    (df_combinado3['Texto breve'].str.contains('ZC', case=False, na=False), 'COLABORATIVO'),
    (df_combinado3['Texto breve'].str.contains('ZS', case=False, na=False), 'SINISTRO'),
    (df_combinado3['Texto breve'].str.contains('ZL', case=False, na=False), 'PPCI'),
    (df_combinado3['Texto breve'].str.contains('PASSIVO', case=False, na=False), 'PASSIVO'),
    (df_combinado3['Texto breve'].str.contains('ZE', case=False, na=False), 'ENCERRAMENTO')
]

# Definir o valor padrão caso nenhuma condição seja atendida
valor_padrao = 'S/CLASSIFICAÇÃO'

# Aplicar as condições usando np.select()
df_combinado3['Classificacao'] = pd.Series(pd.NA, dtype='string')  # Inicializar a coluna com NA
df_combinado3['Classificacao'] = df_combinado3['Classificacao'].astype('string')  # Garantir que a coluna tenha dtype string

# Loop pelas condições e definindo a classificação
for condicao, classificacao in condicoes:
    df_combinado3['Classificacao'] = df_combinado3['Classificacao'].mask(condicao, classificacao)

# Preencher os valores que não foram classificados com o valor padrão
df_combinado3['Classificacao'].fillna(valor_padrao, inplace=True)



# Exibir as primeiras linhas do DataFrame combinado
print(df_combinado3.head())

# Salvar o DataFrame combinado em um arquivo
caminho_saida = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\Bases - DB_MANUTENCAO\tabela_fato.xlsx"

if os.path.exists(caminho_saida):
    os.remove(caminho_saida)

df_combinado3.to_excel(caminho_saida, index=False)
print("Tabela fato salva com sucesso!")




