Projeto ETL para Análise de Dados no Power BI
Este projeto consiste em um processo de ETL (Extract, Transform, Load) desenvolvido em Python para a análise de dados no Power BI. O objetivo principal é gerar uma tabela consolidada e limpa a partir de diversos arquivos de entrada, reduzindo assim a complexidade e o tempo de processamento necessário no Power BI.

Descrição do Projeto:
Desenvolvedor Responsável: Glauber Marques, Analista de Dados da área de Planejamento.
Intuito do Projeto: Migrar o processo de ETL para Python, visando gerar uma tabela mais limpa e reduzir o número de etapas de processamento no Power BI, economizando recursos de processamento.
Fluxo do Projeto:
Extração: Os dados são extraídos de diferentes fontes, como planilhas Excel.
Transformação: São aplicadas diversas transformações nos dados, como seleção de colunas, junção de DataFrames e aplicação de condições para classificação.
Carga: Os dados transformados são carregados em um único DataFrame e salvos em um arquivo Excel, pronto para ser utilizado no Power BI.
Benefícios:
Redução da complexidade no Power BI: Com a tabela já tratada e consolidada, há menos manipulações necessárias no Power BI, resultando em um processo mais eficiente.
Economia de recursos de processamento: Ao realizar as transformações e limpezas no Python, há uma redução no tempo de processamento e no uso de recursos no Power BI, melhorando a performance da análise de dados.
Estrutura do Projeto:
Arquivos de Entrada: São arquivos Excel contendo diferentes conjuntos de dados.
Arquivo Python (ETL): Contém o código Python responsável por extrair, transformar e carregar os dados.
Arquivo de Saída: Tabela final resultante do processo de ETL, salva em formato Excel para utilização no Power BI.
Observações:
O código Python utiliza a biblioteca Pandas para manipulação e análise de dados.
As etapas de transformação são realizadas de forma eficiente, visando otimizar o processo como um todo.
Qualquer atualização nos dados de entrada pode ser facilmente incorporada ao processo, mantendo a consistência e a eficácia da análise no Power BI.
