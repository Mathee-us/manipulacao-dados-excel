# Análise de Cotações e Geração de Planilhas com Python

Este projeto tem como objetivo processar cotações de ações e gerar planilhas com dados históricos e gráficos, utilizando Python e a biblioteca openpyxl. A aplicação realiza a leitura de dados a partir de arquivos, calcula médias móveis e bandas de volatilidade, e gera uma planilha contendo os dados e um gráfico ilustrativo.

## Funcionalidades

- **Leitura de Dados:** Utiliza a classe `Leitor` para ler arquivos de cotações do diretório `./dados/`.
- **Processamento dos Dados:** Calcula médias móveis e bandas (superior e inferior) com base em 20 pontos.
- **Geração de Planilhas:** Cria planilhas para armazenar os dados históricos e o gráfico utilizando a classe `GerenciadorPlanilha`.
- **Criação de Gráficos:** Constrói um gráfico de linhas representando as cotações e as bandas, com personalização de estilo e cores definidas.
- **Tratamento de Erros:** Implementa tratamento para erros como arquivo não encontrado ou formato incorreto.

## Tecnologias Utilizadas

- **Python 3.x**
- **openpyxl:** Para criação e manipulação de planilhas Excel.
- **Pillow (PIL):** Para manipulação de imagens (se necessário).
- **datetime:** Para manipulação de datas.

*Observação:* O projeto depende de classes customizadas (`Leitor`, `PropriedadesGrafico`, `GerenciadorPlanilha`), que devem estar implementadas no diretório `classes`.

