# <p style="color:green">Relatório Automatizado de Vendas por Loja</p>

Este script em **Python** automatiza a geração de um relatório de vendas por loja a partir de um arquivo Excel de dados de vendas. O script realiza as seguintes etapas:

1. <span style="color:blue">**Importação e Preparação dos Dados:**</span>
   - Utiliza a biblioteca `pandas` para importar os dados de vendas de um arquivo Excel.
   - Configura a visualização da base de dados para exibir todas as colunas.

2. <span style="color:blue">**Análise de Faturamento por Loja:**</span>
   - Agrupa os dados de vendas por ID da loja e calcula o faturamento total de cada loja.

3. <span style="color:blue">**Análise da Quantidade de Produtos Vendidos por Loja:**</span>
   - Agrupa os dados de vendas por ID da loja e calcula a quantidade total de produtos vendidos em cada loja.

4. <span style="color:blue">**Cálculo do Ticket Médio por Produto em Cada Loja:**</span>
   - Calcula o ticket médio (faturamento médio por produto) para cada loja.

5. <span style="color:blue">**Envio de Relatório por Email:**</span>
   - Utiliza a biblioteca `win32com` para interagir com o Microsoft Outlook.
   - Cria um email com o relatório em HTML incorporado, incluindo:
     - Faturamento por loja.
     - Quantidade de produtos vendidos por loja.
     - Ticket médio dos produtos em cada loja.

6. <span style="color:blue">**Requisitos:**</span>
   - Python 3.x
   - Bibliotecas: pandas, win32com (para interagir com o Outlook)

7. <span style="color:blue">**Instruções de Uso:**</span>
   - Certifique-se de ter instalado todas as bibliotecas necessárias.
   - Tenha o arquivo Excel "Vendas.xlsx" na mesma pasta que o script.
   - Execute o script e verifique o Outlook para o envio do relatório por email.

Para executar este script, é necessário ter o Python instalado juntamente com as bibliotecas `pandas` e `win32com`. Certifique-se de ter permissões para interagir com o Outlook para o envio automatizado de emails.
