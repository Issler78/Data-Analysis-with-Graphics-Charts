# Análise de Dados com Gráficos

> [!NOTE]
> A ideia do projeto surgiu de apenas um prompt simples no ChatGPT (nada sério, apenas para me testar na minha primeira vez com gráficos em python 😄).

Esse programa tem a ideia principal de realizar uma análise de dados completa a partir de uma tabela excel com 10000 linhas sobre vendas.
No final da execução do script, você terá um relatório em excel, com as seguintes páginas:
- Relatório Geral - Terá informações completas, como o valor total das vendas ou o vendedor com mais vendas;
- Vendas Diárias - Terá a quantidade de produtos vendidos e o valor total das vendas de cada **dia** presente na tabela;
- Vendas por Produto - Terá a quantidade vendida e o valor total das vendas de cada **produto**;
- Vendas por Vendedor - Terá a quantidade vendida e o valor total das vendas de cada **vendedor**;
- Vendas por Localização - Terá a quantidade vendida e o valor total das vendas de cada **localização/loja**;
  
- Gráficos - Terá 3 (três) gráficos com base na tabela. São eles:
  - Valor Total das Vendas por Produto - Um gráfico de barras, com os produtos e seus respectivos lucros;
  - Distribuição das Vendas por Localização - Um gráfico de pizza, com as lojas que participaram das vendas e as suas respectivas porcentagem do quanto cada uma contribuiu;
  - Evolução da Quantidade de Vendas - Um gráfico de linhas, mostrando a evolução da quantidade vendida no período da tabela.
 
> [!TIP]
> Pode ser utilizado o Task Scheduler (windows) ou o Cron (Linux) para automatizar e rodar o script no intervalo que você decidir.

## Funcionalidades e Características

- Entrega de um relatório completo com a análise de toda a tabela base;
- Código todo documentado, fácil de entender e de modificar;
- Relatório formatado para fácil leitura;
- Gráficos para um relatório visual.

## Requisitos

- Python 3
- Instalar as Bibliotecas **Pandas**, **Openpyxl**, **Matplotlib** e **Python-Dotenv**.

  No Windows e Linux:
  ```
  pip install pandas openpyxl matplotlib python-dotenv
  
  ```

  No Mac:
  ```
  pip3 install pandas openpyxl matplotlib python-dotenv
  
  ```

  ## Screenshots

    Tabela base (input)
    
    <img src="https://github.com/user-attachments/assets/ff3b7d49-265a-4a7d-9e47-d6847a9a1cec" alt="Tabela base (input)" width="300" />
    
  <hr>
    
    Tabela do relatório gerado (output)
    
    <img src="https://github.com/user-attachments/assets/5aefe836-13bf-409f-afa3-115e9ca3067b" alt="Tabela do relatório gerado" width="600"/>
    
  <hr>
    
    Gráfico da evolução das vendas
    
    <img src="https://github.com/user-attachments/assets/c130a2ff-61b2-472c-b4ee-b840d58ed0d1" alt="Tabela do relatório gerado" width="600"/>
