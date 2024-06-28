# Projeto de Análise de Vendas

Este projeto consiste em ler uma planilha de vendas de produtos, lidar com dados falhos, e gerar relatórios das vendas. O objetivo é tratar os dados, corrigir informações inconsistentes, e gerar insights sobre as vendas realizadas.

## Funcionalidades

- Leitura e tratamento de dados de uma planilha Excel.
- Validação e correção de dados faltantes ou inconsistentes.
- Geração de relatórios e gráficos sobre as vendas.
- Envio de relatórios por e-mail.

## Requisitos

- Python 3.x
- Pandas
- Dateutil
- Unidecode
- Matplotlib
- Seaborn
- Openpyxl
- Smtplib

## Instalação

1. Clone o repositório do projeto:
    ```sh
    git clone https://github.com/imrickss/projeto-analise-vendas.git
    ```

2. Navegue até o diretório do projeto:
    ```sh
    cd projeto-analise-vendas
    ```

3. Crie e ative um ambiente virtual:
    ```sh
    python -m venv venv
    source venv/bin/activate   # No Windows, use `venv\Scripts\activate`
    ```

4. Instale as dependências:
    ```sh
    pip install -r requirements.txt
    ```

## Como Usar

1. Certifique-se de que a planilha de vendas (`Relacao_Produtos_e_Clientes_2024.xlsx`) esteja no caminho correto.

2. Execute o script principal:
    ```sh
    python app.py
    ```

3. O script irá ler a planilha, tratar os dados e gerar relatórios em formato de gráfico.

4. Os gráficos serão salvos localmente e enviados por e-mail conforme configurado no script.

## Estrutura do Projeto

- `script_an
