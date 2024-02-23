import pandas as pd
from unidecode import unidecode

# Carregar a planilha
planilha = pd.read_excel(r"C:\Users\henri\Documents\python_projects\RPA\RPA\src\Relacao_Produtos_e_Clientes_2024.xlsx")

# Listas de palavras-chave e suas respectivas substituições
metodos_substituicoes = {
    "Cred.": "Cartão de Crédito",
    "Débito": "Cartão de Débito",
    "Transferência": "Transferência Bancária",
    "Dinheiro": "Dinheiro",
    "Cheque": "Cheque"
}

# Função para reescrever o método de pagamento
def reescrever_metodo_pagamento(metodo_pagamento, substituicoes):
    metodo_pagamento_padronizado = metodo_pagamento
    for chave, valor in substituicoes.items():
        if chave in metodo_pagamento_padronizado:
            metodo_pagamento_padronizado = valor
    return metodo_pagamento_padronizado

# Lista para armazenar os métodos de pagamento padronizados
lista_metodo_pagamento_padronizado = []

# Coluna onde os métodos de pagamento estão localizados
coluna_metodo_pagamento = 'Método de Pagamento'

# Iterar sobre as linhas da coluna de métodos de pagamento
for linha in range(2, planilha.shape[0] + 1):
    metodo_pagamento = planilha.loc[linha - 1, coluna_metodo_pagamento]
    metodo_pagamento_padronizado = reescrever_metodo_pagamento(metodo_pagamento, metodos_substituicoes)
    lista_metodo_pagamento_padronizado.append(metodo_pagamento_padronizado)

# Exibir a lista de métodos de pagamento padronizados
print(lista_metodo_pagamento_padronizado)
