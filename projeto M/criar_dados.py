import pandas as pd

# Criando dados fictícios com a estrutura que definimos
dados = {
    'nome_completo': [
        'Cecília Mendes', 
        'João da Silva', 
        'Maria Oliveira', 
        'Pedro Santos', 
        'Ana Costa'
    ],
    'data_nascimento': [
        '15/05/2001', 
        '20/11/1985', 
        '02/03/1990', 
        '10/07/1978', 
        '25/12/1995'
    ],
    'banco': [
        'Banco do Brasil', 
        'Nubank', 
        'Santander', 
        'Itaú', 
        'Bradesco'
    ],
    'data_ultima_aparicao': [
        '10/01/2026', 
        '15/12/2025', 
        '05/01/2026', 
        '20/12/2025', 
        '18/01/2026'
    ],
    'data_inclusao': [
        '01/06/2025', 
        '20/05/2024', 
        '10/08/2025', 
        '15/03/2023', 
        '01/11/2025'
    ],
    'email_banco': [ # E-mails fictícios para teste
        'sac@bb.com.br', 
        'meajuda@nubank.com.br', 
        'sac@santander.com.br', 
        'contato@itau.com.br', 
        'sac@bradesco.com.br'
    ]
}

# Transformando o dicionário em uma Tabela (DataFrame)
df = pd.DataFrame(dados)

# Salvando como arquivo Excel
nome_arquivo = "dados.xlsx"
df.to_excel(nome_arquivo, index=False)

print(f"✅ Planilha '{nome_arquivo}' criada com sucesso!")
print("Agora você pode rodar o script de geração de PDF.")