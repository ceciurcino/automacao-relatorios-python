# 📊 Automação de Relatórios Financeiros Seguros

Este projeto é uma solução completa de **Backend** e **Automação** desenvolvida em Python. O sistema processa dados financeiros de clientes, gera documentos PDF protegidos por senha (criptografia AES-128) e realiza o envio automático por e-mail.

O objetivo principal foi simular um cenário real de **LGPD** (Lei Geral de Proteção de Dados), onde informações sensíveis do Banco Central (Registrato) precisam ser trafegadas com segurança.

## 🚀 Funcionalidades

* **Leitura de Dados:** Extração automática de informações de planilhas Excel (`pandas`).
* **Geração de Documentos:** Criação dinâmica de PDFs formatados (`fpdf`).
* **Segurança da Informação:** Criptografia dos PDFs usando a data de nascimento do cliente como chave de acesso (`pypdf`).
* **Disparo de E-mails:** Envio automatizado via protocolo SMTP (Gmail), com tratamento de erros e anexos.
* **Organização de Arquivos:** Gerenciamento automático de pastas e nomes de arquivos.

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3.12+
* **Análise de Dados:** Pandas, OpenPyXL
* **Manipulação de PDF:** FPDF, PyPDF (Criptografia)
* **Automação de E-mail:** Smtplib (Nativo)

## 📦 Como rodar este projeto

### Pré-requisitos
Certifique-se de ter o Python instalado em sua máquina.

### 1. Instalação das bibliotecas
```bash
pip install pandas openpyxl fpdf pypdf

2. Configuração
Para rodar o script de envio de e-mails, é necessário configurar uma Senha de App do Google para garantir a segurança da conta.

Nota: Por questões de segurança, as credenciais não estão incluídas no repositório. Configure as variáveis MEU_EMAIL e MINHA_SENHA no arquivo automacao_completa.py.

3. Execução
```bash
python automacao_completa.py

📝 Estrutura dos Dados (Exemplo)
A planilha de entrada (dados.xlsx) deve seguir este formato:
nome_completo,data_nascimento,banco,data_inclusao,email_banco
Cecília Mendes,15/05/2001,Nubank,10/01/2026,exemplo@email.com

👩‍💻 Autora
Cecília - Estudante de Análise e Desenvolvimento de Sistemas (ETEP EAD) Foco em Desenvolvimento Backend e Full-Stack.

Projeto desenvolvido para fins de estudo em Python e Automação.