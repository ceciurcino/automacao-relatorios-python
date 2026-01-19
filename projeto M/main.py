import pandas as pd
from fpdf import FPDF
from pypdf import PdfReader, PdfWriter 
import os

# 1. Configurações de Segurança e Pastas
PASTA_SAIDA = "Relatorios_Sigilosos"
if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

# 2. Carregar Dados

try:
    df = pd.read_excel("dados.xlsx")
except FileNotFoundError:
    print("Erro: O arquivo 'dados.xlsx' não foi encontrado.")
    exit()

# 3. Processamento e Geração de PDF
for index, linha in df.iterrows():
    pdf = FPDF()
    pdf.add_page()
    
    # Cabeçalho Formal
    pdf.set_font("Arial", "B", 14)
    pdf.cell(190, 10, "NOTIFICAÇÃO DE REGISTRO - DADOS DO REGISTRATO (BC)", ln=True, align="C")
    pdf.ln(10)
    
    # Dados do Cliente
    pdf.set_font("Arial", "B", 12)
    pdf.cell(190, 10, "DADOS DO TITULAR", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(190, 8, f"Nome Completo: {linha['nome_completo']}", ln=True)
    pdf.cell(190, 8, f"Data de Nascimento: {linha['data_nascimento']}", ln=True)
    pdf.ln(5)
    
    # Detalhes da Dívida
    pdf.set_font("Arial", "B", 12)
    pdf.cell(190, 10, "INFORMAÇÕES DO DÉBITO", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(190, 8, f"Instituição Financeira: {linha['banco']}", ln=True)
    pdf.cell(190, 8, f"Data de Inclusão: {linha['data_inclusao']}", ln=True)
    pdf.cell(190, 8, f"Última Aparição no Sistema: {linha['data_ultima_aparicao']}", ln=True)
    
    # Rodapé de Confidencialidade
    pdf.ln(20)
    pdf.set_font("Arial", "I", 9)
    pdf.multi_cell(190, 5, "Este documento contém dados sensíveis protegidos pela Lei Geral de Proteção de Dados (LGPD). O uso indevido dessas informações está sujeito a penalidades legais.")

    # Nome do arquivo: Nome_Banco.pdf
    nome_arquivo = f"{PASTA_SAIDA}/Relatorio_{linha['nome_completo']}_{linha['banco']}.pdf".replace(" ", "_")
    pdf.output(nome_arquivo)
    
    print(f"✅ Relatório gerado para {linha['nome_completo']} (Banco: {linha['banco']})")