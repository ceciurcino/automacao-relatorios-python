import pandas as pd
from fpdf import FPDF
from pypdf import PdfReader, PdfWriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --- 1. CONFIGURAÇÕES GERAIS ---
MEU_EMAIL = "SEU_EMAIL_AQUI"  # <--- COLOQUE SEU EMAIL AQUI
MINHA_SENHA = "SUA_SENHA_DE_APP_AQUI" # <--- COLOQUE SUA SENHA DE APP AQUI
PASTA_SAIDA = "Relatorios_Finais"

# Cria a pasta se não existir
if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

# --- 2. CARREGAR BANCO DE DADOS ---
try:
    print("📂 Lendo arquivo de dados...")
    df = pd.read_excel("dados.xlsx")
    # Garante que as datas sejam texto para usar como senha
    df['data_nascimento'] = df['data_nascimento'].astype(str)
except FileNotFoundError:
    print("❌ Erro: Arquivo 'dados.xlsx' não encontrado.")
    exit()

# --- 3. CONECTAR AO SERVIDOR DE EMAIL (GMAIL) ---
try:
    print("🔌 Conectando ao servidor de e-mail...")
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(MEU_EMAIL, MINHA_SENHA)
    print("✅ Conexão estabelecida!")
except Exception as e:
    print(f"❌ Falha na conexão com e-mail: {e}")
    exit()

# --- 4. CICLO DE AUTOMAÇÃO (Para cada cliente) ---
print(f"🚀 Iniciando processamento de {len(df)} registros...\n")

for index, linha in df.iterrows():
    nome = linha['nome_completo']
    banco = linha['banco']
    email_destinatario = linha['email_banco']
    
    # ---------------------------------------
    # PASSO A: GERAR O PDF VISUAL
    # ---------------------------------------
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_font("Arial", "B", 14)
    pdf.cell(190, 10, "NOTIFICAÇÃO EXTRAJUDICIAL", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("Arial", size=12)
    pdf.cell(190, 10, f"Devedor: {nome}", ln=True)
    pdf.cell(190, 10, f"Instituição: {banco}", ln=True)
    pdf.cell(190, 10, f"Data da Inclusão: {linha['data_inclusao']}", ln=True)
    pdf.cell(190, 10, f"Valor Pendente (Ref.): {linha['data_ultima_aparicao']}", ln=True)
    
    pdf.ln(20)
    pdf.set_font("Arial", "I", 10)
    pdf.multi_cell(190, 10, "Documento protegido. Senha: Data de Nascimento (DDMMAAAA).")

    # Nome do arquivo temporário
    nome_arquivo = f"{PASTA_SAIDA}/{nome}_{banco}.pdf".replace(" ", "_")
    pdf.output(nome_arquivo)

    # ---------------------------------------
    # PASSO B: CRIPTOGRAFAR (COLOCAR SENHA)
    # ---------------------------------------
    # Limpa a data para virar senha (ex: 15/05/2001 -> 15052001)
    senha = linha['data_nascimento'].replace("/", "").replace("-", "").replace(" ", "")
    
    reader = PdfReader(nome_arquivo)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(senha)

    # Salva o arquivo protegido por cima do original
    with open(nome_arquivo, "wb") as f_out:
        writer.write(f_out)

    # ---------------------------------------
    # PASSO C: ENVIAR O E-MAIL
    # ---------------------------------------
    msg = MIMEMultipart()
    msg['From'] = MEU_EMAIL
    msg['To'] = email_destinatario
    msg['Subject'] = f"Relatório Confidencial - {nome} - {banco}"

    corpo = f"""
    Prezados,
    
    Segue anexo o relatório seguro referente ao cliente {nome}.
    O arquivo está protegido por senha (Data de Nascimento do titular).
    
    Atenciosamente,
    Robô de Automação Python.
    """
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar o PDF
    try:
        with open(nome_arquivo, "rb") as anexo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(anexo.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(nome_arquivo)}")
            msg.attach(part)

        # Enviar
        server.sendmail(MEU_EMAIL, email_destinatario, msg.as_string())
        print(f"✅ [SUCESSO] {nome} -> Enviado para {email_destinatario}")
        
    except Exception as e:
        print(f"⚠️ [ERRO] Falha ao enviar para {nome}: {e}")

# --- 5. FINALIZAÇÃO ---
server.quit()
print("\n🏁 Processo finalizado com sucesso!")