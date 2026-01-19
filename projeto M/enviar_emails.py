import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --- CONFIGURAÇÕES ---
MEU_EMAIL = "SEU_EMAIL_AQUI" # Coloque seu e-mail aqui
MINHA_SENHA = "SUA_SENHA_DE_APP_AQUI" # Coloque a Senha de App aqui (sem espaços se preferir)
PASTA_ANEXOS = "Relatorios_Sigilosos"

# Carregar dados
df = pd.read_excel("dados.xlsx")

# Conectar ao servidor do Gmail
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()

try:
    server.login(MEU_EMAIL, MINHA_SENHA)
    print("✅ Login realizado com sucesso!")
except Exception as e:
    print(f"❌ Erro no login: {e}")
    exit()

# Loop para enviar para cada cliente/banco
for index, linha in df.iterrows():
    destinatario = linha['email_banco']
    nome_cliente = linha['nome_completo']
    nome_banco = linha['banco']
    
    # 1. Montando o E-mail
    msg = MIMEMultipart()
    msg['From'] = MEU_EMAIL
    msg['To'] = destinatario
    msg['Subject'] = f"NOTIFICAÇÃO EXTRAJUDICIAL - {nome_cliente} - {nome_banco}"
    
    corpo_email = f"""
    Prezados,
    
    Segue em anexo o relatório referente aos dados do cliente {nome_cliente} vinculados à instituição {nome_banco}.
    
    Atenciosamente,
    Sistema Automatizado.
    """
    msg.attach(MIMEText(corpo_email, 'plain'))
    
    # 2. Anexando o PDF
    nome_arquivo_pdf = f"Relatorio_{nome_cliente}_{nome_banco}.pdf".replace(" ", "_")
    caminho_completo = os.path.join(PASTA_ANEXOS, nome_arquivo_pdf)
    
    try:
        attachment = open(caminho_completo, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {nome_arquivo_pdf}")
        msg.attach(part)
        attachment.close()
        
        # 3. Enviando
        server.sendmail(MEU_EMAIL, destinatario, msg.as_string())
        print(f"📧 E-mail enviado para {destinatario} (Cliente: {nome_cliente})")
        
    except FileNotFoundError:
        print(f"⚠️ Arquivo não encontrado: {nome_arquivo_pdf}")
    except Exception as e:
        print(f"❌ Erro ao enviar para {destinatario}: {e}")

# Fechar conexão
server.quit()
print("--- Fim do Processo ---")