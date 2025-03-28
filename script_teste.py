import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import getpass  # Para esconder a senha

# Configurações SMTP do Expresso
smtp_server = "smtps.expresso.pe.gov.br"
smtp_port = 587

# Tentativas de login
tentativas = 3

while tentativas > 0:
    email = input("Digite seu e-mail do Expresso: ")
    senha = getpass.getpass("Digite sua senha do Expresso: ")  # Senha oculta

    try:
        # Conectar ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email, senha)
        print("✅ Login bem-sucedido!\n")
        break  # Sai do loop se o login for bem-sucedido
    except smtplib.SMTPAuthenticationError:
        tentativas -= 1
        print(f"❌ Senha incorreta! Você tem mais {tentativas} tentativa(s).")

        if tentativas == 0:
            print("🚨 Número máximo de tentativas atingido. Encerrando o programa.")
            exit()

# Carregar a planilha
arquivo_excel = "notas_desempenho.xlsx"
df = pd.read_excel(arquivo_excel)

# 📬 Enviar e-mails
for index, row in df.iterrows():
    if pd.isna(row['Status de Envio']):  # Se o e-mail ainda não foi enviado
        # Configurar mensagem
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = row['E-mail']
        msg['Subject'] = "Sua Nota de Desempenho - Ano 2025"

        # Corpo do e-mail
        corpo = f"""
        Olá, {row['Nome']},

        Sua nota de desempenho foi: {row['Nota']}.
        
        Caso tenha dúvidas, entre em contato com a equipe da CTI.

        Atenciosamente,
        ARPE - CTI
        """
        msg.attach(MIMEText(corpo, 'plain'))

        # Solicitar Confirmação de Leitura
        msg.add_header('Disposition-Notification-To', email)  # Confirmação de leitura
        msg.add_header('Return-Receipt-To', email)  # Confirmação de entrega

        # Enviar e-mail
        server.send_message(msg)

        # Atualizar status na planilha
        df.loc[index, 'Status de Envio'] = "Enviado"
        df.loc[index, 'Data de Envio'] = datetime.now().strftime('%d/%m/%Y %H:%M')

        print(f"E-mail enviado para: {row['Nome']} ({row['E-mail']})")

# Salvar planilha atualizada
df.to_excel(arquivo_excel, index=False)
server.quit()
print("\n✅ Todos os e-mails foram enviados com sucesso!")