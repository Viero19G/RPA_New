import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do servidor de e-mail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "notificacao@robustec.com.br"  # Substitua pelo seu e-mail
EMAIL_PASSWORD = "ihhhcgqkixvqbekj"  # Substitua pela sua senha

# Função para enviar e-mail
def enviar_email(destino, assunto, mensagem):
    try:
        # Configuração do e-mail
        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = destino
        msg["Subject"] = assunto
        msg.attach(MIMEText(mensagem, "plain"))

        # Envio do e-mail
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
            print(f"E-mail enviado para {destino}")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destino}: {e}")

# Ler o arquivo Excel
arquivo_excel = r"\\192.168.1.28\robustecfs\PUBLICO\00-NOTIFICA\01-TI\controle.xlsx"  # Substitua pelo caminho do seu arquivo
df = pd.read_excel(arquivo_excel)

# Verificar as datas e enviar e-mails
hoje = datetime.now()
intervalo_7 = timedelta(days=7)
intervalo_15 = timedelta(days=15)
intervalo_30 = timedelta(days=30)
intervalo_60 = timedelta(days=60)
intervalo_90 = timedelta(days=90)

for index, row in df.iterrows():
    try:
        data_vencimento = pd.to_datetime(row["DATA"])
        if  data_vencimento.date() == (hoje + intervalo_30).date():
            destinatario = row["DESTINO"]
            copia = row["COPIA"]
            mensagem = row["MENSAGEM"]
            item = row["ITEM"]
            assunto = f"Alerta 30 dias de vencimento do {item}"
            enviar_email(destinatario, assunto, mensagem)
            enviar_email(copia, assunto, mensagem)
        elif data_vencimento.date() == (hoje + intervalo_60).date():
            destinatario = row["DESTINO"]
            mensagem = row["MENSAGEM"]
            copia = row["COPIA"]
            item = row["ITEM"]
            assunto = f"Alerta 60 dias de vencimento do {item}"
            enviar_email(copia, assunto, mensagem)
            enviar_email(destinatario, assunto, mensagem)
        elif data_vencimento.date() == (hoje + intervalo_7).date():
            destinatario = row["DESTINO"]
            mensagem = row["MENSAGEM"]
            copia = row["COPIA"]
            item = row["ITEM"]
            assunto = f"Alerta 7 dias de vencimento do {item}"
            enviar_email(copia, assunto, mensagem)
            enviar_email(destinatario, assunto, mensagem)
        elif data_vencimento.date() == (hoje + intervalo_15).date():
            destinatario = row["DESTINO"]
            mensagem = row["MENSAGEM"]
            copia = row["COPIA"]
            item = row["ITEM"]
            assunto = f"Alerta 15 dias de vencimento do {item}"
            enviar_email(copia, assunto, mensagem)
            enviar_email(destinatario, assunto, mensagem)
        elif data_vencimento.date() == (hoje + intervalo_90).date():
            destinatario = row["DESTINO"]
            mensagem = row["MENSAGEM"]
            item = row["ITEM"]
            assunto = f"Alerta 90 dias de vencimento do {item}"
            copia = row["COPIA"]
            enviar_email(copia, assunto, mensagem)
            enviar_email(destinatario, assunto, mensagem)
    except Exception as e:
        destino = 'ti3@robustec.com.br'
        assunto = 'ERRO NA LISTA 01-TI'
        mensagem = f'O erro: {e} aconteceu na linha: {row}'
        enviar_email(destino, assunto, mensagem)
