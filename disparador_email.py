# SMTP

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from senha_app_gmail import senha

# load_dotenv()

def enviar_email(assunto, corpo_email, arquivo_anexo, nome_arquivo):
    
    msg = MIMEMultipart()
    msg["Subject"] = assunto
    msg["From"] = "felipe.monsef.vasc@gmail.com"
    msg["To"] = "felipe.monsef.vasc@gmail.com"
    # msg["Cc"] = "seuemailcopia@gmail.com;outroemailcopia@hotmail.com"
    # msg["Bcc"] = "seuemailcopiaoculta@gmail.com"
    

    msg.attach(MIMEText(corpo_email, "html"))

    # anexar arquivos
    with open(arquivo_anexo, "rb") as arquivo:
        msg.attach(MIMEApplication(arquivo.read(), Name=nome_arquivo))

    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(msg["From"], senha)
    servidor.send_message(msg)
    servidor.quit()
    print("Email enviado")


# enviar_email()