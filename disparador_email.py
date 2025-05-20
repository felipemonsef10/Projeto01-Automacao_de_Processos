# SMTP

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from senha_app_gmail import senha

# load_dotenv()

def enviar_email(from_email, to_email, assunto, corpo_email, arquivo_anexo, nome_arquivo, qtd_anexos=1):
    
    msg = MIMEMultipart()
    msg["Subject"] = assunto
    msg["From"] = from_email
    msg["To"] = to_email
    # msg["Cc"] = "seuemailcopia@gmail.com;outroemailcopia@hotmail.com"
    # msg["Bcc"] = "seuemailcopiaoculta@gmail.com"
    

    msg.attach(MIMEText(corpo_email, "html"))

    if qtd_anexos == 1:
        # anexar arquivos
        with open(arquivo_anexo, "rb") as arquivo:
            msg.attach(MIMEApplication(arquivo.read(), Name=nome_arquivo))
    else:
        for i, arquivo in enumerate(arquivo_anexo):
            # anexar arquivos
            with open(arquivo, "rb") as arquivo:
                msg.attach(MIMEApplication(arquivo.read(), Name=nome_arquivo[i]))


    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(msg["From"], senha)
    servidor.send_message(msg)
    servidor.quit()
    print("Email enviado")


# enviar_email()