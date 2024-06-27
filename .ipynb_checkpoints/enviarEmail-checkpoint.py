import email, smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env


def enviar_email():

    subject = "Relatório Notebooks" #titulo do email
    body = '''
    Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.

    Atenciosamente,
    Robô
    '''
    
    sender_email = "email"
    receiver_email = ""
    password = ""


    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  


    message.attach(MIMEText(body, "plain"))

    filename = r"C:\Users\eliez\Desktop\via sul\magazine\output\Notebooks.xlsx"  #caminho do arquivo a ser enviado


    with open(filename, "rb") as attachment:

        part = MIMEBase("application", "vnd.ms-excel") #le o exel
        part.set_payload(attachment.read())


    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment", filename='Notebooks.xlsx', #texto que vai no arquivo
    )


    message.attach(part)
    text = message.as_string()


    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)