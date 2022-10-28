import pandas as pd
import ssl
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText


global user
global passwd

with open("subject_message.txt", encoding='utf-8') as fil:
    subject = fil.read()

with open("message.txt", encoding='utf-8') as file:
    text = file.read()

df = pd.read_excel("Email_password.xlsx")
user = df['Email'][0]
passwd = df['Password'][0]

server = "smtp.mail.ru"
port = 587

mime = 'MIME-Version: 1.0'



def Send_email(user, to, subject, text):
    em = EmailMessage()
    em['From'] = user
    em['To'] = to
    em['Subject'] = subject
    em.set_content(text)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.yandex.ru', 465, context=context) as smtp:
        smtp.login(user, passwd)
        smtp.sendmail(user, to, em.as_string())


base = pd.read_excel("Pbase.xlsx")
i = 0


for mail in base['Email']:
    name = base['Name'][i]
    if ', ' in base['Email']:
        to = base['Email'].split(",")
        for mails in to:
            if '__NAME__' in text:
                texts = text.replace('__NAME__', name)
                texts = MIMEText(texts, "plain", "utf-8")
            Send_email(user, to, subject, texts)
    if '__NAME__' in text:
        texts = text.replace('__NAME__', name)
    texts = MIMEText(texts, "plain", "utf-8")
    Send_email(user, mail, subject, texts)
    i = i + 1
