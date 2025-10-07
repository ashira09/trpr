import os
import smtplib
import imaplib
from csv2pdf import convert
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email import message_from_bytes
from email.utils import make_msgid

# входные данные
password = "faja gqpg uzqg unzi"
message_subject = "Report"
sender_user = "trpr.email.lsa@gmail.com"
receiver_user = "vecnivtil89@gmail.com"
reply_user = "vecnivtil@yandex.ru"
path_to_file = "./lr3/data.csv"
filename = "data.pdf"
smtp_server = "smtp.gmail.com"
imap_server = "imap.gmail.com"
smtp_port = 587

convert(source=path_to_file, destination=f"./lr3/{filename}") # формирование pdf документа из csv файла с данными

# формирование письма
msg = MIMEMultipart()
msg['Subject'] = message_subject
msg['From'] = sender_user
msg['To'] = receiver_user
msg['Reply-To'] = reply_user

# создание и прикрепление файлового приложения к письму
fp = open(f"./lr3/{filename}", "rb")
att = MIMEApplication(fp.read())
fp.close()
att.add_header('Content-Disposition', 'attachment', filename=filename)
msg.attach(att)

# отправка письма
s = smtplib.SMTP(smtp_server, smtp_port)
s.ehlo()
s.starttls()
s.ehlo()
s.login(sender_user, password)
s.sendmail(sender_user, receiver_user, msg.as_string())
s.quit()

mail = imaplib.IMAP4_SSL(imap_server) # подключение к почтовому серверу IMAP
mail.login(sender_user, password) # авторизация в аккаунте почты
mail.select('"[Gmail]/Sent Mail"') # выбор папки, в которой находятся искомые письма
result, data = mail.search(None, f'SUBJECT "{message_subject}"') # поиск писем для перессылки
for msg_id in data[0].split(): # цикл по найденным письмам
    
    # считывание найденного письма
    result, msg_data = mail.fetch(msg_id, "(RFC822)")
    raw_email = msg_data[0][1]
    original = message_from_bytes(raw_email) 
    
    # считывание файлового приложения к письму
    for part in original.walk():
        if part.get_content_type() == "application/octet-stream":
            filename = part.get_filename()
            if filename:
                filepath = os.path.join('./lr3/', 'readed_data.pdf')
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
    
    # формирование письма-перессылки
    original = message_from_bytes(raw_email)
    reply = MIMEMultipart()
    reply["Message-ID"] = make_msgid()
    reply["In_Reply-To"] = original["Reply-To"]
    reply["References"] = original["Message-Id"]
    reply['Subject'] = "Re: " + original["Subject"]
    reply["To"] = original["Reply-To"] or original["From"]
    reply["From"] = sender_user

    # создание и прикрепление считанного выше файлового приложения к письму-перессылки
    fp = open(f"./lr3/{filename}", 'rb')
    att = MIMEApplication(fp.read())
    fp.close()
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    reply.attach(att)

    #отправка письма-перессылки
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(sender_user, password)
    server.sendmail(sender_user, reply["In_Reply-To"], reply.as_string())
    server.quit()