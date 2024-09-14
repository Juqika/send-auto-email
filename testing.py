import pandas as pd
import smtplib
import re
import time
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

def check_email(email):
    regex = r'^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w+$'
    return re.search(regex, email) is not None
    
def get_name(email):
    name_strip = email.split('@')[0]
    name = re.sub(r'\d+', '', name_strip)
    return name.capitalize()
    
def send_email(to_address, subject, body, from_address, smtp_server, smtp_port, login, password):

    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(login, password)
        server.sendmail(from_address, to_address, msg.as_string())
        server.quit()
        return True
    except smtplib.SMTPAuthenticationError:
        print(f"gagal login -_-, check password atau email.")
    except smtplib.SMTPRecipientsRefused:
        print(f"Noooo email gagal terkirim, email ({to_address}) tidak valid.")
    except smtplib.SMTPException as e:
        print(f"Error dalam kirim email ke {to_address}. Error: {str(e)}")
    return False

def sendEmail_to_excel(file_path, from_address, smtp_server, smtp_port, login, password, subject, delay=2):

    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Tidak dapat membaca file Excel. Error: {str(e)}")
        return
    

    for index, row in data.iterrows():
        to_address = row.get('Email')
        # subject = row.get('Subject')
        # message = row.get('Message')

        if pd.isna(to_address):
            print(f"No {index + 1} terlewat karna data kosong.")
            continue

        if not check_email(to_address):
            print(f"Alamat email tidak valid pada No {index + 1}: {to_address}.")
            continue

        name = get_name(to_address)
        body = f"""
what up {name} mate!!


"""

        if send_email(to_address, subject, body, from_address, smtp_server, smtp_port, login, password):
            print(f"yaattaaa... send to {to_address}, No: {index + 1}.")
        else:
            print(f"!!!Gagal terkirim!!! ke {to_address}, No: {index + 1}----!!!")

        time.sleep(delay)


file_path = 'D:\coding\email\data.xlsx'
from_address = os.getenv("EMAIL_USER")
smtp_server = 'smtp.gmail.com'
smtp_port = 587
login = os.getenv("EMAIL_USER")
password = os.getenv("PASSWORD_USER")
send_time = 2
# name = get_name()
subject = "----___TESTING___----"
# body_t = f"""
# what up {name} a

# """


sendEmail_to_excel(file_path, from_address, smtp_server, smtp_port, login, password, subject, delay=send_time)