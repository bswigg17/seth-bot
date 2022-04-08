import csv, os
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import smtplib
from datetime import date


def send_mail():
    # Create a multipart message
    msg = MIMEMultipart()
    body_part = MIMEText("Here's Data", 'plain')
    msg['Subject'] = f"Data for {date.today()}"
    msg['From'] = "bswigg17@gmail.com"
    msg['To'] = "tessa.swier@toshibagcs.com"
    # Add body to email
    msg.attach(body_part)
    # open and read the CSV file in binary
    with open("data.csv",'rb') as file:
    # Attach the file with filename to the email
        msg.attach(MIMEApplication(file.read(), Name="data.csv"))

    try:
        # Create SMTP object
        smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_obj.starttls()
        # Login to the server
        smtp_obj.login('bswigg17', passwordHash())

        # Convert the message to a string and send it
        smtp_obj.sendmail(msg['From'], msg['To'], msg.as_string())
        smtp_obj.quit()
    except Exception as e:
        print(e)
    finally:
        os.remove('data.csv')



def passwordHash():
    password = '0202wB027101'[::-1]
    return password
        
