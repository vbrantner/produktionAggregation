#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import smtplib
import base64
import configparser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
from datetime import datetime

def sendMail(sendto='v.brantner@brantner.de', subj='', msg=''):
    today = datetime.today()
    config = configparser.ConfigParser()
    config.sections()
    config.read('credentials.ini')
    email = 'v.brantner@brantner.de'
    password = config['O365']['password']
    send_to_email = 'v.brantner@brantner.de'
    subject = 'ProduktionsAnalyse, Monat: {}'.format(today.month)
    message = 'automatisch generierte E-Mail'
    file_location = './produktionsAnalyse.xlsx'

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = send_to_email
    msg['Subject'] = subject
#    msg['CC'] = "j.brantner@brantner.de"

    msg.attach(MIMEText(message, 'plain'))

    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server = smtplib.SMTP('smtp.office365.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo
    server.login(email, password)
    text = msg.as_string()
    server.sendmail(email, send_to_email, text)
    server.quit()

if __name__ == "__main__":
    sendMail()

