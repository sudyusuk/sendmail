import gspread
from oauth2client.service_account import ServiceAccountCredentials
#mail
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


from email.utils import formatdate
import settings

import datetime

FROM_ADDRESS = settings.FROM_ADDRESS
MY_PASSWORD = settings.MY_PASSWORD
TO_ADDRESS = settings.TO_ADDRESS
BCC = ''
file_path = "./test.zip"

def create_message(from_addr, to_addr, bcc_addrs, subject, body):
    # msg = MIMEText(body)
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Bcc'] = bcc_addrs
    msg['Date'] = formatdate()
    msg.attach(MIMEText(body))
    attach = MIMEBase('application', 'zip')
    with open(file_path, "br") as f:
        attach.set_payload(f.read())
    encoders.encode_base64(attach)
    attach.add_header('Content-Disposition', 'attachment',
                      filename='attachment.zip')
    msg.attach(attach)
    return msg

def send(from_addr, to_addrs, msg):
    smtpobj = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=10)
    smtpobj.login(FROM_ADDRESS, MY_PASSWORD)
    smtpobj.sendmail(from_addr, to_addrs, msg.as_string())
    smtpobj.close()

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('api.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open('test_sendmail').sheet1

i=2
while wks.cell(i,11).value != "":
    if wks.cell(i, 18).value =="○":
        pass
    else:
        if wks.cell(i, 14).value !=datetime.date.today().strftime('%Y/%m/%d'):
            pass
        else:
            if wks.cell(i, 15).value !="":
                pass
            else:
                subject = (wks.cell(i, 5).value + "様")
                body = (wks.cell(i, 6).value + "はいかが")
                # subject = SUBJECT
                # body = BODY
                to_addr = wks.cell(i,12).value
                msg = create_message(FROM_ADDRESS, to_addr, BCC, subject, body)
                send(FROM_ADDRESS, to_addr, msg)
                wks.update_cell(i, 15, datetime.datetime.today().strftime('%Y/%m/%d'))
                print(to_addr)
    i+=1

i=2
while wks.cell(i,11).value != "":
    if wks.cell(i, 18).value =="○":
        pass
    else:
        if wks.cell(i, 16).value !=datetime.datetime.today().strftime('%Y/%m/%d'):
            pass
        else:
            if wks.cell(i, 17).value !="":
                pass
            else:
                subject = (wks.cell(i, 5).value + "様")
                body = (wks.cell(i, 6).value + "はいかが")
                # subject = SUBJECT
                # body = BODY
                to_addr = wks.cell(i,12).value
                msg = create_message(FROM_ADDRESS, to_addr, BCC, subject, body)
                send(FROM_ADDRESS, to_addr, msg)
                wks.update_cell(i, 17, datetime.datetime.today().strftime('%Y/%m/%d'))
                print(to_addr)

    i+=1