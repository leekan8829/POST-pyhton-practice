import os
import sys
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def sendmail(file_subject,file_path1):
    COMMASPACE = ','
    sender = '寄件人帳號'
    gmail_password = '寄件人密碼'
    recipients = ['收件人01的EMAIL','收件人02的EMAIL']
    # 建立郵件主題
    outer = MIMEMultipart()
    outer['Subject'] = file_subject
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'You will not see this in a MIME-aware mail reader.\n'

    # 檔案位置 在windows底下記得要加上r 如下 要完整的路徑
    attachments = [file_path1]

    # 內容文字部分
    index_text = "Hi this is test"

    # 處理我們的文字 MIMEtext
    mine_text = MIMEText(_text=index_text, _subtype="plain", _charset="UTF8")
    outer.attach(mine_text)
    # 加入檔案到MAIL底下
    for file in attachments:
        try:
            with open(file, 'rb') as fp:
                print ('can read faile')
                msg = MIMEBase('application', "octet-stream")
                msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            outer.attach(msg)
        except:
            print("Unable to open one of the attachments. Error: ", sys.exc_info()[0])
            raise

    composed = outer.as_string()

    # 寄送EMAIL
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, gmail_password)
            s.sendmail(sender, recipients, composed)
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ", sys.exc_info()[0])
        raise