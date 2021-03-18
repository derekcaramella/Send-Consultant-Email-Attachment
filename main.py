import pyodbc
import pandas as pd
import pyjokes
from datetime import datetime
import settings
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

con = pyodbc.connect(Trusted_Connection='no',
                     driver='{SQL Server}',
                     server='192.168.15.32',
                     database='Alpha_Live',
                     UID=settings.database_username,
                     PWD=settings.database_password)
cursor = con.cursor()
sn_bagger_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[SN Bagger]', con)
sn_bagger_df.to_excel('Updated SN Data.xlsx', sheet_name='SN Bagger', index=False)


sender_email = 'derekcaramella@gmail.com'
sender_password = settings.email_password
receiver_email = 'tenderby@amt-mep.org'
subject = 'SN Update| ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M')
body = 'Hi Tom, attached is the updated SN data.\n' + str(pyjokes.get_joke()) + '\n\nBest Regards,\nDerek Caramella'


message = MIMEMultipart('alternative')
message['Subject'] = subject
message['From'] = sender_email
message['To'] = receiver_email
message.attach(MIMEText(body, 'plain'))
attachment_file_path = 'Updated SN Data.xlsx'
attachment = open(attachment_file_path, 'rb')
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename= {attachment_file_path}',)
message.attach(part)
text = message.as_string()

port = 465  # For SSL
context = ssl.create_default_context()  # Create a secure SSL context
server = smtplib.SMTP_SSL('smtp.gmail.com', port, context=context)
server.login(sender_email, sender_password)
server.sendmail(sender_email, receiver_email, text)
