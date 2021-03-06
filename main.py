# Import necessary
import pyodbc
import os
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

os.chdir(r'C:\Users\carmelld\OneDrive - Yildiz Holding\Documents\Send Consultant Email Attachment')  # Switch directory
# Connect to database
con = pyodbc.connect(Trusted_Connection='no',
                     driver='{SQL Server}',
                     server=settings.database_ip,
                     database='Alpha_Live',
                     UID=settings.database_id,
                     PWD=settings.database_password)
cursor = con.cursor()  # Create database cursor
chocotech_kitchen_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[Chocotech Kitchen]', con)  # Kitchen data query
kgm_080_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[KGM 080]', con)  # Run KGM 080 data query
ttb_015_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[TTB 015]', con)  # Run TTB 015 data query
m5_090_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[M5 090]', con)  # Run M5 090 data query
tt_100_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[TT 100]', con)  # Run TT 100 data query
m5_140_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[M5 140]', con)  # Run M5 140 data query
tt_150_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[TT 150]', con)  # Run TT 150 data query
dbs_080_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[DBS 080]', con)  # Run DBS 080 data query
dfr_031_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[DFR 031]', con)  # Run DFR 031 data query
hcm_273_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[HCM 273]', con)  # Run HCM 273 data query
hcm_274_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[HCM 274]', con)  # Run HCM 274 data query
ttm_147_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[TTM 147]', con)  # Run TTM 147 data query
sn_bagger_df = pd.read_sql('SELECT * FROM [Alpha_Live].[dbo].[SN Bagger]', con)  # Run SN data query
# Goal: Loop through dataframes to compile into Excel sheets.
dataframes_dic = {'Kitchen': chocotech_kitchen_df, 'KGM 080': kgm_080_df, 'TTB 015': ttb_015_df, 'M5 090': m5_090_df,
                  'TT 100': tt_100_df, 'M5 140': m5_140_df, 'TT 150': tt_150_df, 'DBS 080': dbs_080_df,
                  'DFR 031': dfr_031_df, 'HCM 273': hcm_273_df, 'HCM 274': hcm_274_df, 'TTM 147': ttm_147_df,
                  'SN': sn_bagger_df}  # Dictionary that enable sheet naming
excel_writer = pd.ExcelWriter('Updated Site Data.xlsx', engine='xlsxwriter')  # Creates Excel Writer
for dataframe in dataframes_dic:  # dataframe is the string, use the dictionary to obtain the dataframe object
    dataframes_dic[dataframe].to_excel(excel_writer, sheet_name=dataframe, index=False)
excel_writer.save()  # Save the Excel workbook

sender_email = 'derekcaramella@gmail.com'  # Sender's email address
sender_password = settings.email_password  # Sender's email password
receiver_email = 'tenderby@amt-mep.org'  # Receiver's email
subject = 'BF Site Update| ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M')  # Email subject line
# Body of email
body = 'Hi Tom, attached is the updated SN data.\n' + str(pyjokes.get_joke()) + '\n\nBest Regards,\nDerek Caramella'

# Goal: Send email with attachment
message = MIMEMultipart('alternative')  # Sets the email to include
message['Subject'] = subject  # Add predefined subject line
message['From'] = sender_email  # Add predefined sender's address
message['To'] = receiver_email  # # Add predefined receiver's address
message.attach(MIMEText(body, 'plain'))  # Add predefined body as plain text
attachment_file_path = 'Updated SN Data.xlsx'  # Saved Excel file path from previous SQL query
attachment = open(attachment_file_path, 'rb')  # Open the attachment with reading binary parameter
part = MIMEBase('application', 'octet-stream')  # Circle back later
part.set_payload(attachment.read())  # Circle back later
encoders.encode_base64(part)  # Circle back later
part.add_header('Content-Disposition', f'attachment; filename= {attachment_file_path}', )  # Add the name to attachment
message.attach(part)  # Attach the document
text = message.as_string()  # Aggregates the email into a package for sending

port = 465  # For SSL
context = ssl.create_default_context()  # Create a secure SSL context
server = smtplib.SMTP_SSL('smtp.gmail.com', port, context=context)  # Circle back later
server.login(sender_email, sender_password)  # Login into email
server.sendmail(sender_email, receiver_email, text)  # Send the message to the receiver from sender's email
