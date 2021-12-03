import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

from datetime import timedelta, time
import pickle
import numpy as np
import pandas as pd
import os
import json
import base64
import math

import xlsxwriter
import smtplib
import gspread
import re
import io
import tempfile
import pyodbc
import time

import tempfile

from genericFunctions import googleCloudConnection, list_diff,getDriverField, SQLconnection,PSQLconecction
from functionEncription import decrypt_symmetric, crc32c, decryptCatalogs

from google.cloud import bigquery as bq
from google.oauth2 import service_account
from google.cloud.exceptions import NotFound
from google.cloud import storage

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


dfTrips=final_info.copy()

output = io.BytesIO()

# Use the BytesIO object as the filehandle.
writer = pd.ExcelWriter(output, engine='xlsxwriter')

# Write the data frame to the BytesIO object.
dfTrips.to_excel(writer, sheet_name='informacion clientes looker', startrow=2, header=False, index=False)
workbook  = writer.book
worksheet = writer.sheets['informacion clientes looker']

# Add formats.
header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center', 'fg_color': '#4e0000' , 'color': '#ffffff', 'border': 1})
title_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'font_size': 16, 'fg_color': '#ffffff' , 'color': '#4e0000'})
cell_format = workbook.add_format({'align': 'center', 'border': 1, 'border_color': '#000000'}) 
# Write the column headers with the defined format.
for col_num , value in enumerate(dfTrips.columns.values):
    worksheet.write(1, col_num , value, header_format)

title = 'informacion clientes looker'
worksheet.merge_range('A1:Q1', title, title_format)
worksheet.set_column(0, 16, 25, cell_format)

writer.save()
xlsx_data = output.getvalue()


# Send file
smtpHostName = mailSettings['amazonSMTP']['smtpHostName']
smtpPort = mailSettings['amazonSMTP']['smtpPort']
smtpUserName = mailSettings['amazonSMTP']['smtpUserName']
smtpPassword = mailSettings['amazonSMTP']['smtpPassword']
fromMail = "info@carsync.com"

msg = MIMEMultipart('alternative')
msg['Subject'] = "Informacion clientes looker y dispositivos"
msg['From'] = fromMail
msg['To'] = toMail
msg['Cc'] = ccMail

text = "Informacion clientes looker y dispositivos"
fileName = 'Informacion clientes looker.xlsx'
fileName = fileName.replace(" ", "_")
part1 = MIMEText(text, 'plain')
msg.attach(part1)

p = MIMEBase('application', 'octet-stream') 
p.set_payload(xlsx_data) 
encoders.encode_base64(p) 
p.add_header('Content-Disposition', "attachment; filename= %s" % fileName) 
msg.attach(p) 

# Send via local SMTP server.
s = smtplib.SMTP(smtpHostName,smtpPort)
s.starttls()
s.ehlo()
s.login(smtpUserName,smtpPassword)
s.sendmail(fromMail, toMail.split(',') + ccMail.split(',') + hiddenMail.split(','), msg.as_string())
s.quit()
