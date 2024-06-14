import json
import pandas as pd
import  requests
from requests.auth import HTTPBasicAuth
from sqlalchemy import create_engine
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime as dt
from datetime import datetime, timedelta
from openpyxl.styles import Alignment


def send_mail(send_from, send_to, subject, text, server, port, username='', password='', filename=None):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    if filename is not None:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(filename, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

    smtp = smtplib.SMTP_SSL(server, port)
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

start_time = datetime.utcnow() - timedelta(hours=2)

end_time = datetime.utcnow()- timedelta(hours=1)

# Extract only the date and hour from the current system datetime
start_datetime = start_time.strftime('%Y-%m-%dT%H:00:00.000Z')

end_datetime = end_time.strftime('%Y-%m-%dT%H:00:00.000Z')

cust_url = 'https://adminwebapi.iqsoftllc.com/api/Main/ApiRequest?TimeZone=0&LanguageId=en'

cust_data = {
    "Controller": "Client",
    "Method": "GetClients",
    "RequestObject": {
        "Controller": "Client",
        "Method": "GetClients",
        "SkipCount": 0,
        "TakeCount": 9999,
        "OrderBy": None,
        "FieldNameToOrderBy": "",
        "CreatedFrom": start_datetime,
        "CreatedBefore": end_datetime
    },
    "UserId": "1780",
    "ApiKey": "betfoxx_api_key"
}

cust_response = requests.post(cust_url, json=cust_data)
cust_response_data = cust_response.json()
cust_entities = cust_response_data['ResponseObject']['Entities']
customers = pd.DataFrame(cust_entities)

if customers is not None and customers.shape[0] > 0:
    customers_1 = customers[['Id', 'Email', 'FirstName','LastName','MobileNumber','CountryName','AffiliateId','LastDepositDate','CreationTime']]
    customers_2 = customers_1[customers['LastDepositDate'].isnull()]
    customers_2["CreationTime"] = pd.to_datetime(customers_2["CreationTime"])
    customers_2 = customers_2.rename(columns={"CreationTime": "CreationTime_utc"})   

    filename = f'RND_{end_datetime}.xlsx'

    valid_filename = filename.replace(':', '-').replace('T', '_')
    
    with pd.ExcelWriter(valid_filename, engine='openpyxl') as writer:
        customers_2.reset_index(drop=True).to_excel(writer, sheet_name="RND_Customers", index=False)

    sub = f'Registered and Non Depositors {end_datetime}'
    
    subject = sub
    body = f"Hi,\n\n Attached contains the during the hour of  {end_datetime} (UTC) for Betfoxx \n\nThanks,\nSaketh"
    sender = "sakethg250@gmail.com"
    recipients = ["saketh@crystalwg.com","sebastian@crystalwg.com","SANDRA@CRYSTALWG.COM","ron@crystalwg.com","camila@crystalwg.com","celeste@crystalwg.com","cristina@crystalwg.com","lina.betcoco@gmail.com","erika@crystalwg.com"]
    password = "xjyb jsdl buri ylqr"
    send_mail(sender, recipients, subject, body, "smtp.gmail.com", 465, sender, password, valid_filename)
    
else:
    subject = f'Registered and Non Depositors {end_datetime}'
    body = "Hi,\n\nNo RND customers were found during the specified period.\n\nThanks,\nSaketh"
    sender = "sakethg250@gmail.com"
    recipients = ["saketh@crystalwg.com","sebastian@crystalwg.com","SANDRA@CRYSTALWG.COM","ron@crystalwg.com","camila@crystalwg.com","celeste@crystalwg.com","cristina@crystalwg.com","lina.betcoco@gmail.com","erika@crystalwg.com"]
    password = "xjyb jsdl buri ylqr"

    send_mail(sender, recipients, subject, body, "smtp.gmail.com", 465, sender, password)