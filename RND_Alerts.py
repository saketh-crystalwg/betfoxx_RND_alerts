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
from datetime import datetime, timedelta, timezone
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

start_time = datetime.now(timezone.utc) - timedelta(minutes=60)

end_time = datetime.now(timezone.utc) - timedelta(minutes=5)

end_time_1 = datetime.now(timezone.utc)

# Extract only the date and hour from the current system datetime
start_datetime = start_time.strftime('%Y-%m-%dT%H:00:00.000Z')

end_datetime = end_time.strftime('%Y-%m-%dT%H:00:00.000Z')

end_datetime_1 = end_time_1.strftime('%Y-%m-%dT%H:00:00.000Z')

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

txn_url = 'https://adminwebapi.iqsoftllc.com/api/Main/ApiRequest?TimeZone=0&LanguageId=en'

txn_data = {"Controller":"PaymentSystem",
            "Method":"GetPaymentRequestsPaging",
            "RequestObject":{
                "Controller":"PaymentSystem",
                "Method":"GetPaymentRequestsPaging",
                "SkipCount":0,
                "TakeCount":1000,
                "OrderBy":None,
                "FieldNameToOrderBy":"",
                "Type":2,
                "HasNote":False,
                "FromDate":start_datetime,"ToDate":end_datetime_1},
            "UserId":"1780","ApiKey":"betfoxx_api_key"}

txn_response = requests.post(txn_url, json=txn_data)

txn_response_data = txn_response.json()

txn_entities = txn_response_data['ResponseObject']['PaymentRequests']['Entities']

txns = pd.DataFrame(txn_entities)

end_datetime_1 = end_time.strftime('%Y-%m-%dT%H:%M:%S.000Z')

if txns is not None and txns.shape[0] > 0:
    txns['Status'] = ['Approved' if x == 8 \
                      else 'ApprovedManually' if x == 12 \
                      else 'Cancelled' if x == 2 \
                      else 'CancelPending' if x == 14 \
                      else 'Confirmed' if x == 7 \
                      else 'Declined' if x == 6 \
                      else 'Deleted' if x == 11 \
                      else 'Expired' if x == 13 \
                      else 'Failed' if x == 9 \
                      else 'Frozen' if x == 4 \
                      else 'InProcess' if x == 3 \
                      else 'Pay Pending' if x == 10 \
                      else 'Pending' if x == 1 \
                      else 'Splitted' if x == 15 \
                      else 'Waiting For KYC' if x == 5 \
                      else 'NA' for x in txns['State']]
    txns["CreationTime"] = pd.to_datetime(txns["CreationTime"])
    txns = txns.rename(columns={"CreationTime": "CreationTime_utc"})

if customers is not None and customers.shape[0] > 0:
    customers_1 = customers[['Id', 'Email', 'FirstName','LastName','MobileNumber','CountryName','AffiliateId','LastDepositDate','CreationTime']]
    customers_2 = customers_1[customers['LastDepositDate'].isnull()]
    customers_2["CreationTime"] = pd.to_datetime(customers_2["CreationTime"])
    customers_2 = customers_2.rename(columns={"CreationTime": "CreationTime_utc"})
    
    if customers_2  is not None and customers_2.shape[0] > 0 and txns is not None and txns.shape[0] > 0:
        customers_2['_key'] = 1
        txns['_key'] = 1
        cross_joined = pd.merge(customers_2, txns, on='_key').drop('_key', axis=1)
        cross_joined_filtered = cross_joined[(cross_joined['Id_x'] == cross_joined['ClientId']) & (cross_joined['CreationTime_utc_y'] > cross_joined['CreationTime_utc_x'])]
        filtered_client_ids = cross_joined_filtered['ClientId'].unique()
        customers_2_filtered = customers_2[~customers_2['Id'].isin(filtered_client_ids)]
        customers_2_filtered.drop('_key', axis=1,inplace = True)
        
    filename = f'RND_{end_datetime}.xlsx'

    valid_filename = filename.replace(':', '-').replace('T', '_')
    
    with pd.ExcelWriter(valid_filename, engine='openpyxl') as writer:
        customers_2_filtered.reset_index(drop=True).to_excel(writer, sheet_name="RND_Customers", index=False)

    sub = f'Registered and Non Depositors {end_datetime}'
    
    subject = sub
    body = f"Hi,\n\n Attached contains the during the hour of  {end_datetime} (UTC) for Betfoxx \n\nThanks,\nSaketh"
    sender = "sakethg250@gmail.com"
    recipients = ["saketh@crystalwg.com","ron@crystalwg.com","camila@crystalwg.com","celeste@crystalwg.com","cristina@crystalwg.com","lina@crystalwg.com","erika@crystalwg.com","isaac@crystalwg.com",
    "sakethg250@gmail.com","alberto@crystalwg.com", "ximena@crystalwg.com","camila.betcoco@gmail.com"]
    password = "xjyb jsdl buri ylqr"
    send_mail(sender, recipients, subject, body, "smtp.gmail.com", 465, sender, password, valid_filename)
    
else:
    subject = f'Registered and Non Depositors {end_datetime}'
    body = "Hi,\n\nNo RND customers were found during the specified period.\n\nThanks,\nSaketh"
    sender = "sakethg250@gmail.com"
    recipients = ["sakethg250@gmail.com"]
    password = "xjyb jsdl buri ylqr"

    send_mail(sender, recipients, subject, body, "smtp.gmail.com", 465, sender, password)