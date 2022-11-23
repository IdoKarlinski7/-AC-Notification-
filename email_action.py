import os
import pickle
import datetime
import google.auth.exceptions as ex
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from base64 import urlsafe_b64decode, urlsafe_b64encode
import email.encoders as encoder
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type


EMAIL = 'shared@fusmobile.com'
REAL_CONTACTS = ['ron@fusmobile.com', 'arik@fusmobile.com', 'brians@fusmobile.com', 'niv@fusmobile.com',
                 'ericm@fusmobile.com', 'stephen@fusmobile.com', 'ido@fusmobile.com']
SUPERVISOR_EMAIL = 'ido@fusmobile.com'

SCOPES = ['https://mail.google.com/']


def gmail_authenticate():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        os.remove("token.pickle")
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open("token.pickle", "wb") as token:
        pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


def build_message(destination, title, body, attachments=[]):
    if not attachments:
        message = MIMEText(body)
    else:
        message = MIMEMultipart()
        message.attach(MIMEText(body))
        for filename in attachments:
            add_attachment(message, filename)
    message['to'] = destination
    message['from'] = EMAIL
    message['subject'] = title
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}


def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    fp = open(filename, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()
    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', "attachment", filename=filename)
    encoder.encode_base64(msg)
    message.attach(msg)


def send_message(service, destination, obj, body, attachments=[]):
    msg = build_message(destination, obj, body, attachments)
    return service.users().messages().send(userId="me", body=msg).execute()


def send_info_msg(msg):
    service = gmail_authenticate()
    send_message(service, SUPERVISOR_EMAIL, 'Recalibration update for all active systems', msg)


def send_notification(sys_name, due_date, location=None, attachments=[], error_msg=None):
    subject = str(sys_name) + ' acoustic recalibration'
    due_date_format = str(due_date).split()[0]
    countdown = due_date - datetime.datetime.today() + datetime.timedelta(days=1)
    countdown_days = str(countdown).split(',')[0]

    if error_msg:  # error handling
        message = 'In ' + due_date_format + ' System ' + str(sys_name) + \
                  ' was not able to complete script, failed to - ' + error_msg

    elif countdown_days[0] == '-':  # recalibration date passed
        message = 'System calibration date has passed and a new QF-00005 was not uploaded.' \
              ' Please update Arena as soon as possible.'
    else:
        message = 'System ' + sys_name + ' located in ' + location + ' is due to recalibration at ' + \
                  due_date_format + ', in ' + countdown_days

    service = gmail_authenticate()
    '''for contact in REAL_CONTACTS:
        send_message(service, contact, subject, message, attachments)'''
   # send_message(service, SUPERVISOR_EMAIL, 'test', 'test msg')
