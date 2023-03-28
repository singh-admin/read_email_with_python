#-------------   OUTLOOK IMPLEMENTATION ---------------#


import imaplib
import email
from bs4 import BeautifulSoup
import email
from email.header import decode_header
import webbrowser
import os

# account credentials
username = "mrigendrasing@outlook.com"
password = "Mrigendra@12"

imap_server = "outlook.office365.com"

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# create an IMAP4 class with SSL 
imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
imap.login(username, password)
status, messages = imap.select("INBOX")
# number of top emails to fetch
# status, messages = imap.list("INBOX")
N = 1
# total number of emails
messages = int(messages[0])
print(messages)


for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode(encoding)
            # decode email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)
                print(From, "Sent:")
            print("Subject:", subject)
            print("From:", From)
            # if the email message is multipart
            if msg.is_multipart():
                print("inside", "is_multipart")
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    # print(content_type, "content_type")
                    content_disposition = str(part.get("Content-Disposition"))
                    # print(content_disposition, "content_disposition")
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                        soup = BeautifulSoup(body, 'html.parser')

                        # Extract the "From" address from the email using the email module
                        msg = email.message_from_string(soup.prettify())
                        from_address = msg['Sent']

                        # Print the "From" address
                        print(from_address, "Sent")
                        # print(body, "body")
                    except Exception:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body)
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            folder_name = clean(subject)
                            if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                os.mkdir(folder_name)
                            filepath = os.path.join(folder_name, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                # webbrowser.open(filepath)
            # print("="*100)
# close the connection and logout
imap.close()
imap.logout()





#-------------   GMAIL IMPLEMENTATION ---------------#


import os
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'mrigendrasingh434@gmail.com'


def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            print(vars(flow))
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()


# Adds the attachment with the given filename to the given message
def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(filename, 'rb')
        msg = MIMEText(fp.read().decode(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(filename, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(filename, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(filename, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)



def build_message(destination, obj, body, attachments=[]):
    if not attachments: # no attachments given
        message = MIMEText(body)
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
    else:
        message = MIMEMultipart()
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message.attach(MIMEText(body))
        for filename in attachments:
            add_attachment(message, filename)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}


def send_message(service, destination, obj, body):
    return service.users().messages().send(
      userId="me",
      body=build_message(destination, obj, body)
    ).execute()


# test send email
send_message(service, "singhvictor300@gmail.com", "For Testing", 
            "Hi This is mrigendra just using for Testing.")



def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


def read_message(service, message):
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Prints email basic information (To, From, Subject & Date) and plain/text parts
        - Creates a folder for each email based on the subject
        - Downloads text/html content (if available) and saves it under the folder created as index.html
        - Downloads any file that is attached to the email and saves it in the folder created
    """
    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    folder_name = "email"
    has_subject = False
    if headers:
        # this section prints email basic info & creates a folder for the email
        for header in headers:
            name = header.get("name")
            value = header.get("value")
            if name.lower() == 'from':
                # we print the From address
                print("From:", value)
            if name.lower() == "to":
                # we print the To address
                print("To:", value)
            # if name.lower() == "subject":
                # make our boolean True, the email has "subject"
                # has_subject = True
                # make a directory with the name of the subject
                # folder_name = clean(value)
                # we will also handle emails with the same subject name
                # folder_counter = 0
                # while os.path.isdir(folder_name):
                #     folder_counter += 1
                #     # we have the same folder name, add a number next to it
                #     if folder_name[-1].isdigit() and folder_name[-2] == "_":
                #         folder_name = f"{folder_name[:-2]}_{folder_counter}"
                #     elif folder_name[-2:].isdigit() and folder_name[-3] == "_":
                #         folder_name = f"{folder_name[:-3]}_{folder_counter}"
                #     else:
                #         folder_name = f"{folder_name}_{folder_counter}"
                # os.mkdir(folder_name)
                # print("Subject:", value)
            # if name.lower() == "date":
            #     # we print the date when the message was sent
            #     print("Date:", value)
    # if not has_subject:
    #     # if the email does not have a subject, then make a folder with "email" name
    #     # since folders are created based on subjects
    #     if not os.path.isdir(folder_name):
    #         os.mkdir(folder_name)
    # parse_parts(service, parts, folder_name, message)
    # print("="*50)



# get emails that match the query you specify
results = search_messages(service, "Python Code")
print(f"Found {len(results)} results.")
# for each email matched, read it (output plain/text to console & save HTML and attachments)
for msg in results:
    read_message(service, msg)