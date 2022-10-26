from __future__ import print_function
import collections
import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd
import json
from dotenv import load_dotenv
# Import smtplib for the actual sending function
import smtplib
def line():
    for i in range(30):
        print('=',end="")
    print('=')
# Import the email modules we'll need
def email_send():
    pass


import base64
from email.message import EmailMessage

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
# # Open the plain text file whose name is in textfile for reading.
# with open(textfile) as fp:
#     # Create a text/plain message
#     msg = EmailMessage()
#     msg.set_content(fp.read())

# # me = the sender's email address
# # you = the recipient's email address
# msg['Subject'] = f'The contents of {textfile}'
# msg['From'] = me
# msg['To'] = you

# # Send the message via our own SMTP server.
# s = smtplib.SMTP('localhost')
# s.send_message(msg)
# s.quit()

# Initializing a GoogleAuth Object
# gauth = GoogleAuth()

# # client_secrets.json file is verified
# # and it automatically handles authentication
# gauth.LocalWebserverAuth()

# # GoogleDrive Instance is created using
# # authenticated GoogleAuth instance
# drive = GoogleDrive(gauth)

# # Initialize GoogleDriveFile instance with file id
# file_obj = drive.CreateFile(
#     {'id': '1UCbxwVUZFdHyMQyqlNhPsJe8RJkJTvsa2v8pFXI4wpE'})
# file_obj.GetContentFile('Data Collect (Responses).xls',
#          mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

dataframe = pd.read_excel('Hostel Allocation Grouping (Responses).xlsx')
# print(dataframe)
# line()
# print(type(dataframe['Block No.']))


def match_mate():
    pass


rooms_names_1 = []
rooms_names_2 = []
rooms_names_3 = []
rooms_names_4 = []
room_dict_1 = collections.defaultdict(list) 
room_dict_2 = collections.defaultdict(list) 
room_dict_3 = collections.defaultdict(list) 
room_dict_4 = collections.defaultdict(list) 


for block in dataframe['Block No.'].tolist():

#   print(block)
    match str(block):
        case '2':
            df1 = dataframe.loc[dataframe['Block No.'] == block].head()
            n = 0

            for room in df1['Room No. ']:
                # print(room)
                df2 = df1.loc[df1['Room No. '] == room].head()
                if str(room) in rooms_names_2:
                    for length in range(len(rooms_names_2)):
                        if str(room) == rooms_names_2[length]:
                            if len(room_dict_2[str(room)])==0:
                                room_dict_2[str(room)].append(df2['Name'].tolist())
                                # print(str(room))
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_2.append(str(listdf1[n]))
                    if len(room_dict_2[str(room)])==0:
                                room_dict_2[str(room)].append(df2['Name'].tolist())                    
                    
                n+=1
            
        case '1':
            df1 = dataframe.loc[dataframe['Block No.'] == block].head()
            n = 0
            
            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room].head()
                
                if str(room) in rooms_names_1:
                    
                    for length in range(len(rooms_names_1)):
                        if str(room) == rooms_names_1[length]:
                            if len(room_dict_1[str(room)])==0:
                                room_dict_1[str(room)].append(df2['Name'].tolist())
                                # print(str(room))
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_1.append(str(listdf1[n]))
                    if len(room_dict_1[str(room)])==0:
                                room_dict_1[str(room)].append(df2['Name'].tolist())
                n+=1
        case '3':
            df1 = dataframe.loc[dataframe['Block No.'] == block].head()
            n = 0

            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room].head()
                if str(room) in rooms_names_3:
                    for length in range(len(rooms_names_3)):
                        if str(room) == rooms_names_3[length]:
                            if len(room_dict_3[str(room)])==0:
                                room_dict_3[str(room)].append(df2['Name'].tolist())
                                print(str(room))
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_3.append(str(listdf1[n]))
                    if len(room_dict_3[str(room)])==0:
                                room_dict_3[str(room)].append(df2['Name'].tolist())
                n+=1
        case '4':
            df1 = dataframe.loc[dataframe['Block No.'] == block].head()
            n = 0

            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room].head()
                if str(room) in rooms_names_4:
                    for length in range(len(rooms_names_4)):
                        if str(room) == rooms_names_4[length]:
                            if len(room_dict_4[str(room)])==0:
                                room_dict_4[str(room)].append(df2['Name'].tolist())
                                print(str(room))
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_4.append(str(listdf1[n]))
                    if len(room_dict_4[str(room)])==0:
                                room_dict_4[str(room)].append(df2['Name'].tolist())
                n+=1
        case _:
            print("Bhai Block 5 Under Construction")
            pass

# print(room_dict_1)
# data2 = json.loads(json.dumps(room_dict_2))
# print("room_dict_two: ",data2)
# data1 = json.loads(json.dumps(room_dict_1))
# print("room_dict_one: ",data1)
# # print(rooms_names_1)
# print("room__two:",rooms_names_2)
# print("room__one:",rooms_names_1)

def gmail_send_message():
    """Create and send an email message
    Print the returned  message id
    Returns: Message object, including message id

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"]= "./gmail-cred.json"
    creds, _ = google.auth.default()

    try:
        service = build('gmail', 'v1', credentials=creds)
        message = EmailMessage()

        message.set_content('This is automated draft mail')

        message['To'] = "bikramjeetdg@gmail.com"
        message['From'] = "quilo.bikcodes@gmail.com"
        message['Subject'] = 'Automated draft'

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()) \
            .decode()

        create_message = {
            'raw': encoded_message
        }
     
        # pylint: disable=E1101
        send_message = (service.users().messages().send
                        (userId="me", body=create_message).execute())
        print(F'Message Id: {send_message["id"]}')
    except HttpError as error:
        print(F'An error occurred: {error}')
        send_message = None
    return send_message


SCOPES = ['https://mail.google.com/']

    
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
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()


gmail_send_message()
for email in dataframe['Email Address'].tolist():
    df2 = dataframe.loc[dataframe['Email Address'] == email].head()
    blklist = df2['Block No.'].tolist()
    blockno = ''.join([str(element) for element in blklist]) 
    # print(blockno)
    if blockno == '2':
        roomlist = df2['Room No. '].tolist()

        roomno = ''.join([str(element) for element in roomlist]) 
        # print(room_dict_2[roomno])
    # print(str(email))



        # msg = EmailMessage()
        # content = room_dict_2[roomno][0]
        # msg.set_content("Hello from Bikram")

        # me = "bikramjeetdg@gmail.com"
        # you = "quilo.bikcodes@gmail.com"
        # msg['Subject'] = f'Hey Mate of Room{roomno}. Here are your roommates'
        # msg['From'] = me
        # msg['To'] = you
        # # s = smtplib.SMTP('localhost')

        # load_dotenv(".env")

        # SENDER = os.environ.get("GMAIL_USER")
        # PASSWORD = os.environ.get("GMAIL_PASSWORD")
        # s = smtplib.SMTP_SSL("smtp.gmail.com", 587)
        # s.login(me, "Babi1234")
        # s.send_message(msg)
        # s.quit()
        
# https://docs.google.com/spreadsheets/d/1UCbxwVUZFdHyMQyqlNhPsJe8RJkJTvsa2v8pFXI4wpE/edit?usp=sharing