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

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


# Writes gmail message
def create_message(sender, content, subject):

    msg = EmailMessage()
   
    # msg.set_content("Hello from Bikram")
    msg.set_content(content)

    me = "bikramjeetdasgupta@gmail.com"
    you = sender
    # msg['Subject'] = f'Hey Mate of Room. Here are your roommates'
    msg['Subject'] = subject
    msg['From'] = me
    msg['To'] = you
    return {'raw': base64.urlsafe_b64encode(msg.as_string().encode()).decode()}


# Sends gmail message
def send_message(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except HttpError as error:
        print('An error occurred: %s' % error)


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    msg = create_message()
    send_message(service, 'me', msg)



for email in dataframe['Email Address'].tolist():

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('gmail', 'v1', credentials=creds)

   
    

    

    df2 = dataframe.loc[dataframe['Email Address'] == email].head()
    blklist = df2['Block No.'].tolist()
    blockno = ''.join([str(element) for element in blklist]) 
    # print(blockno)
    if blockno == '2':
        roomlist = df2['Room No. '].tolist()
        
        roomno = ''.join([str(element) for element in roomlist]) 
        # print(room_dict_2[roomno])
    # print(str(email))

        sender = "quilo.bikcodes@gmail.com"
        context = "So these are your roommates:"
        for i in room_dict_2[roomno][0]:
            context = context + "\n" + i
        context = context + "\n\nThanks for filling the G Forms\n\n\n\nFrom the operator \nI am Bikramjeet Dasgupta \nCSE \nAI ML \nFresher\n"
        print(context)
        subject = f'Hey Mate of Room {roomno}. Here are your roommates'
        msg = create_message(sender, context , subject )
        # send_message(service, 'me', msg)        
        break
        # msg = EmailMessage()
        # content = room_dict_2[roomno][0]
        # msg.set_content("Hello from Bikram")

        # me = "bikramjeetdasgupta@yahoo.com"
        # you = "quilo.bikcodes@gmail.com"
        # msg['Subject'] = f'Hey Mate of Room{roomno}. Here are your roommates'
        # msg['From'] = me
        # msg['To'] = you

        # # s = smtplib.SMTP('smtp.zoho.com', 465)

        # load_dotenv(".env")

        # SENDER = os.environ.get("GMAIL_USER")
        # PASSWORD = os.environ.get("GMAIL_PASSWORD")
        # # s.ehlo()
        # # s.starttls()
        # # s.ehlo()
        # # s = smtplib.SMTP_SSL("smtp.gmail.com", 587)
        # session = smtplib.SMTP_SSL('smtp.zoho.com', 465)
        # # session.starttls()
        # session.login("bikramjeetdg@zohomail.in", "Babi1234")

        # session.sendmail(me, you, msg)
        # session.close()
        # # s.login(SENDER,PASSWORD)
        # # s.sendmail(SENDER, you,msg)
        # # # s.send_message(msg)
        # # s.quit()
        
# https://docs.google.com/spreadsheets/d/1UCbxwVUZFdHyMQyqlNhPsJe8RJkJTvsa2v8pFXI4wpE/edit?usp=sharing
# G6Mu*QG6FzYd9*a