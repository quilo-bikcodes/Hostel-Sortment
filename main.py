from __future__ import print_function
import collections
import os
import pickle
from time import sleep
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
    for i in range(100):
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
block_1_dict = collections.defaultdict(list) 
block_2_dict = collections.defaultdict(list) 
block_3_dict = collections.defaultdict(list) 
neighbors_dict_2 = collections.defaultdict(list)
# print(dataframe['Block No.'].tolist())
# print(dataframe['Room No. '].tolist())


for block in dataframe['Block No.'].tolist():
    sleep(0.001)

    
    match str(block):
        
        case '2' :
            df1 = dataframe.loc[dataframe['Block No.'] == block]
            
            n = 0
            block_2_dict[str(block)].append(df1['Name'].tolist())
            
            for room in df1['Room No. ']:
                # print(room)
                df2 = df1.loc[df1['Room No. '] == room]
                if str(room) in rooms_names_2:
                    for length in range(len(rooms_names_2)):
                        if str(room) == rooms_names_2[length]:
                            if len(room_dict_2[str(room)])==0:
                                room_dict_2[str(room)].append(df2['Name'].tolist())
                                room_dict_2[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_2[str(room)].append(df2['HomeTown'].tolist())  
                            for i in range(-10,11):
                                
                                if 'A' in str(room):
                                    
                                    neighbour = str(int(str(room)[1:]) + i)
                                    if  neighbour in  rooms_names_2:
                                        if len(neighbors_dict_2[str(room)])==0:
                                            neighbors_dict_2[str(room)].append('A'+neighbour)
                                            for j in df2['Room No. ']:
                                                if str(j) == 'B'+neighbour:
                                                    room_dict_2[str(room)].append(df2['Name'].tolist())
                                                                                
                                elif 'B' in str(room):
                                    neighbour = str(int(str(room)[1:]) + i)
                                    if  neighbour in  rooms_names_2:
                                        if len(neighbors_dict_2[str(room)])==0:
                                            neighbors_dict_2[str(room)].append('B'+neighbour)
                                            for j in df2['Room No. ']:
                                                # print(str(j),neighbour)
                                                if str(j) == 'B'+neighbour:
                                                    room_dict_2[str(room)].append(df2['Name'].tolist())
                                                    # print(df2['Name'].tolist())
                                                    # line()
                                            
                                                    

                            

    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_2.append(str(listdf1[n]))
                    if len(room_dict_2[str(room)])==0:
                                room_dict_2[str(room)].append(df2['Name'].tolist())                    
                                room_dict_2[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_2[str(room)].append(df2['HomeTown'].tolist())                    
                    
                n+=1
        
        case '1':
            df1 = dataframe.loc[dataframe['Block No.'] == block]
            n = 0
            block_1_dict[str(block)].append(df1['Name'].tolist())
            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room]
                
                if str(room) in rooms_names_1:
                    
                    for length in range(len(rooms_names_1)):
                        if str(room) == rooms_names_1[length]:
                            if len(room_dict_1[str(room)])==0:
                                room_dict_1[str(room)].append(df2['Name'].tolist())
                                room_dict_1[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_1[str(room)].append(df2['HomeTown'].tolist())                    
                                # print(str(room))
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_1.append(str(listdf1[n]))
                    if len(room_dict_1[str(room)])==0:
                                room_dict_1[str(room)].append(df2['Name'].tolist())
                                room_dict_1[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_1[str(room)].append(df2['HomeTown'].tolist()) 
                n+=1
        case '3':
            df1 = dataframe.loc[dataframe['Block No.'] == block]
            n = 0
            block_3_dict[str(block)].append(df1['Name'].tolist())
            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room]
                if str(room) in rooms_names_3:
                    for length in range(len(rooms_names_3)):
                        if str(room) == rooms_names_3[length]:
                            if len(room_dict_3[str(room)])==0:
                                room_dict_3[str(room)].append(df2['Name'].tolist())
                                room_dict_3[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_3[str(room)].append(df2['HomeTown'].tolist()) 
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_3.append(str(listdf1[n]))
                    if len(room_dict_3[str(room)])==0:
                                room_dict_3[str(room)].append(df2['Name'].tolist())
                                room_dict_3[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_3[str(room)].append(df2['HomeTown'].tolist()) 
                n+=1
        case '4':
            df1 = dataframe.loc[dataframe['Block No.'] == block]
            n = 0

            for room in df1['Room No. ']:
                df2 = df1.loc[df1['Room No. '] == room]
                if str(room) in rooms_names_4:
                    for length in range(len(rooms_names_4)):
                        if str(room) == rooms_names_4[length]:
                            if len(room_dict_4[str(room)])==0:
                                room_dict_4[str(room)].append(df2['Name'].tolist())
                                room_dict_4[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_4[str(room)].append(df2['HomeTown'].tolist())
    
                else:
                    listdf1 = df1['Room No. '].tolist()
                    rooms_names_4.append(str(listdf1[n]))
                    if len(room_dict_4[str(room)])==0:
                                room_dict_4[str(room)].append(df2['Name'].tolist())
                                room_dict_4[str(room)].append(df2['Branch with specialization (if any)'].tolist())                    
                                room_dict_4[str(room)].append(df2['HomeTown'].tolist())
                n+=1
        case _:
            print("Bhai Block 5 Under Construction")
            pass

# print(room_dict_1)
data2 = json.loads(json.dumps(room_dict_2))
# print("room_dict_two: ",data2)
# data2 = json.loads(json.dumps(neighbors_dict_2))
# print("room_dict_two: ",data2)

# data1 = json.loads(json.dumps(room_dict_1))
# print("room_dict_one: ",data1)
# # # print(rooms_names_1)
# print("room__two:",rooms_names_2)
# print("room__one:",rooms_names_1)
# print("room__three:",rooms_names_3)

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

   
    

    

    df2 = dataframe.loc[dataframe['Email Address'] == email]
    # print(email)
    blklist = df2['Block No.'].tolist()
    blockno = ''.join([str(element) for element in blklist]) 
    
    match str(blockno):
        case '2':
            roomlist = df2['Room No. '].tolist()
            # print(roomlist)
            roomno = ''.join([str(element) for element in roomlist]) 
            # print(room_dict_2[roomno])
        # print(str(email))

            sender = "quilo.bikcodes@gmail.com"
            context = f"Room: {roomno} \nBlock {blockno} \nSo these are your roommates:"
            loop = 0
            for i in room_dict_2[roomno][0]:
                j = room_dict_2[roomno][1][loop]
                k = room_dict_2[roomno][2][loop]
                context = context + "\n\n\n" + i +" studying in the branch of  " +j + " from " + k
                loop += 1 
            context = context + "\n\n\n" + f"These are the hostel mates of your block 2: \n"
            for name in block_2_dict['2'][0]:
                context = context + "\n"+ name
            context = context + "\n\nThanks for filling the G Forms\n\n\n\nFrom the developer \nI am Bikramjeet Dasgupta \nCSE \nAI ML \nFresher\nBlock 1\nRoom No.: 548\nI shall be adding neighbours also from the next mail. If You dont get a room mate no worries...The sample size is only 140 so ya It is quite often to happen I would rather say that matched roommates will be rare."
            print(context)
            subject = f'Hey Mate of Room {roomno}. Here are your roommates'
            msg = create_message(email, context , subject )
            send_message(service, 'me', msg)
            
                  
            
                   
            
        case '1':
            roomlist = df2['Room No. '].tolist()
            
            roomno = ''.join([str(element) for element in roomlist]) 
            # print(room_dict_2[roomno])
        # print(str(email))

            sender = "quilo.bikcodes@gmail.com"
            context = f"Room: {roomno} \nBlock {blockno} \nSo these are your roommates:"
            loop = 0
            for i in room_dict_1[roomno][0]:
                j = room_dict_1[roomno][1][loop]
                k = room_dict_1[roomno][2][loop]
                context = context + "\n" + i +" studying in the branch of  " +j + " from " + k
                loop += 1 
            context = context + "\n\n\n" + f"These are the hostel mates of your block 1: \n"
            for name in block_1_dict['1'][0]:
                context = context + "\n"+ name
            context = context + "\n\nThanks for filling the G Forms\n\n\n\nFrom the developer \nI am Bikramjeet Dasgupta \nCSE \nAI ML \nFresher\nBlock 1\nRoom No.: 548\nI shall be adding neighbours also from the next mail. If You dont get a room mate no worries...The sample size is only 140 so ya It is quite often to happen I would rather say that matched roommates will be rare."
            print(context)
            subject = f'Hey Mate of Room {roomno}. Here are your roommates'
            msg = create_message(email, context , subject )
            send_message(service, 'me', msg)  
        
        case '3':

            roomlist = df2['Room No. '].tolist()
            
            roomno = ''.join([str(element) for element in roomlist]) 
            # print(room_dict_2[roomno])
        # print(str(email))

            sender = "quilo.bikcodes@gmail.com"
            context = f"Room: {roomno} \nBlock {blockno} \nSo these are your roommates:"
            loop = 0
            for i in room_dict_3[roomno][0]:
                j = room_dict_3[roomno][1][loop]
                k = room_dict_3[roomno][2][loop]
                context = context + "\n" + i +" studying in the branch of  " +j + " from " + k
                loop += 1 
            context = context + "\n\n\n" + f"These are the hostel mates of your block 3: \n"
            for name in block_3_dict['3'][0]:
                context = context + "\n"+ name
            context = context + "\n\nThanks for filling the G Forms\n\n\n\nFrom the developer \nI am Bikramjeet Dasgupta \nCSE \nAI ML \nFresher\nBlock 1\nRoom No.: 548\nI shall be adding neighbours also from the next mail. If You dont get a room mate no worries...The sample size is only 140 so ya It is quite often to happen I would rather say that matched roommates will be rare."
            print(context)
            subject = f'Hey Mate of Room {roomno}. Here are your roommates'
            msg = create_message(email, context , subject )
            send_message(service, 'me', msg)  
        


    line()
    sleep(0.2)


