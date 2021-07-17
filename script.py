from __future__ import print_function
import os.path
import xlwt as xl
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
SEARCH_SUBJECT = 'Meeting: 15th July 2021 @5pm'
SEARCH_MSG = '+1'
SEARCH_MSG_NEG = '-1'

def Looker(str):
    newStr = []
    flag = False
    for i in range(len(str)):
        if(str[i].isalpha() or str[i] == ' '):
            flag = True
            newStr.append(str[i])
        elif flag:
            break
    return newStr

def Writer(plus_list, minus_list):
    wb = xl.Workbook()
    sheet = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)
    style_string = "font: bold on; borders: bottom dashed"
    style = xl.easyxf(style_string)
    sheet.write(0, 0, "PLUS ONES", style = style)
    sheet.write(0, 1, "MINUS ONES", style = style)
    row = 1
    for email in plus_list:
        sheet.write(row, 0, Looker(email))
        row+=1
    row = 1
    for email in minus_list:
        sheet.write(row, 1, Looker(email))
        row+=1
    wb.save('autoGenAttendance.xls')
    print('xls is generated')

def counter(service, user_id='me'):
    threads = service.users().threads().list(userId=user_id).execute().get('threads', [])
    flag = False
    for thread in threads:
        tdata = service.users().threads().get(userId=user_id, id=thread['id']).execute()

        msg = tdata['messages'][0]['payload']

        for header in msg['headers']:
            if header['name'] == 'Subject':
                subject = header['value']
                if (subject == SEARCH_SUBJECT):
                    print ("Found the thread!!")
                    nmsgs = len(tdata['messages'])
                    print ("nmsgs = %d" % nmsgs)
                    flag = True
                break

        if flag:
            plus_ctr = 0
            plus_list = []
            minus_ctr = 0
            minus_list = []
            for Dmsgs in tdata['messages']:
                if Dmsgs['snippet'][:2] == SEARCH_MSG:
                    plus_ctr += 1
                    msg = Dmsgs['payload']
                    for header in msg['headers']:
                        if header['name'] == 'From':
                            plus_list.append(header['value'])
                            break

                elif Dmsgs['snippet'][:2] == SEARCH_MSG_NEG:
                    minus_ctr += 1
                    msg = Dmsgs['payload']
                    for header in msg['headers']:
                        if header['name'] == 'From':
                            minus_list.append(header['value'])
                            break

            print ('plus_ctr : %d' % plus_ctr)
            print ('minus_ctr : %d' % minus_ctr)
            break
    Writer(plus_list, minus_list)

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    counter(service)

if __name__ == '__main__':
    main()