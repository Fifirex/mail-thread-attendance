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
XL_PATH = 'autoGenAttendance.xls'
DATE = '15 July 2021'

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

def Writer(plus_list, minus_list, plus_ctr, minus_ctr, reason_list):
    wb = xl.Workbook()
    sheet = wb.add_sheet(DATE, cell_overwrite_ok=True)
    # column width adjusted as per need (taken from width.py)
    sheet.col(0).width = 6092
    sheet.col(1).width = 6092
    sheet.col(2).width = 720*20
    sheet.col(3).width = 2769
    sheet.col(4).width = 2769

    style_string = "font: bold on;\
                    align: wrap on, horiz centre;\
                    borders: left thick, right thick, bottom thick, top thick"
    style = xl.easyxf(style_string)
    sheet.write(0, 0, "PLUS ONES", style = style)
    sheet.write(0, 1, "MINUS ONES", style = style)
    sheet.write(0, 2, "RESPECTIVE REASONS", style = style)
    sheet.write_merge(0, 0, 3, 4, "SUMMARY", style = style)
    sheet.write(1, 3, "DATE", style = style)
    sheet.write(2, 3, "RESP_CTR", style = style)
    sheet.write(3, 3, "PLUS_CTR", style = style)
    sheet.write(4, 3, "MINS_CTR", style = style)
    sheet.write(5, 3, "SCRIPT", style = style)

    style_string = "font: bold off;\
                    align: wrap on, horiz centre;\
                    borders: left thick, right thick, bottom thick, top thick"
    style = xl.easyxf(style_string)
    sheet.write(1, 4, DATE, style = style)
    sheet.write(2, 4, str(plus_ctr + minus_ctr), style = style)
    sheet.write(3, 4, str(plus_ctr), style = style)
    sheet.write(4, 4, str(minus_ctr), style = style)
    sheet.write(5, 4, "https://sed.lol/fptw87", style = style)

    style_string = "align: wrap on"
    style = xl.easyxf(style_string)
    row = 1
    for email in plus_list:
        sheet.write(row, 0, Looker(email), style = style)
        row+=1

    row = 1
    for email in minus_list:
        sheet.write(row, 1, Looker(email), style = style)
        sheet.write(row, 2, reason_list[row - 1], style = style)
        row+=1

    wb.save(XL_PATH)
    print('xls is generated!!')

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
                    flag = True
                break

        if flag:
            plus_ctr = 0
            plus_list = []
            minus_ctr = 0
            minus_list = []
            reason_list = []
            for Dmsgs in tdata['messages']:
                if Dmsgs['snippet'][:2] == SEARCH_MSG:
                    msg = Dmsgs['payload']
                    for header in msg['headers']:
                        if header['name'] == 'From':
                            if(plus_list.count(header['value']) <= 0):
                                plus_list.append(header['value'])
                                plus_ctr += 1
                                if(minus_list.count(header['value']) > 0):
                                    minus_list.pop(minus_list.index(header['value']))
                                    minus_ctr -= 1
                            break

                elif Dmsgs['snippet'][:2] == SEARCH_MSG_NEG:
                    index = 0
                    while not Dmsgs['snippet'][index].isalpha():
                        index+=1
                    index2 = index
                    while (Dmsgs['snippet'][index2 : index2+2] != "On"):
                        index2+=1
                    index2-=1
                    while not Dmsgs['snippet'][index2].isalpha():
                        index2-=1
                    reason_list.append(Dmsgs['snippet'][index : index2+1])
                    msg = Dmsgs['payload']
                    for header in msg['headers']:
                        if header['name'] == 'From':
                            if(minus_list.count(header['value']) <= 0):
                                minus_list.append(header['value'])
                                minus_ctr += 1
                                if(plus_list.count(header['value']) > 0):
                                    plus_list.pop(plus_list.index(header['value']))
                                    plus_ctr -= 1
                            break

            print ('plus_ctr : %d' % plus_ctr)
            print ('mins_ctr : %d' % minus_ctr)
            break

    Writer(plus_list, minus_list, plus_ctr, minus_ctr, reason_list)

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