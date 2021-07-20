from __future__ import print_function
import os.path
import json
import xlwt as xl
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
SCRIPT = 'https://github.com/Fifirex/mail-thread-attendance/blob/main/script.py'
SEARCH_SUBJECT = 'Meeting: 20 July 2021, @9:45pm'
SEARCH_MSG = '+1'
SEARCH_MSG_NEG = '-1'
XL_PATH = 'database/autoGenAttendance.xls'
DATE = '20 July 2021'
TOT_COUNT = 36

def Looker(str):
    newStr = ""
    flag = False
    for i in range(len(str)):
        if(str[i].isalpha() or str[i] == ' '):
            flag = True
            newStr += str[i]
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
                    align: wrap on, horiz centre, vert centre;\
                    pattern: pattern solid, fore_colour gray40;\
                    borders: left thick, right thick, bottom thick, top thick"
    style_header = xl.easyxf(style_string)

    style_string = "font: bold on;\
                    align: wrap on, horiz centre, vert centre;\
                    borders: left thick, right thick, bottom thick, top thick"
    style_null = xl.easyxf(style_string)

    style_string = "font: bold on;\
                    align: wrap on, horiz centre, vert centre;\
                    pattern: pattern solid, fore_colour gray25;\
                    pattern: pattern solid, fore_colour light_green;\
                    borders: left thick, right thick, bottom thick, top thick"
    plus_header = xl.easyxf(style_string)

    style_string = "font: bold on;\
                    align: wrap on, horiz centre, vert centre;\
                    pattern: pattern solid, fore_colour coral;\
                    borders: left thick, right thick, bottom thick, top thick"
    minus_header = xl.easyxf(style_string)

    style_string = "font: bold on;\
                    align: wrap on, horiz centre, vert centre;\
                    pattern: pattern solid, fore_colour light_yellow;\
                    borders: left thick, right thick, bottom thick, top thick"
    null_header = xl.easyxf(style_string)

    sheet.write(0, 0, "NAMES", style = style_header)
    sheet.write(0, 1, "STATE", style = style_header)
    sheet.write(0, 2, "RESPECTIVE REASONS", style = style_header)
    sheet.write_merge(0, 0, 3, 4, "SUMMARY", style = style_header)
    sheet.write(1, 3, "DATE", style = style_null)
    sheet.write(2, 3, "RESP_CTR", style = style_null)
    sheet.write(3, 3, "PLUS_CTR", style = plus_header)
    sheet.write(4, 3, "MINS_CTR", style = minus_header)
    sheet.write(5, 3, "NR_CTR", style = null_header)
    sheet.write(6, 3, "SCRIPT", style = style_null)

    style_string = "font: bold off;\
                    align: wrap on, horiz centre, vert centre;\
                    borders: left thick, right thick, bottom thick, top thick"
    style = xl.easyxf(style_string)
    sheet.write(1, 4, DATE, style = style)
    sheet.write(2, 4, str(plus_ctr + minus_ctr), style = style)
    sheet.write(3, 4, str(plus_ctr), style = style)
    sheet.write(4, 4, str(minus_ctr), style = style)
    sheet.write(5, 4, str(TOT_COUNT - plus_ctr - minus_ctr), style = style)
    sheet.write(6, 4, xl.Formula('HYPERLINK("%s";"script")' % SCRIPT), style = style)

    style_string = "align: wrap on"
    style = xl.easyxf(style_string)

    style_string = "align: wrap on;\
                    borders: left thin, right thin, bottom thin;\
                    pattern: pattern solid, fore_colour gray25"
    name_style = xl.easyxf(style_string)

    style_string = "align: wrap on;\
                    borders: left thin, right thin, bottom thin;\
                    pattern: pattern solid, fore_colour light_green"
    plus_style = xl.easyxf(style_string)

    style_string = "align: wrap on;\
                    borders: left thin, right thin, bottom thin;\
                    pattern: pattern solid, fore_colour coral"
    minus_style = xl.easyxf(style_string)

    style_string = "align: wrap on;\
                    borders: left thin, right thin, bottom thin;\
                    pattern: pattern solid, fore_colour light_yellow"
    null_style = xl.easyxf(style_string)

    file = open("database/data.json")
    data = json.load(file)

    row = 1
    for email in plus_list:
        str1 = Looker(email)
        sheet.write(data[str1][1], 0, data[str1][0], style = name_style)
        sheet.write(data[str1][1], 1, "+1", style = plus_style)
        data[str1][2] = True
        row+=1

    i = 1
    for email in minus_list:
        str2 = Looker(email)
        sheet.write(data[str2][1], 0, data[str2][0], style = name_style)
        sheet.write(data[str2][1], 1, "-1", style = minus_style)
        sheet.write(data[str2][1], 2, reason_list[i - 1], style = style)
        data[str2][2] = True
        row+=1
        i+=1

    for em in data:
        if not data[em][2]:
            sheet.write(data[em][1], 0, data[em][0], style = name_style)
            sheet.write(data[em][1], 1, "NR", style = null_style)

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

            print ('resp_ctr : %d' % (plus_ctr + minus_ctr)) 
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