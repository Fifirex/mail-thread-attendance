from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
SEARCH_SUBJECT = 'Meeting: 15th July 2021 @5pm'
SEARCH_MSG = '+1'
SEARCH_MSG_NEG = '-1'

def show_chatty_threads(service, user_id='me'):
    threads = service.users().threads().list(userId=user_id).execute().get('threads', [])
    for thread in threads:
        tdata = service.users().threads().get(userId=user_id, id=thread['id']).execute()
        nmsgs = len(tdata['messages'])

        if nmsgs > 2:    # skip if <3 msgs in thread
            msg = tdata['messages'][0]['payload']
            subject = ''
            for header in msg['headers']:
                if header['name'] == 'Subject':
                    subject = header['value']
                    break
            if subject:  # skip if no Subject line
                print('- %s (%d msgs)' % (subject, nmsgs))
                print ('Message size is %d' % msg['body']['size'])

def counter(service, user_id='me'):
    threads = service.users().threads().list(userId=user_id).execute().get('threads', [])
    flag = False
    for thread in threads:
        tdata = service.users().threads().get(userId=user_id, id=thread['id']).execute()

        msg = tdata['messages'][0]['payload']

        # header sector
        # for collecting the subject
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
            # print ('Message size is %d' % msg['body']['size'])
            # str = tdata['raw']
            # base64_bytes = str.encode('ascii')
            # message_bytes = base64.b64decode(base64_bytes)
            # message = message_bytes.decode('ascii')
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
            # print (tdata['messages'][1]['snippet'])
            break

def main():
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
    # show_chatty_threads(service)
    counter(service)

    # Call the Gmail API
    # results = service.users().labels().list(userId='me').execute()
    # labels = results.get('labels', [])
    # msgs = results.get('messages', [])

    # if not msgs:
    #     print('No msgs found.')
    # else:
    #     print('msg ids:')
    #     for msg in msgs:
    #         print(msg['id'])

    # if not labels:
    #     print('No labels found.')
    # else:
    #     print('Labels:')
    #     for label in labels:
    #         print(label['name'])

if __name__ == '__main__':
    main()