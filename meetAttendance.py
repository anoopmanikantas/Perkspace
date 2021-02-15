# NOTE: This code works only with google chrome
# Add chrome browser to your system's path
# Get google sheets api credentials form the following url (https://developers.google.com/sheets/api/quickstart/python) and paste it in the same folder as the code.
# Run chrome.cmd file from this repository to start chrome in debugging mode

from __future__ import print_function
import pickle
import re
import os
import string
import datetime
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from collections import defaultdict
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# 1.attendance 
# 2.score 
# 3.score_sheet 
# 4.sub_sheet{createsheet{createworksheets}} 
# 5. mail 
# 6.if__name__ == "__main__" 


a = list(string.ascii_uppercase)
a = a*3
ranges = []
for i, j in enumerate(a):
    if i < 26:
        ranges.append(f'Sheet1!{j}1')
        continue
    elif i > 25 and i < 52:
        ranges.append(f'Sheet2!{j}1')
    else:
        ranges.append(f'Sheet3!{j}1')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']


def attendance(path, path2):
    """
    This module extracts data from meet and saves the chats to text file and marks attendance in excel sheet.
    param: Path of the file and name of the file.
    """
    opt = Options()
    opt.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driverLoc = ''
    if os.path.exists('chromedriverloc.txt'):
        with open('chromedriverloc.txt', 'r') as f:
            driverLoc = f.readline()
    else:
        driverLoc = input("Enter the path of chrome driver (This is only for the first time, details entered will\
be saved in chromedriverloc.txt, do not delete the file) \n>>> ")
        with open('chromedriverloc.txt', 'a') as f:
            f.write(driverLoc)
    driver = webdriver.Chrome(driverLoc, options=opt)

    date = datetime.date.today()
    d = date.strftime('%d-%m-%Y')
    date = date.strftime('%d/%m/%Y')
    lis = defaultdict(list)

    try:
        driver.find_element_by_xpath(
            r'//*[@id="ow3"]/div[1]/div/div[8]/div[3]/div[6]/div[3]/div/div[2]/div[3]/span/span'
        ).click()
        print("\nChats open")
        sleep(2)

    except:
        try:
            driver.find_element_by_xpath(
                r'//*[@id="ow3"]/div[1]/div/div[8]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[2]'
            ).click()
            sleep(2)
        except:
            pass

        print('\nChats open')

    finally:
        x = driver.find_elements_by_class_name('GDhqjd')
        for i in x:
            lis['name'].append(i.find_element_by_class_name('YTbUzc').text)
            lis['replies'].append(i.find_element_by_class_name('Zmm6We').text)

    for i, j in enumerate(lis['replies']):
        x = j.split("\n")
        x = ' '.join(x)
        lis['replies'][i] = x

    if not os.path.exists("chats"):
        os.mkdir("chats")

    with open("chats/" + f'Meet_Chat_History_{d}_{path2}.txt', 'a', encoding='utf-8') as f:
        for i, j in zip(lis['name'], lis['replies']):
            f.write(i + ':-- ' + j + '\n')

    att = []
    sk1 = input("\nEnter the first secret key given to students: \n>>> ")
    sk2 = input("\nEnter the second secret key given to students (If there is no second key press n): \n>>> ")

    with open("chats/" + f'Meet_Chat_History_{d}_{path2}.txt', 'r', encoding='utf-8') as f:
        x = f.readlines()

    print(f"\nReading xlsx file from {path}")

    df = pd.read_excel(path, engine='openpyxl')
    df.drop(df.columns[df.columns.str.contains(
        'unnamed', case=False)], axis=1, inplace=True)
    print(f'\nCreating a column named "{date}"..')
    df[f'{date}'] = ['ab']*len(df['name'])

    for i in x:
        try:
            if (sk1.lower() in i.lower()) or (sk2.lower() in i.lower()):
                x = re.split(':-- ', i)
                print(x[0])
                att.append(x[0])
        except:
            print("Given key is not alphabetic\n\n Trying numeric key....\n\n")
            if (sk1 in i.lower()) or (sk2 in i.lower()):
                x = re.split(':-- ', i)
                print(x[0])
                att.append(x[0])
    print("Successfully checked")
    
    print("\nParticipants identified with secret key (Note: This count is with repetition): ", len(att))
    print("\nMarking attendance...")

    for i, j in enumerate(df['name']):
        for z in att:
            if (j.lower() in z.lower()) or (df['usn'][i].lower() in z.lower()):
                df[f'{date}'][i] = 'present'

    df.to_excel(path)

    print(f"\nAttendance marked and saved in {path}")
    print('\nTotal participants attended today (Note:This count is without repetition): ', sum(
        [1 for i in df[f'{date}'] if i.lower() == 'present']))
    driver.close()


def score(path):
    """
    This module marks the score in the excel sheet (score gets incremented if the student is present).
    param: Path of the excel sheet.
    return: Name of the file
    """

    path2 = re.split(".xlsx", path)[0]
    if os.path.exists("sheets/total") == False:
        os.mkdir("sheets/total")

    path2 = "sheets/total/"+path2.split('\\')[-1]
    df = pd.read_excel(path)
    d = defaultdict(list)
    d['name'] = list(df['name'])
    d['usn'] = list(df['usn'])
    d['score'] = [0] * len(d['name'])
    d2 = defaultdict(list)
    date = datetime.date.today()
    date = date.strftime("%d/%m/%Y")

    if os.path.exists(path2+"_total.xlsx"):
        df2 = pd.read_excel(path2+"_total.xlsx")
        d2['name'] = list(df2['name'])
        d2['usn'] = list(df2['usn'])
        d2['score'] = list(df2['score'])
        for i, j in enumerate(df[f'{date}']):
            if j == 'present':
                d2['score'][i] += 1
            else:
                d2['score'][i] += 0

        df2 = pd.DataFrame(d2)
        df2.to_excel(path2+"_total.xlsx")

    else:
        for i, j in enumerate(df[f'{date}']):
            print(i)
            if j == 'present':
                d['score'][i] += 1

            else:
                d['score'][i] += 0

        df2 = pd.DataFrame(d)
        df2.to_excel(path2+"_total.xlsx")

    return path2


def score_sheet(path):
    """
    This module updates the score sheet in drive.
    param: Name of the file
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    print(creds)

    date = datetime.date.today()
    date = date.strftime("%d/%m/%Y")

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    sleep(2)
    sub_name = path
    x = os.listdir("sheets/total")
    print("\nFiles in the directory -> ", x)

    names = [i for i in x if sub_name in re.search(r"\w+.xlsx", i)[0]][0]
    print("\nFile read -> ", names)

    f_names = re.split(".xlsx", names)[0]
    data = defaultdict(list)
    if not os.path.exists('id'):
        os.mkdir('id')

    if os.path.exists("id/sheetIDs.csv"):
        data = pd.read_csv("id/sheetIDs.csv")

    else:

        spreadsheet = {
            'properties': {
                'title': 'a_'+f_names
            }
        }
        spreadsheet = sheet.create(
            body=spreadsheet, fields='spreadsheetId').execute()
        print("Spreadsheet id: ", spreadsheet.get('spreadsheetId'))
        data['name'].append('a_'+f_names)
        data['id'].append(spreadsheet.get("spreadsheetId"))

    data = pd.DataFrame(data)
    data.to_csv('id/sheetIDs.csv')
    df = pd.read_excel("sheets/total/"+f_names+'.xlsx')
    spredid = data['id'][list(data['name']).index('a_'+f_names)]

    print("Spread sheet id: ", str(spredid))
    l = [[i, j, k] for i, j, k in zip(df['name'], df['usn'], df['score'])]

    req = sheet.values().update(
        spreadsheetId=str(spredid),
        range='B2:D',
        valueInputOption="USER_ENTERED",
        body={'values': l[:len(l)]}
    ).execute()

    print("\nUpdated all the columns and rows\nResult: ", req)
    req = sheet.values().update(
        spreadsheetId=str(spredid),
        range='B1:D1',
        valueInputOption="USER_ENTERED",
        body={'values': [['name', 'usn', "score"]]}
    ).execute()

    print("\nUpdated all the column names\nResult: ", req)
    print("\nDone...!\nThank you!")


def sub_sheet(path, path2):
    """
    This module adds/ updates the attendance sheet in the drive
    param: path of the file and name of the file
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    print(creds)
    date = datetime.date.today()
    date = date.strftime("%d/%m/%Y")

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    df = pd.read_excel('id/subsheets.xlsx')
    a = defaultdict(list)
    a['name'].extend(list(df['name']))
    a['id'].extend(list(df['id']))
    print(a['name'])
    print(path2)

    x = os.listdir("sheets")
    print(x)
    names = [i for i in x if path in i][0]
    print(names)
    f_names = re.split(".xlsx", names)[0]  # gets file name

    spreadsheet_id = None

    def createsheet(path, a):
        """
        This module creates a spreadsheet if it does not exist in drive
        param: path of the file
        return: id of the spreadsheet
        """
        spreadsheet = {
            'properties': {
                'title': f_names
            }
        }

        spreadsheet = sheet.create(
            body=spreadsheet, fields='spreadsheetId').execute()
        print(spreadsheet.get('spreadsheetId'))
        id1 = spreadsheet.get('spreadsheetId')
        a['name'].append(f_names)
        a['id'].append(id1)
        a = pd.DataFrame(a)
        a.to_excel('id/subsheets.xlsx')

        def createworksheets(id, sheet_name):
            """
            This module creates multiple worksheets inside a spreadsheet
            param: id of the sheet and name of the sheet
            """
            request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                        }
                    }
                }]
            }
            response = sheet.batchUpdate(
                spreadsheetId=id,
                body=request_body
            ).execute()

            print("Created worksheets: ", response)
        for i in ['Sheet2', 'Sheet3']:
            createworksheets(id1, i)

        df2 = pd.read_excel("sheets/"+path)
        l = [[i, j] for i, j in zip(list(df2['name']), list(df2['usn']))]

        for i in ['Sheet1', 'Sheet2', 'Sheet3']:

            req = sheet.values().update(
                spreadsheetId=id1,
                range=f"{i}!A2:B",
                valueInputOption="USER_ENTERED",
                body={'values': l[:len(l)]}
            ).execute()
            print(req)

            req = sheet.values().update(
                spreadsheetId=id1,
                range=f'{i}!A1:B1',
                valueInputOption="USER_ENTERED",
                body={'values': [['name', 'usn']]}
            ).execute()
            print(req)

        return id1

    if path2 in ' '.join(a['name']):
        print('a')
        spreadsheet_id = a['id'][a['name'].index(path2)]
        
    else:
        print("b")
        spreadsheet_id = createsheet(path, a)
        

    value_render_option = 'FORMULA'
    date_time_render_option = 'SERIAL_NUMBER'
    print("before request")
    request = sheet.values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges)
    print("re")
    response = request.execute()
    print(response)

    # generating standard form of row and column id for spreadsheet
    c_range = None
    for i in response['valueRanges']:
        if 'values' not in ' '.join(i.keys()):
            c_range = i['range']
            break

    print(c_range)
    r_range = c_range.split('!')
    r_range = r_range[0] + "!" + ''.join(r_range[-1].split(
        '1')) + '2' + ':' + ''.join(r_range[-1].split('1'))
    print(r_range)
    df = pd.read_excel("sheets/"+f_names+'.xlsx')
    vales = [f'{date}']
    l = [[k] for k in df[f'{date}']]
    print(l)
    # update rows
    req = sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=r_range,
        valueInputOption="USER_ENTERED",
        body={'values': l[:len(l)]}
    ).execute()
    print(req)
    # update column names
    req = sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=c_range,
        valueInputOption="USER_ENTERED",
        body={'values': [vales]}
    ).execute()
    print(req)


def mail(fname):
    date = datetime.datetime.today()
    date = date.strftime('%d/%m/%Y')
    fromaddr = input("Enter sender's mail id: \n>>> ")
    # toaddr = input("Enter receiver's mail id:\n>>>")
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    # msg['To'] = toaddr
    sub = input("Enter the subject name: ")
    msg['Subject'] = sub +" Attendance"+f" {date}"
    body = f"Enter body of the mail"
    
    msg.attach(MIMEText(body,'plain'))
    smtp = smtplib.SMTP('smtp.gmail.com',587)
    smtp.starttls()
    smtp.login(fromaddr, "Enter your password")
    tex = msg.as_string()
    df = pd.read_excel("sheets/mails.xlsx")
    d = defaultdict(list)
    d['mail'].extend(list(df['mail']))
    d['status'].extend(df['status'])
    df2 = pd.read_excel(fname)
    for i, j in enumerate(df2[f'{date}']):
        if j == 'present' and d['status'][i] != 'p':
            print(d['mail'][i])
            d['status'][i] = 'p'
            smtp.sendmail(fromaddr, d['mail'][i], tex)
    d = pd.DataFrame(d)
    d.to_excel("sheets/mails.xlsx")
    smtp.quit()
    print("\n\nMail sent")


if __name__ == '__main__':
    print("Welcome".center(90, '-'))
    print("\n\nExtracting participants data.......\n")

    path = input(
        '\nEnter the path of attendance sheet (File must be of xlsx type only. Example: "D:/user/py/test2.xlsx"):\n>>> ')
    fname = path

    path2 = re.split(".xlsx", path)[0]
    path2 = path2.split('\\')[-1]
    path3 = path.split('\\')[-1]
    path4 = path3.split('.xlsx')[0]

    attendance(path, path2)
    a = input("\nWould you like to update the attendance sheet in drive? Y or n\n>>> ")
    if a.lower() == 'y':
        print("Updating attendance sheet...")
        try:
            sub_sheet(path3, path4)
        except Exception as e:
            print(e)

    else:
        print("\nYou chose not to update")
    print("\nConverting data.......")

    try:
        path = score(path).split('/')
    except Exception as e:
        print(e)

    a = input("\nWould you like to update the score sheet in drive? Y or n\n>>> ")

    if a.lower() == 'y':
        try:
            print("\nUpdating in spreadsheets")
            score_sheet(path[-1])
            input("\nPress enter to quit.. ")
        except Exception as e:
            print(e)

    else:
        print("\nYou chose not to update\nThank you!")
        input("\nPress enter to quit")

    a = input("\nWould you like to mail? Y or n\n>>> ")
        
    if a.lower() == 'y':
        try:
            mail(fname)
        except Exception as e:
            print(e)
