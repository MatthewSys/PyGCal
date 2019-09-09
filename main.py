import xlrd

import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from tkinter import *
#from tkinter.ttk import *

year = datetime.datetime.today().year

SCOPES = ['https://www.googleapis.com/auth/calendar']
CREDENTIALS_FILE = 'Secr.json'


def get_calendar_service():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service


def main():
    CID = str(txt.get())
    name = int(txt2.get())
    path = str(txt3.get())
    newlist = []
    wb = xlrd.open_workbook(path)
    sheet_c = len(wb.sheets()) - 1
    print(sheet_c)
    for i in range(sheet_c):
        sheet = wb.sheet_by_index(i)
        days = sheet.row_values(11, 6, 37)
        days = list(filter(lambda elm: isinstance(elm, float), days))
        M = max(days)
        M = int(M)
        # print(M)
        if name == 1:
            c1 = 13
            c2 = 6
        elif name == 2:
            c1 = 18
            c2 = 6
        elif name == 3:
            c1 = 23
            c2 = 6
        elif name == 4:
            c1 = 28
            c2 = 6
        elif name == 5:
            c1 = 33
            c2 = 6
        else:
            break
        sun = sheet.row_values(c1, c2, M + 6)
        dict = {"-8": 1, "8": 1, "8-": 2, '': 0, "отпуск": 0, -8.0: 1}
        new = [dict.get(n, n) for n in sun]
        newlist.insert(i, [new])
        continue
    b = newlist
    # print(len(b))

    for mounth in range(len(b)):
        a = b[mounth][0]
        print(a)
        for smenes in range(len(a)):
            s = a[smenes]
            i1 = str(mounth + 1)
            j1 = str(smenes + 1)
            if s == 2 and not mounth == 0:
                ss = str(year) + "-"
                min = ('-')
                kk = ss + i1 + min + j1
                print(kk, "day")
                service = get_calendar_service()
                event_result = service.events().insert(
                    calendarId=CID,
                    body={
                        "summary": 'DAY',
                        "description": 'DAY OF PAIN',
                        "start": {"dateTime": kk + 'T01:00:00-04:00',
                                  "timeZone": 'Europe/Moscow'},
                        "end": {"dateTime": kk + 'T09:00:00-08:00',
                                "timeZone": 'Europe/Moscow'},
                    }
                ).execute()
            elif s == 2 and mounth == 0:
                ss = str(year) + "-"
                min = ('-')
                kk = ss + i1 + min + j1
                print(kk, "night")
                service = get_calendar_service()
                event_result = service.events().insert(calendarId=CID,
                                                       body={

                                                           "summary": 'NIGHT',
                                                           "description": 'NIGHT OF PAIN',
                                                           "start": {"dateTime": kk + 'T17:00:00-00:00',
                                                                     "timeZone": 'Europe/Moscow'},
                                                           "end": {"dateTime": kk + 'T23:00:00-06:00',
                                                                   "timeZone": 'Europe/Moscow'},
                                                       }
                                                       ).execute()

            elif s == 1 and not smenes == 0:
                smenes = str(smenes)
                ss = str(year) + "-"
                min = ('-')
                kk = ss + i1 + min + smenes
                print(kk, "night")
                service = get_calendar_service()
                event_result = service.events().insert(calendarId=CID,
                                                       body={

                                                           "summary": 'NIGHT',
                                                           "description": 'NIGHT OF PAIN',
                                                           "start": {"dateTime": kk + 'T17:00:00-00:00',
                                                                     "timeZone": 'Europe/Moscow'},
                                                           "end": {"dateTime": kk + 'T23:00:00-06:00',
                                                                   "timeZone": 'Europe/Moscow'},
                                                       }
                                                       ).execute()

            elif s == 1 and smenes == 0 and not mounth == 0:
                i1 = str(mounth)
                j1 = str(len(newlist[mounth - 1][0]))
                ss = str(year) + "-"
                min = ('-')
                kk = ss + i1 + min + j1
                print(kk, "night")
                service = get_calendar_service()
                event_result = service.events().insert(calendarId=CID,
                                                       body={

                                                           "summary": 'NIGHT',
                                                           "description": 'NIGHT OF PAIN',
                                                           "start": {"dateTime": kk + 'T17:00:00-00:00',
                                                                     "timeZone": 'Europe/Moscow'},
                                                           "end": {"dateTime": kk + 'T23:00:00-06:00',
                                                                   "timeZone": 'Europe/Moscow'},
                                                       }
                                                       ).execute()
        print()


window = Tk()

window.title("PyGCal")
window.geometry('300x150')

lbl = Label(window, text="Введи ID")
lbl.grid(column=0, row=0)
txt = Entry(window, width=10)
txt.grid(column=1, row=0)

lbl2 = Label(window, text="Введи номер в календаре")
lbl2.grid(column=0, row=2)
txt2 = Entry(window, width=10)
txt2.grid(column=1, row=2)

lbl3 = Label(window, text="Введи путь к файлу")
lbl3.grid(column=0, row=3)
txt3 = Entry(window, width=10)
txt3.grid(column=1, row=3)

btn = Button(window, text="Обработать", command=main)

btn.grid(column=1, row=5)

if __name__ == '__main__':
    window.mainloop()
