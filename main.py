import win32com.client
import re

import pandas as pd

WORKING_DIRECTORY = 'C:/Users/randy/Documents/'

def read_oor_emails(filename:str=None, save_to_excel:bool=False):
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    mail_folder = mapi.GetDefaultFolder(6).Folders['Odometer OOR']
    emails = mail_folder.items
    rows = []
    row = {}

    for email in emails:
        nums = [float(s) for s in re.findall(r'[-+]?(?:\d*\.*\d+)', email.body)]
        date = pd.to_datetime(email.ReceivedTime.date())
        row = {
            'Date': date.date(),
            'Tractor #': nums[0],
            'Miles': nums[1],
            'Percentage': nums[2],
            'Order Number': nums[3]
        }
        #Filter out bad data
        if row['Order Number'] > 190000 and row['Miles'] < 1000:
            rows.append(row)
        row = {}
    dataframe = pd.DataFrame(rows)
    if save_to_excel and filename:
        print(f'Saved {filename}.xlsx in Documents Folder')
        dataframe.to_excel(f'{WORKING_DIRECTORY}{filename}.xlsx')