import random
import time,os
import requests
import csv
import openpyxl
from pathlib import Path
import pandas as pd
import json 
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

total_records=1500
datas = requests.get(f"https://data.nashville.gov/api/id/479w-kw2x.json?$select=*&$order=`date_received`+DESC&$limit={total_records}&$offset=0&$$read_from_nbe=true&$$version=2.1")

datas_recent = requests.get("https://data.nashville.gov/api/id/479w-kw2x.json?$select=*&$order=`date_received`+DESC&$limit=1&$offset=0&$$read_from_nbe=true&$$version=2.1")


recent_date=(datetime.datetime.strptime(datas_recent.json()[0]['date_received'].split('T')[0], '%Y-%m-%d'))
    
limit = datetime.timedelta(days = 10)
min_date=recent_date-limit
print(recent_date)
print(min_date)
print('Collecting Data')
last_date=datetime.datetime.strptime(datas.json()[-1]['date_received'].split('T')[0], '%Y-%m-%d')
while True:
    if last_date > min_date:
        total_records=total_records+1000
        datas = requests.get(f"https://data.nashville.gov/api/id/479w-kw2x.json?$select=*&$order=`date_received`+DESC&$limit={total_records}&$offset=0&$$read_from_nbe=true&$$version=2.1")
        last_date=datetime.datetime.strptime(datas.json()[-1]['date_received'].split('T')[0], '%Y-%m-%d')
    else:
        break
print('Data Collected!')
data_recent=[]
data_rest=[]
for i in (datas.json()):
    if recent_date==datetime.datetime.strptime(i['date_received'].split('T')[0], '%Y-%m-%d'):
        try:
            i.update({'mapped_location':str(i['mapped_location']['coordinates']).replace('[','').replace(']','')})
#             print(i['mapped_location'])
        except Exception as e:
            pass
#         i['mapped_location']=str(i['mapped_location']['coordinates']).replace('[','').replace(']','')
        data_recent.append(i)
    else:
        if datetime.datetime.strptime(i['date_received'].split('T')[0], '%Y-%m-%d')>=min_date:
            try:
                i.update({'mapped_location':str(i['mapped_location']['coordinates']).replace('[','').replace(']','')})
#                 print(i['mapped_location'])
            except Exception as e:
                pass
#         i['mapped_location']=str(i['mapped_location']['coordinates']).replace('[','').replace(']','')
            data_rest.append(i)
df_recent = pd.DataFrame(data_recent)
df_rest = pd.DataFrame(data_rest)
gsheet_columns_raw=df_recent.columns.values.tolist()
gsheet_columns=[]
for i in gsheet_columns_raw:
    gsheet_columns.append(i.replace('_',' ').capitalize())
rest_data_for_gsheet=df_rest.fillna('').values.tolist()
recent_data_for_gsheet=df_recent.fillna('').values.tolist()
rest_data_for_gsheet.insert(0,gsheet_columns)
recent_data_for_gsheet.insert(0,gsheet_columns)
SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SAMPLE_SPREADSHEET_ID = '1brIiqmWaW3ge0A1HAQO3diD6-kEEJb1zPrrBPx40nCE'
try:
    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
                                                      
    sheet = service.spreadsheets()
    request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="Recent Day Data!A1", valueInputOption="USER_ENTERED", body={"values":recent_data_for_gsheet}).execute()                                         
    request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="Rest of the Data!A1", valueInputOption="USER_ENTERED", body={"values":rest_data_for_gsheet}).execute()                   
except HttpError as err:
    print(err)
print('Data Collected Successfully!')