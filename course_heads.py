import gspread
from gspread_formatting import *
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import pandas as pd


def design(worksheet, rows_num):

    worksheet.update_cell(1, 7, 'Статус')
    worksheet.update_cell(1, 8, 'Комментарий')


    set_column_width(worksheet, 'A:G', 120)

    # Добавление столбца с выпадающими списками
    data_validation_rule = {
        "condition": {
            "type": "ONE_OF_LIST",
            "values": [
                {"userEnteredValue": "закрылся"},
                {"userEnteredValue": "задолж."},
                {"userEnteredValue": "академ."},
                {"userEnteredValue": "отчислен"}
            ]
        },
        "showCustomUi": True,
        "strict": True
    }
    request = {
        "setDataValidation": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,
                "endRowIndex": rows_num + 1,
                "startColumnIndex": 6,
                "endColumnIndex": 7
            },
            "rule": data_validation_rule
        }
    }
    worksheet.spreadsheet.batch_update({"requests": [request]})

    # Раскрашиавние в чередующиеся цвета
    start_row = 1
    end_row = rows_num + 1
    requests = []
    for row in range(start_row, end_row + 1):
        if row % 2 == 0:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 0,
                        "endColumnIndex": worksheet.col_count
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 1.0
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        else:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 0,
                        "endColumnIndex": worksheet.col_count
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.91,
                                "green": 0.94,
                                "blue": 1.0
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })

    batch_update_body = {"requests": requests}
    worksheet.spreadsheet.batch_update(batch_update_body)

    # Выриванивание текста и обозначение заголовка таблицы
    worksheet.format(
        "1:1", {
            "horizontalAlignment": "CENTER",
            "textFormat": {
            "fontSize": 10,
            "bold": True
            }
        })
    worksheet.format(
        "A2:H100", {
            "horizontalAlignment": "CENTER",
            "textFormat": {
            "fontSize": 10
            }
    })



# Аутентификация
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_file('credentials.json')
drive_service = build('drive', 'v3', credentials=credentials)
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)


# Загрузка нашей базы
table = client.open_by_key('19DuW3CRvYameij1eFNCyyijRyC9O7HtFvNjgbazg_zk')
df = [pd.DataFrame()] * 3 
for i in range(3):
    sheet = table.get_worksheet(i)
    df[i] = (pd.DataFrame(sheet.get_all_records()))[['Фамилия', 'Имя', 'Отчество', 'Направление', 'Курс', 'Студенческий']]


# Создание папки, где будут лежать 6 таблиц
parent_folder_id = '1c6BiUKFSt6nr7_jKcHOS2wJW5haoJxcw'
folder_metadata = {
    'name': 'Таблицы начальникам курсов',
    'parents': [parent_folder_id],
    'mimeType': 'application/vnd.google-apps.folder'
}
folder = drive_service.files().create(body=folder_metadata, fields='id').execute()


# Создание таблиц бакалавриата с доступом редактора для всех
sheet_names = ['Бюджет ЧП', 'Контракт ЧП']
for i in range(4):
    table = client.create(f'{i + 1} курс', folder_id=folder['id'])
    for j in range(2):
        worksheet = table.add_worksheet(title = sheet_names[j], rows = 100, cols = 10)
        tmp_df = df[j]
        tmp_df = tmp_df[((tmp_df['Курс'] == (i + 1)) | (tmp_df['Курс'] == f'{i + 1}(академ)')) & (tmp_df['Направление'] == 'б')]
        rows_num = len(tmp_df)
        worksheet.update([tmp_df.columns.values.tolist()] + tmp_df.values.tolist())
        design(worksheet, rows_num)

    table.del_worksheet(table.sheet1)
    table.share(None, perm_type='anyone', role='writer') 
    print(f'Таблица {i + 1} курса бакалавриата создана')

# Аналогично для магистратуры
for i in range(2):
    table = client.create(f'{i + 5} курс', folder_id=folder['id'])
    for j in range(2):
        worksheet = table.add_worksheet(title = sheet_names[j], rows = 100, cols = 10)
        tmp_df = df[j]
        tmp_df = tmp_df[((tmp_df['Курс'] == (i + 1)) | (tmp_df['Курс'] == f'{i + 1}(академ)')) & (tmp_df['Направление'] == 'м')]
        rows_num = len(tmp_df)
        worksheet.update([tmp_df.columns.values.tolist()] + tmp_df.values.tolist())
        design(worksheet, rows_num)

    table.del_worksheet(table.sheet1)
    table.share(None, perm_type='anyone', role='writer') 
    print(f'Таблица {i + 1} курса магистратуры создана')

