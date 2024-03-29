import gspread
from gspread_formatting import *
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import pandas as pd
import pygsheets
import apiclient
import httplib2
from datetime import datetime

#-----------------------------------------------------------------------------------------#
#--------------------------------------ТАБЛИЦА--------------------------------------------#
#-----------------------------------------------------------------------------------------#

# Аутентификация
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_file('credentials.json')
drive_service = build('drive', 'v3', credentials=credentials)
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)

#Создание таблицы
sh = client.create('test_table', folder_id="1jp9bDnn225CvC250JadqWft2z4RSP11l")
ws = sh.sheet1

#Заполняем и форматируем первую строку
data = ["Фамилия", "Имя", "Отчество", "Студ. билет", "Профбилет", "Бюджет\контракт", "Направление", "Курс", "Счет", "Факультет", "Адрес", "Категория", "Справки", "Комментарии"]
ws.insert_row(data)
fmt = CellFormat(horizontalAlignment='CENTER',
                 textFormat=textFormat(bold=True, fontSize=12),
                 borders=borders(bottom=border('SOLID'), right=border('SOLID')),
                 padding=padding(bottom=10, right=10))
format_cell_range(ws, 'A1:N1', fmt)
set_row_height(ws, '1', 40)
set_column_width(ws, 'A:N', 200)

#Копирование и сортировка таблицы со студентами
gc = pygsheets.authorize(service_file = 'credentials.json')
base = gc.open_by_key('1xc5GyCMXDV82ssD1ZSyFReiJfuxIpkMigEOh1ojLrCA')

df_base = base[0]
df = df_base.get_as_df()
df2 = df.loc[(df['Статус'] == 'Внести') & (df['База данных'] != 'Ок')]
df2 = df2.sort_values(['ФИО'])
df2_final = df2.iloc[:,[1,3,4,5,6,7,8,9,20,21]]

#Заполняем таблицу на включение данными из отсортированной таблицы со студентами
i = 2
for index, row in df2_final.iterrows():
    insert_data = row['ФИО'].split()
    j = 0
    for elem in row:
        if j == 7:
            insert_data.append('ВМК')
        if j > 0:
            insert_data.append(str(elem))
        j += 1
    ws.update([insert_data], f'A{i}:N{i}')
    i += 1

#-----------------------------------------------------------------------------------------#


#-----------------------------------------------------------------------------------------#
#-----------------------------------------ДОКУМЕНТ----------------------------------------#
#-----------------------------------------------------------------------------------------#
    
# Получение текущей даты
current_date = datetime.now().date()

# Подставьте путь к файлу с вашими учетными данными и ID вашей целевой папки
SERVICE_ACCOUNT_FILE = 'credentials.json'
FOLDER_ID = '1jp9bDnn225CvC250JadqWft2z4RSP11l'

# Аутентификация с использованием учетных данных службы
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']
)

# Создание экземпляра клиента для работы с API Google Docs
docs_service = build('docs', 'v1', credentials=credentials)

# Создание метаданных для нового файла
file_metadata = {
    'name': 'Egor_test_enabling',
    'parents': [FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.document'
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print('File created: %s' % new_file['id'])


# Текст, который нужно вставить в начало документа
text_to_insert = f"""
Список студентов, рекомендованных для включения в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n """

# Создание таблички с данными студентов, которые будем заносить в документ
df = df2_final.iloc[:,[5,1,2]]
insert_data = [[], [], []]
for elem in df2_final['ФИО']:
    tmp = elem.split()
    insert_data[0].append(tmp[0])
    insert_data[1].append(tmp[1])
    insert_data[2].append(tmp[2])
df.insert(loc=0, column='Отчество', value=insert_data[2])
df.insert(loc=0, column='Имя', value=insert_data[1])
df.insert(loc=0, column='Фамилия', value=insert_data[0])
indexes = [str(i) for i in range(1, len(df2_final) + 1)]
df.insert(loc=0, column=' ', value=indexes)
df.to_csv(r'enabling.csv', index=False)

df = pd.read_csv("enabling.csv")
df[' '] = df[' '].astype(str)
df['Курс'] = df['Курс'].astype(str)
df['Номер студенческого билета'] = df['Номер студенческого билета'].astype(str)
df['Номер профсоюзного билета'] = df['Номер профсоюзного билета'].astype(str)
df.fillna('   ', inplace=True)
x, y = df.shape

#Вставка начального текста и создание таблицы
requests = [
    {
        'insertText': {
            'location': {
                'index': 1,
            },
            'text': text_to_insert,
        }
    },
    {
        "insertTable":
        {
            "rows": int(x) + 1,
            "columns": int(y),
            "location":
            {
                "index": len(text_to_insert)
            }
        }
    }]

#Вставка первой строки в таблице
x, y = df.shape
names_column = df.columns.values
print(names_column)
ind = len(text_to_insert) + 4
for i in names_column:
    requests.insert(2, {
        "insertText":
            {
                "text": i,
                "location":
                    {
                        "index": ind
                    }
            }
    })
    ind += 2
ind += 1

#Вставка всех остальных строк
for j in range(0, x):
    for i in names_column:
        requests.insert(2, {
            "insertText":
                {
                    "text": df[i][j],
                    "location":
                        {
                            "index": ind
                        }
                }
        })
        ind += 2
    ind += 1


#Вставка текста после таблицы
text_to_insert = f"""
Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.
"""
requests.insert(2, {
    'insertText': {
        'location': {
            'index': ind,
        },
        'text': text_to_insert,
    }
})

# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(documentId=new_file['id'], body={'requests': requests}).execute()