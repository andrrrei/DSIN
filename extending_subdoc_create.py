# _________ ОБЯЗАТЕЛЬНАЯ ЧАСТЬ В ЛЮБОЙ ПРОГРАММЕ _________ #
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

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

# Создание экземпляра клиента для работы с API Google Drive и Google Docs
drive_service = build('drive', 'v3', credentials=credentials)
docs_service = build('docs', 'v1', credentials=credentials)

# ________________________________________________________ #


# Создание метаданных для нового файла
file_metadata = {
    'name': 'Продление справок',
    'parents': [FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.document'
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print('File created: %s' % new_file['id'])

months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
# Текст, который нужно вставить в документ
text_to_insert = f"""
Справки для БДНС
Факультет ВМК
\n
 "{current_date.day}" {months[current_date.month - 1]} {current_date.year}г.\n """

# Считывание файла с данными в датафрейм, приведение всех данных в строковый тип
df = pd.read_csv(" my_data.csv")
df.fillna('   ', inplace=True)
df['X'] = df['X'].astype(str)
df.rename(columns={'X': '№'}, inplace=True)
for i in df.columns:
    df[i] = df[i].astype(str)

# Вставка первоначального текста и первых 20 строк таблицы
df_temp = df
x, y = df_temp.shape
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
# Запоминаем параметры нашего датафрейма
strs, cols = df.shape
names_column = df.columns.values
# Настройка индекса на начало таблицы
ind = len(text_to_insert) + 4
# Вставка названий столбцов
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

# Вставка значений
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
# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(documentId=new_file['id'], body={'requests': requests}).execute()

# Функция для настройки ширины столбцов (в разработке)
def modify_table(t_idx):
    """
    t_idx = Table start index.
    """
    request = [{
              'updateTableColumnProperties': {
              'tableStartLocation': {'index': t_idx},
              'columnIndices': [0],
              'tableColumnProperties': {
                'widthType': 'FIXED_WIDTH',
                'width': {
                  'magnitude': 40,
                  'unit': 'PT'
               }
             },
             'fields': '*'
           }
        }
        ]

    result = docs_service.documents().batchUpdate(documentId=new_file['id'], body={'requests': request}).execute()

