# _________ ОБЯЗАТЕЛЬНАЯ ЧАСТЬ В ЛЮБОЙ ПРОГРАММЕ _________ #
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
    'name': 'New Google Document',
    'parents': [FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.document'
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print('File created: %s' % new_file['id'])


# Текст, который нужно вставить в документ
text_to_insert = f"""
Список студентов, рекомендованных для продления в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n
Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.
"""

# Запрос на вставку текста в начало документа
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
        'insertTable': {
            'rows': 3,
            'columns': 3,
            'endOfSegmentLocation': {
                'segmentId': ''
            }
        }
    }
]

# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(documentId=new_file['id'], body={'requests': requests}).execute()
