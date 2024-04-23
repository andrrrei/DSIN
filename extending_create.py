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
    'name': 'Продление',
    'parents': [FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.document'
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print('File created: %s' % new_file['id'])


# Текст, который нужно вставить в документ
text_to_insert = f"""
Список студентов, рекомендованных для продления в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n """
text_to_insert2 = f"""
Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.
"""
#Текст для подписи (нижнего колонтитула)
footer_text = f"""Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Лебедев А. А."""

# Считывание файла с данными в датафрейм, приведение всех данных в строковый тип
df = pd.read_csv(" my_data.csv")
df['X'] = df['X'].astype(str)
df.rename(columns={'X': '\t'}, inplace=True)
for i in df.columns:
    df[i] = df[i].astype(str)
df.fillna(' \t  ', inplace=True)

# Вставка первоначального текста и первых 20 строк таблицы
df_temp = df[:20]
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
requests.insert(2, {
    'insertText': {
        'location': {
            'index': ind,
        },
        'text': footer_text,
    }
})

# Вставляем пачками части таблицы
begin = 20
step = 35 # шаг вставки
end = begin + step
put_index = 2

# Цикл со вставкой части таблицы + текста для подписей
while end <= strs:
    requests.insert(put_index, {
        "insertTable":
        {
            "rows": int(step),
            "columns": int(y),
            "location":
            {
                "index": ind
            }
        }})
    ind += 4 # Перемещение на начало таблицы, для вставки значений
    for j in range(begin, end):
        for i in names_column:
            requests.insert(put_index + 1, {
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
    # Вставка текста для подписи
    requests.insert(put_index + 1, {
        'insertText': {
            'location': {
                'index': ind-1,
            },
            'text': footer_text,
        }
    })
    put_index += 2
    ind += len(footer_text) - 1
    begin += step
    end += step

# Проверка на оставшуюся часть датафрейма и её вставка
if begin < strs:
    requests.insert(put_index, {
        "insertTable":
            {
                "rows": int(strs - begin),
                "columns": int(y),
                "location":
                    {
                        "index": ind
                    }
            }})
    ind += 4 # Перемещение на начало таблицы, для вставки значений
    for j in range(begin, strs):
        for i in names_column:
            requests.insert(put_index + 1, {
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
    requests.insert(put_index + 1, {
        'insertText': {
            'location': {
                'index': ind - 1,
            },
            'text': footer_text,
        }
    })
    # Донастройка индексов для вставки текста после таблицы
    put_index += 1
    ind += -1

# Текст для вставки после таблицы
text_to_insert2 = f"""


Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.
Декан ф-та ВМК МГУ  			       ___________________ Соколов И. А. 


							М.П.

Зам. декана по учебной работе
ф-та ВМК МГУ                                                    ___________________ Федотов М. В.



Председатель профкома ф-та ВМК МГУ          ___________________ Поставничая В. А

							М.П.

Председатель студенческой комиссии 
ф-та ВМК МГУ 				       ___________________ Худяков Д. В.


Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Лебедев А. А.
"""
requests.insert(put_index, {
    'insertText': {
        'location': {
            'index': ind,
        },
        'text': text_to_insert2,
    }
})

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

