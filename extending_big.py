# __________ #
import datetime

import pygsheets
import pandas as pd

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
credentials = service_account.Credentials.from_service_account_file(
        'credentials.json', scopes=['https://www.googleapis.com/auth/drive']
    )
drive_service = build("drive", "v3", credentials)

import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials
CREDENTIALS_FILE = 'credentials.json'  # имя файла с закрытым ключом
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                  'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
sheets_service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

gc = pygsheets.authorize(service_file='credentials.json')

# Первый этап (Сбор информации из таблицы ответов на форму)
# Считывание даты, с которой нужно брать данные
date_period_str = input("Введите дату, начиная с которой отсчитывать период: ")
date_period_obj = datetime.datetime.strptime(date_period_str, '%d.%m.%Y')

# Считывание ID таблицы ответов на форму
# table = input()
table = '1VCBrFI6nnEU-omIzAlIuhqDrY74JusRBQ5EDXlAfnkI'
base = gc.open_by_key(table)
df_base = base[0]
df = df_base.get_as_df()

# Отсеиваем только нужный период
df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format = '%d.%m.%Y %H:%M:%S')
df = df.loc[df['Отметка времени'] >= date_period_obj]

# Считали данные таблицы в датафрейм, вытаскиваем только нужные нам (продление), сортируем
df2 = df.loc[df['Статус'] == 'Продлить']

# Если людей на продление нет, завершаем работу
if df2.empty:
    print("Нет людей на продление в этом периоде")
    exit(1)

df2 = df2.sort_values(['ФИО']).reset_index(drop=True)

# Разделили ФИО на разные столбцы, достаём остальные необходимые столбцы
new_df = df2['ФИО'].str.split(expand=True)
new_df.columns=['Фамилия','Имя','Отчество']
df2_final = df2.iloc[:,[3,4,7]]

# Соединяем датафрейм с Фамилией, Именем, Отчеством с остальной информацией
final_df = pd.concat([new_df, df2_final], axis=1)
# Первый этап завершён

# Второй этап сверка с базой ОПК
base = gc.open_by_key('1KjM9rWLjl4T8Vp4DGWnb4pNRTTtwdh12-qdDld5X9jA')
df_base = base[0]
df = df_base.get_as_df()
x, y = final_df.shape
# С помощью базы данных заполняем недостающие данные
for i in range(0, int(x)):
    surname = final_df.iat[i, 0]
    name = final_df.iat[i, 1]
    df2 = df.loc[(df['Фамилия'] == surname) & (df['Имя'] == name)]
    if str(df2.iat[0, 6]) == "м":
        final_df.iat[i, 5] = str(df2.iat[0, 7]) + 'м'
    else:
        final_df.iat[i, 5] = str(df2.iat[0, 7])
    final_df.iat[i, 3] = str(df2.iat[0, 3])
    final_df.iat[i, 4] = str(df2.iat[0, 4])
# Добавляем столбец с нумерацией
indexes = range(1, len(final_df) + 1)
final_df.insert(loc=0, column='X', value=indexes)
final_df['X'] = final_df['X'].astype(str)

# Нужный порядок столбцов
final_df = final_df[['X', 'Фамилия', 'Имя', 'Отчество', 'Курс', 'Номер студенческого билета', 'Номер профсоюзного билета']]

#------------------- Работа с гугл-документом -------------------------------------#

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
start_text_to_insert = f"""
Список студентов, рекомендованных для продления в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n """
#Текст для подписи (нижнего колонтитула)
footer_text = f"""Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Лебедев А. А."""

# Переименование переменной для удобства и замена названия столбцов на нужные
df = final_df
df.rename(columns={'X': '\t'}, inplace=True)
df.rename(columns={'Номер студенческого билета': '№ студенческого билета'}, inplace=True)
df.rename(columns={'Номер профсоюзного билета': '№ профсоюзного билета'}, inplace=True)


# Вставка первоначального текста и первых 20 строк таблицы + устанавливаем размер текста и ширину столбцов
df_temp = df[:20]
x, y = df_temp.shape
start_idx = len(start_text_to_insert)
requests = [
    {
        'insertText': {
            'location': {
                'index': 1,
            },
            'text': start_text_to_insert,
        }
    },
    {
        "insertTable":
        {
            "rows": int(x) + 1,
            "columns": int(y),
            "location":
            {
                "index": start_idx
            }
        }
    }]

# Выделяем начальный текст жирным шрифтом
requests.insert(2, {
    'updateTextStyle': {
            'range': {
                'startIndex': 1,  # Начальный индекс текста
                'endIndex': start_idx + 1  # Конечный индекс текста
            },
            'textStyle': {
                'bold': True
            },
            'fields': '*'
    }
})

# Настраиваем ширину столбцов
widths = [30, 90, 70, 90, 35, 95, 95]
for i in range(y):
    requests.insert(2, {
            'updateTableColumnProperties': {
                'tableStartLocation': {'index': start_idx + 1},  # Индекс начальной строки таблицы (начинается с 1)
                'columnIndices': [i],  # Индексы столбцов, для которых нужно установить ширину
                'tableColumnProperties': {
                    'widthType': 'FIXED_WIDTH',
                    'width': {'magnitude': widths[i], 'unit': 'PT'}
                },  # Установка ширины столбцов в пунктах (PT)
                'fields': '*'
            }
        }
    )

# Запоминаем параметры нашего датафрейма
strs, cols = df.shape
names_column = df.columns.values
# Настройка индекса на начало таблицы
ind = len(start_text_to_insert) + 4

first_ind = ind
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
last_ind = ind - 1

# Выделяем названия столбцов жирным шрифтом
requests.insert(2, {
    'updateTextStyle': {
            'range': {
                'startIndex': first_ind,  # Начальный индекс текста
                'endIndex': last_ind  # Конечный индекс текста
            },
            'textStyle': {
                'bold': True
            },
            'fields': '*'
    }
})

first_ind = ind

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

last_ind = ind

# Меняем размер текста в таблице
requests.insert(2, {
    'updateTextStyle': {
            'range': {
                'startIndex': first_ind,  # Начальный индекс текста
                'endIndex': last_ind  # Конечный индекс текста
            },
            'textStyle': {
                'fontSize': {
                    'magnitude': 10,  # Размер шрифта в пунктах
                    'unit': 'PT'
                }
            },
            'fields': '*'
    }
})

if strs > 20:
    # Вставка текста с подписью
    requests.insert(2, {
        'insertText': {
            'location': {
                'index': ind,
            },
            'text': '\n\n' + footer_text,
        }
    })
elif strs > 7:
    # Вставка текста с подписью
    requests.insert(2, {
        'insertText': {
            'location': {
                'index': ind,
            },
            'text': '\n\n' + footer_text,
        }
    })
    # Вставка разрыва страницы перед финальным текстом
    requests.insert(2, {
        "insertPageBreak": {
            "location": {
                "index": ind
            }
        }
    })


# Вставляем пачками части таблицы
begin = 20
step = 26  # шаг вставки
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
    # Настраиваем ширину столбцов
    for i in range(y):
        requests.insert(put_index + 1, {
                'updateTableColumnProperties': {
                    'tableStartLocation': {'index': ind + 1},  # Индекс начальной строки таблицы (начинается с 1)
                    'columnIndices': [i],  # Индексы столбцов, для которых нужно установить ширину
                    'tableColumnProperties': {
                        'widthType': 'FIXED_WIDTH',
                        'width': {'magnitude': widths[i], 'unit': 'PT'}
                    },  # Установка ширины столбцов в пунктах (PT)
                    'fields': '*'
                }
            }
        )
    # Перемещение на начало таблицы и вставка значений
    ind += 4
    first_ind = ind
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
        last_ind = ind - 1
    # Меняем размер текста в таблице
    requests.insert(put_index + 1, {
        'updateTextStyle': {
                'range': {
                    'startIndex': first_ind,  # Начальный индекс текста
                    'endIndex': last_ind  # Конечный индекс текста
                },
                'textStyle': {
                    'fontSize': {
                        'magnitude': 10,  # Размер шрифта в пунктах
                        'unit': 'PT'
                    }
                },
                'fields': '*'
        }
    })
    # Вставка текста для подписи
    requests.insert(put_index + 1, {
        'insertText': {
            'location': {
                'index': ind-1,
            },
            'text': '\n\n' + footer_text,
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
    # Настраиваем ширину столбцов
    for i in range(y):
        requests.insert(put_index + 1, {
                'updateTableColumnProperties': {
                    'tableStartLocation': {'index': ind + 1},  # Индекс начальной строки таблицы (начинается с 1)
                    'columnIndices': [i],  # Индексы столбцов, для которых нужно установить ширину
                    'tableColumnProperties': {
                        'widthType': 'FIXED_WIDTH',
                        'width': {'magnitude': widths[i], 'unit': 'PT'}
                    },  # Установка ширины столбцов в пунктах (PT)
                    'fields': '*'
                }
            }
        )
    # Перемещение на начало таблицы и вставка значений
    ind += 4
    first_ind = ind
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
    last_ind = ind - 1
    # Меняем размер текста в таблице
    requests.insert(put_index + 1, {
        'updateTextStyle': {
                'range': {
                    'startIndex': first_ind,  # Начальный индекс текста
                    'endIndex': last_ind  # Конечный индекс текста
                },
                'textStyle': {
                    'fontSize': {
                        'magnitude': 10,  # Размер шрифта в пунктах
                        'unit': 'PT'
                    }
                },
                'fields': '*'
        }
    })

    if strs - begin >= 9:
        # Вставка текста для подписи
        requests.insert(put_index + 1, {
            'insertText': {
                'location': {
                    'index': ind - 1,
                },
                'text': '\n\n' + footer_text,
            }
        })
        # Вставка разрыва страницы перед финальным текстом
        requests.insert(put_index + 1, {
            "insertPageBreak": {
                "location": {
                    "index": ind - 1
                }
            }
        })
    # Донастройка индексов для вставки текста после таблицы
    put_index += 1
    ind -= 1

# Текст для вставки в конец документа
end_text_to_insert = f"""
Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.\n
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
        'text': end_text_to_insert,
    }
})


# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(documentId=new_file['id'], body={'requests': requests}).execute()



