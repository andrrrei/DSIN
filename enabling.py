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
import os

SERVICE_ACCOUNT_FILE = 'credentials.json'
PARENT_FOLDER_ID = '14blDt4JalikhJUoenyKLUeT0dxj-sq7G'
FORM_RESPONSES_ID = '1fZhfUDWSGGr6uHQVdMpA1O2KNX32uXpKe8hMMNkoeMM'
#-----------------------------------------------------------------------------------------#
#--------------------------------------ТАБЛИЦА--------------------------------------------#
#-----------------------------------------------------------------------------------------#

# Аутентификация
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
drive_service = build('drive', 'v3', credentials=credentials)
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
client = gspread.authorize(credentials)

# Создание таблицы
sh = client.create('Таблица включения', folder_id=PARENT_FOLDER_ID)
ws = sh.sheet1

# Заполняем и форматируем первую строку
data = ["Фамилия", "Имя", "Отчество", "Студ. билет", "Профбилет", "Бюджет\контракт", "Направление", "Курс", "Счет", "Факультет", "Адрес", "Категория", "Справки", "Комментарии"]
ws.insert_row(data)
fmt = CellFormat(horizontalAlignment='CENTER',
                 textFormat=textFormat(bold=True, fontSize=12),
                 borders=borders(bottom=border('SOLID'), right=border('SOLID')),
                 padding=padding(bottom=10, right=10))
format_cell_range(ws, 'A1:N1', fmt)
set_row_height(ws, '1', 40)
set_column_width(ws, 'A:N', 200)

# Копирование и сортировка таблицы со студентами
gc = pygsheets.authorize(service_file = SERVICE_ACCOUNT_FILE)
base = gc.open_by_key(FORM_RESPONSES_ID)

df_base = base[0]
df = df_base.get_as_df()
df2 = df.loc[(df['Статус'] == 'Внести') & (df['База данных'] != 'Ок')]
df2 = df2.sort_values(['ФИО'])
df2_final = df2.iloc[:,[1, 3, 4, 5, 6, 7, 8, 9, 20, 22]]

# Заполняем таблицу на включение данными из отсортированной таблицы со студентами
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
#-------------------------Вставляем "Ок" в пустые строки в таблице------------------------#
#-----------------------------------------------------------------------------------------#

# x, y = df.shape
# sheet = base.sheet1
# for i in range(x):
#     if df['Статус'][i] == 'Внести' and df['База данных'][i] != 'Ок':
#         sheet.update_cell(i + 2, 26, 'Ок')

#-----------------------------------------------------------------------------------------#


#-----------------------------------------------------------------------------------------#
#-----------------------------------------ДОКУМЕНТ----------------------------------------#
#-----------------------------------------------------------------------------------------#
    
# Получение текущей даты
current_date = datetime.now().date()

# Аутентификация с использованием учетных данных службы
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']
)

# Создание экземпляра клиента для работы с API Google Docs
docs_service = build('docs', 'v1', credentials=credentials)

# Создание метаданных для нового файла
file_metadata = {
    'name': 'Включение',
    'parents': [PARENT_FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.document'
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()
doc_id = new_file['id']
print('File created: %s' % new_file['id'])


# Текст, который нужно вставить в начало документа
start_text_to_insert = f"""Список студентов, рекомендованных для включения в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n """

# Текст для подписи (нижнего колонтитула)
footer_text = f"""Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Лебедев А. А."""


# Создание таблички с данными студентов, которые будем заносить в документ
df = df2_final.iloc[:,[5, 1, 2]]
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

df.to_csv(r'my_data.csv', index=False)
df = pd.read_csv("my_data.csv")
os.remove("my_data.csv")


# Убираем числа типа float и приводим все к строкам + убираем Nan
# def process_float(x):
#     if isinstance(x, float):
#         return str(int(x))
#     else:
#         return x
# df = df.map(lambda x: process_float(x)).astype(str)

df.fillna(' ', inplace=True)
for i in df.columns:
    df[i] = df[i].astype(str)
print(df)

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
widths = [30, 100, 70, 100, 40, 90, 95]
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