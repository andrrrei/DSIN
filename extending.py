# __________ #
import os
import sys

# --------------DEBUG------------#
# path = 'scripts/debug/extending.txt'
# sys.stderr = open(path, 'w')

import pygsheets
import pandas as pd
import gspread
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

credentials = service_account.Credentials.from_service_account_file(
    "credentials.json", scopes=["https://www.googleapis.com/auth/drive"]
)
drive_service = build("drive", "v3", credentials)

import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = "credentials.json"  # имя файла с закрытым ключом
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ],
)
httpAuth = credentials.authorize(httplib2.Http())
sheets_service = apiclient.discovery.build("sheets", "v4", http=httpAuth)

gc = pygsheets.authorize(service_file="credentials.json")

# ИЗМЕНЯЕМАЯ ИНФОРМАЦИЯ
# Folder ID
ROOT_FOLDER_ID = "10oXbgb7tBzFp41KpfXwmfWbepmnhZ9ch"
# ID базы данных БДНС
status_table = "1Cqa_CERAIpnf3jCPoczB498na8drEMZpDAlUrz9_1cU"
# Считывание ID таблицы ответов на форму
table = "1fZhfUDWSGGr6uHQVdMpA1O2KNX32uXpKe8hMMNkoeMM"

# -----------------------------------------------------------#

# ----------------------РАБОТА С ПАПКАМИ----------------------------#

# ----------------------АУТЕНТИФИКАЦИЯ------------------------------#
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
drive_service = build("drive", "v3", credentials=credentials)
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(credentials)


def find_folder(service, folder_name, parent_folder_id=0):
    if parent_folder_id == 0:
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    else:
        query = f"name = '{folder_name}' and '{parent_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get("files", [])
    return items


def create_folder(service, folder_name, parent_folder_id=0):
    if parent_folder_id == 0:
        file_metadata = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
        }
    else:
        file_metadata = {
            "name": folder_name,
            "parents": [parent_folder_id],
            "mimeType": "application/vnd.google-apps.folder",
        }
    file = service.files().create(body=file_metadata, fields="id").execute()
    return file.get("id")


# Поиск папки "УГ" в корневой папке
current_year = datetime.now().year
if datetime.now().month >= 9:
    # Если сейчас сентябрь или позже, то учебный год начинается в этом году
    start_year = current_year
else:
    # Иначе учебный год начался в предыдущем году
    start_year = current_year - 1
end_year = start_year + 1
folder_name = f"УГ {start_year}-{end_year}"

folders = find_folder(drive_service, folder_name, ROOT_FOLDER_ID)
if not folders:
    folder_id = create_folder(drive_service, folder_name, ROOT_FOLDER_ID)
else:
    folder_id = folders[0]["id"]

# Поиск папки текущего месяца в папке УГ
months = {
    "January": "Январь",
    "February": "Февраль",
    "March": "Март",
    "April": "Апрель",
    "May": "Май",
    "June": "Июнь",
    "July": "Июль",
    "August": "Август",
    "September": "Сентябрь",
    "October": "Октябрь",
    "November": "Ноябрь",
    "December": "Декабрь",
}
current_month = datetime.now().strftime("%B")
current_month = months[current_month]
folders = find_folder(drive_service, current_month, folder_id)
if not folders:
    folder_id = create_folder(drive_service, current_month, folder_id)
else:
    folder_id = folders[0]["id"]

# Поиск папки "Продление" в папке текущего месяца
folder_name = "Продление"
folders = find_folder(drive_service, folder_name, folder_id)
if not folders:
    folder_id = create_folder(drive_service, folder_name, folder_id)
else:
    folder_id = folders[0]["id"]

FOLDER_ID = folder_id

# ---------НУЖНЫЕ ПАПКИ НАЙДЕНЫ ИЛИ СОЗДАНЫ----------------

# Первый этап (Сбор информации из таблицы ответов на форму)
inter_file = " my_data.csv"  # Временный файл
base = gc.open_by_key(table)
df_base = base[0]
df = df_base.get_as_df()

# Отсеиваем только нужный период
df = df.loc[df["База данных"] != "Ок"]

# Считали данные таблицы в датафрейм, вытаскиваем только нужные нам (продление), сортируем
df2 = df.loc[df["Статус"] == "Продлить"]

# Если людей на продление нет, завершаем работу
if df2.empty:
    print("Нет людей на продление в этом периоде")
    exit(1)

# Проставляем ОК
# Установка доступа к Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    "credentials.json", scope
)
gc2 = gspread.authorize(credentials)
answer_base = gc2.open_by_key(table).sheet1
# Получаем все данные из таблицы в виде списка списков
data = answer_base.get_all_values()

# Создаем DataFrame из полученных данных
df = pd.DataFrame(data[1:], columns=data[0])

x, y = df2.shape
for i in range(0, x):
    # Находим индекс строки, в которой нужно изменить значение
    value = df2.iat[i, 1]
    index = df[df["ФИО"] == str(value)].index.tolist()
    # Меняем значение в другом столбце
    if index:
        index = (
            index[0] + 2
        )  # Прибавляем 2, потому что индексация в Google Sheets начинается с 1 и нумерация строк считается с 2
        answer_base.update_cell(index, 26, "Ок")
        print("Ок!")
    else:
        print("Не Ок")

df2 = df2.sort_values(["ФИО"]).reset_index(drop=True)
new_df = df2["ФИО"].str.split(expand=True)
new_df.columns = ["Фамилия", "Имя", "Отчество"]

# Информация для документа о продлении
df2_final = df2.iloc[:, [3, 4, 7, 21, 22, 16, 17, 18, 19]]


# Соединяем датафрейм с Фамилией, Именем, Отчеством с остальной информацией
final_df = pd.concat([new_df, df2_final], axis=1)
# Первый этап завершён

# Второй этап сверка с базой БДНС
base = gc.open_by_key(status_table)
df_base = base[0]
df = df_base.get_as_df()
x, y = final_df.shape
# С помощью базы данных заполняем недостающие данные
for i in range(0, int(x)):
    surname = final_df.iat[i, 0]
    name = final_df.iat[i, 1]
    df2 = df.loc[(df["Фамилия"] == surname) & (df["Имя"] == name)]
    if str(df2.iat[0, 8]) == "м":
        final_df.iat[i, 5] = str(df2.iat[0, 9]) + "м"
    else:
        final_df.iat[i, 5] = str(df2.iat[0, 9])
    final_df.iat[i, 3] = str(df2.iat[0, 5])
    final_df.iat[i, 4] = str(df2.iat[0, 6])


# Обновление сроков истечение документов в таблице БДНС
# Подключаемся к базе статусов БДНС
# Установка доступа к Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    "credentials.json", scope
)
gc2 = gspread.authorize(credentials)
status_base = gc2.open_by_key(status_table).sheet1

# Получаем все данные из таблицы в виде списка списков
data = status_base.get_all_values()

# Создаем DataFrame из полученных данных
df = pd.DataFrame(data[1:], columns=data[0])

# Извлекаем номер студенческого и срок действия документов
x, y = final_df.shape
for i in range(0, x):
    # Находим индекс строки, в которой нужно изменить значение
    value = final_df.iat[i, 3]
    index = df[df["Студенческий"] == str(value)].index.tolist()
    # Меняем значение в другом столбце
    if index:
        index = (
            index[0] + 2
        )  # Прибавляем 2, потому что индексация в Google Sheets начинается с 1 и нумерация строк считается с 2
        status_base.update_cell(index, 20, final_df.iat[i, 7])
        print("Значение обновлено успешно!")
    else:
        print("Значение для обновления данных не найдено.")


# Отбираем информацию для документа со справками
df3 = final_df.copy()
sub_df = new_df.copy()
x, y = sub_df.shape
for i in range(0, int(x)):
    sub_df.iat[i, 0] = (
        str(sub_df.iat[i, 0]) + str(sub_df.iat[i, 1])[0] + str(sub_df.iat[i, 2])[0]
    )

# Переименование столбцов и подготовка записи датафрейма в CSV-файл
sub_df = sub_df["Фамилия"]
df3 = df3.iloc[:, [6, 7, 8, 9, 10, 11]]
final_sub_df = pd.concat([sub_df, df3], axis=1)
final_sub_df.rename(columns={"Тип справки": "Документы"}, inplace=True)
final_sub_df.rename(columns={"Фамилия": "Название файла"}, inplace=True)
indexes = range(1, len(final_sub_df) + 1)
final_sub_df.insert(loc=0, column="X", value=indexes)
final_sub_df["X"] = final_sub_df["X"].astype(str)

# Скопируем файлы каждого студента в отдельную папку на диске

# ----------------------КОПИРОВАНИЕ ФАЙЛОВ-------------------------------#

# Укажите путь к вашему JSON файлу учетной записи службы
SERVICE_ACCOUNT_FILE = "credentials.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]


# -----------Описание вспомогательных функций---------------------#
def authenticate():
    """Авторизация и получение service."""
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build("drive", "v3", credentials=creds)
    return service


def get_folder_id(service, folder_name, parent_folder_id):
    """Проверяет наличие папки на Google Диске и возвращает её ID."""
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and '{parent_folder_id}' in parents and trashed=false"
    results = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name)")
        .execute()
    )
    items = results.get("files", [])
    if items:
        # Папка найдена, возвращаем её ID
        return items[0]["id"]
    return None


def create_folder(service, folder_name, parent_folder_id):
    """Создает папку на Google Диске, если её ещё нет."""
    folder_id = get_folder_id(service, folder_name, parent_folder_id)
    if folder_id:
        print(f"Папка '{folder_name}' уже существует с ID '{folder_id}'.")
        return folder_id
    file_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }
    folder = service.files().create(body=file_metadata, fields="id").execute()
    print(f"Папка '{folder_name}' создана внутри '{parent_folder_id}'.")
    return folder.get("id")


def copy_file(service, file_id, destination_folder_id):
    """Копирует файл на Google Диске в другую папку."""
    copied_file = {"parents": [destination_folder_id]}
    try:
        copied_file = service.files().copy(fileId=file_id, body=copied_file).execute()
        print(
            f"Файл с ID {file_id} успешно скопирован в папку с ID {destination_folder_id}."
        )
        return copied_file.get("id")
    except Exception as e:
        print(f"Не удалось скопировать файл: {e}")


def get_file_id_from_url(file_url):
    """Получает идентификатор файла из ссылки на файл Google Диска."""
    # Ссылка формата https://drive.google.com/open?id=++++++++++++++++++++++
    pos = file_url.rfind("id=")  # находит последнее вхождение 'id='
    return file_url[pos + len("id=") :]


def rename_file(file_id, new_filename, service):
    file_metadata = {"name": new_filename}
    # Выполняем запрос к Google Drive API для обновления метаданных файла
    try:
        file = service.files().update(fileId=file_id, body=file_metadata).execute()
    except Exception as e:
        print(f"Произошла ошибка при изменении имени файла: {e}")


def copy_files_of_user(parent_folder_id, file_url, new_folder_name):
    service = authenticate()

    # Получение идентификатора файла из ссылки
    file_id = get_file_id_from_url(file_url)

    # Создание новой папки
    new_folder_id = create_folder(service, new_folder_name, parent_folder_id)

    # Копирование файла в новую папку
    copy_file(service, file_id, new_folder_id)

    rename_file(file_id, "Подтверждающий документ.pdf", service)


# ------------------------------------------------------------------#


# Запуск для каждого из студентов
# Для каждого студента создаётся папка с его документами
x, y = final_sub_df.shape
for i in range(0, int(x)):
    for j in range(4, 8):
        elem = final_sub_df.iat[i, j]
        if not (pd.isnull(elem) or (elem == "")):
            copy_files_of_user(FOLDER_ID, elem, final_sub_df.iat[i, 1])

# --------------------РАБОТА С ПАПКАМИ ЗАВЕРШЕНА------------------------------#

# Оставим только нужную информацию для документа "Справки"
final_sub_df = final_sub_df.iloc[:, [0, 1, 2, 3]]

# Второй этап обработки данных завершён, создаём промежуточный CSV-файл с данными (Без промежуточного файла возникает ошибка)
csv_data = final_sub_df.to_csv(inter_file, index=False)

# ------------------------------------------------------------

# Далее следует работа с документом о продлении
final_df = final_df.iloc[:, [0, 1, 2, 3, 4, 5]]

# Добавляем столбец с нумерацией
indexes = range(1, len(final_df) + 1)
final_df.insert(loc=0, column="X", value=indexes)
final_df["X"] = final_df["X"].astype(str)

# Нужный порядок столбцов
final_df = final_df[
    [
        "X",
        "Фамилия",
        "Имя",
        "Отчество",
        "Курс",
        "Номер студенческого билета",
        "Номер профсоюзного билета",
    ]
]

# ------------------- Работа с гугл-документом -------------------------------------#

from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Получение текущей даты
current_date = datetime.now().date()

# Подставьте путь к файлу с вашими учетными данными и ID вашей целевой папки
SERVICE_ACCOUNT_FILE = "credentials.json"


# Аутентификация с использованием учетных данных службы
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=[
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents",
    ],
)

# Создание экземпляра клиента для работы с API Google Drive и Google Docs
drive_service = build("drive", "v3", credentials=credentials)
docs_service = build("docs", "v1", credentials=credentials)

# ________________________________________________________ #


# Создание метаданных для нового файла
file_metadata = {
    "name": "Продление",
    "parents": [FOLDER_ID],
    "mimeType": "application/vnd.google-apps.document",
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print("File created: %s" % new_file["id"])


# Текст, который нужно вставить в документ
start_text_to_insert = f"""
Список студентов, рекомендованных для продления в БДНС факультета ВМК МГУ.\n
{'Осенний' if current_date.month in [1, 8, 9, 10, 11, 12] else 'Весенний'} семестр {current_date.year} г.\n """
# Текст для подписи (нижнего колонтитула)
footer_text = f"""Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Гудов Д. О."""

# Переименование переменной для удобства и замена названия столбцов на нужные
df = final_df
df.rename(columns={"X": "\t"}, inplace=True)
df.rename(
    columns={"Номер студенческого билета": "№ студенческого билета"}, inplace=True
)
df.rename(columns={"Номер профсоюзного билета": "№ профсоюзного билета"}, inplace=True)


# Вставка первоначального текста и первых 20 строк таблицы + устанавливаем размер текста и ширину столбцов
df_temp = df[:20]
x, y = df_temp.shape
start_idx = len(start_text_to_insert)
requests = [
    {
        "insertText": {
            "location": {
                "index": 1,
            },
            "text": start_text_to_insert,
        }
    },
    {
        "insertTable": {
            "rows": int(x) + 1,
            "columns": int(y),
            "location": {"index": start_idx},
        }
    },
]

# Выделяем начальный текст жирным шрифтом
requests.insert(
    2,
    {
        "updateTextStyle": {
            "range": {
                "startIndex": 1,  # Начальный индекс текста
                "endIndex": start_idx + 1,  # Конечный индекс текста
            },
            "textStyle": {"bold": True},
            "fields": "*",
        }
    },
)

# Настраиваем ширину столбцов
widths = [27, 90, 70, 90, 37, 95, 95]
for i in range(y):
    requests.insert(
        2,
        {
            "updateTableColumnProperties": {
                "tableStartLocation": {
                    "index": start_idx + 1
                },  # Индекс начальной строки таблицы (начинается с 1)
                "columnIndices": [
                    i
                ],  # Индексы столбцов, для которых нужно установить ширину
                "tableColumnProperties": {
                    "widthType": "FIXED_WIDTH",
                    "width": {"magnitude": widths[i], "unit": "PT"},
                },  # Установка ширины столбцов в пунктах (PT)
                "fields": "*",
            }
        },
    )

# Запоминаем параметры нашего датафрейма
strs, cols = df.shape
names_column = df.columns.values
# Настройка индекса на начало таблицы
ind = len(start_text_to_insert) + 4

first_ind = ind
# Вставка названий столбцов
for i in names_column:
    requests.insert(2, {"insertText": {"text": i, "location": {"index": ind}}})
    ind += 2
ind += 1
last_ind = ind - 1

# Выделяем названия столбцов жирным шрифтом
requests.insert(
    2,
    {
        "updateTextStyle": {
            "range": {
                "startIndex": first_ind,  # Начальный индекс текста
                "endIndex": last_ind,  # Конечный индекс текста
            },
            "textStyle": {"bold": True},
            "fields": "*",
        }
    },
)

first_ind = ind

# Вставка значений
for j in range(0, x):
    for i in names_column:
        requests.insert(
            2, {"insertText": {"text": df[i][j], "location": {"index": ind}}}
        )
        ind += 2
    ind += 1

last_ind = ind

# Меняем размер текста в таблице
requests.insert(
    2,
    {
        "updateTextStyle": {
            "range": {
                "startIndex": first_ind,  # Начальный индекс текста
                "endIndex": last_ind,  # Конечный индекс текста
            },
            "textStyle": {
                "fontSize": {"magnitude": 10, "unit": "PT"}  # Размер шрифта в пунктах
            },
            "fields": "*",
        }
    },
)

if strs > 20:
    # Вставка текста с подписью
    requests.insert(
        2,
        {
            "insertText": {
                "location": {
                    "index": ind,
                },
                "text": "\n\n" + footer_text,
            }
        },
    )
elif strs > 7:
    # Вставка текста с подписью
    requests.insert(
        2,
        {
            "insertText": {
                "location": {
                    "index": ind,
                },
                "text": "\n\n" + footer_text,
            }
        },
    )
    # Вставка разрыва страницы перед финальным текстом
    requests.insert(2, {"insertPageBreak": {"location": {"index": ind}}})


# Вставляем пачками части таблицы
begin = 20
step = 26  # шаг вставки
end = begin + step
put_index = 2

# Цикл со вставкой части таблицы + текста для подписей
while end <= strs:
    requests.insert(
        put_index,
        {
            "insertTable": {
                "rows": int(step),
                "columns": int(y),
                "location": {"index": ind},
            }
        },
    )
    # Настраиваем ширину столбцов
    for i in range(y):
        requests.insert(
            put_index + 1,
            {
                "updateTableColumnProperties": {
                    "tableStartLocation": {
                        "index": ind + 1
                    },  # Индекс начальной строки таблицы (начинается с 1)
                    "columnIndices": [
                        i
                    ],  # Индексы столбцов, для которых нужно установить ширину
                    "tableColumnProperties": {
                        "widthType": "FIXED_WIDTH",
                        "width": {"magnitude": widths[i], "unit": "PT"},
                    },  # Установка ширины столбцов в пунктах (PT)
                    "fields": "*",
                }
            },
        )
    # Перемещение на начало таблицы и вставка значений
    ind += 4
    first_ind = ind
    for j in range(begin, end):
        for i in names_column:
            requests.insert(
                put_index + 1,
                {"insertText": {"text": df[i][j], "location": {"index": ind}}},
            )
            ind += 2
        ind += 1
        last_ind = ind - 1
    # Меняем размер текста в таблице
    requests.insert(
        put_index + 1,
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": first_ind,  # Начальный индекс текста
                    "endIndex": last_ind,  # Конечный индекс текста
                },
                "textStyle": {
                    "fontSize": {
                        "magnitude": 10,  # Размер шрифта в пунктах
                        "unit": "PT",
                    }
                },
                "fields": "*",
            }
        },
    )
    # Вставка текста для подписи
    requests.insert(
        put_index + 1,
        {
            "insertText": {
                "location": {
                    "index": ind - 1,
                },
                "text": "\n\n" + footer_text,
            }
        },
    )
    put_index += 2
    ind += len(footer_text) - 1
    begin += step
    end += step

# Проверка на оставшуюся часть датафрейма и её вставка
if begin < strs:
    requests.insert(
        put_index,
        {
            "insertTable": {
                "rows": int(strs - begin),
                "columns": int(y),
                "location": {"index": ind},
            }
        },
    )
    # Настраиваем ширину столбцов
    for i in range(y):
        requests.insert(
            put_index + 1,
            {
                "updateTableColumnProperties": {
                    "tableStartLocation": {
                        "index": ind + 1
                    },  # Индекс начальной строки таблицы (начинается с 1)
                    "columnIndices": [
                        i
                    ],  # Индексы столбцов, для которых нужно установить ширину
                    "tableColumnProperties": {
                        "widthType": "FIXED_WIDTH",
                        "width": {"magnitude": widths[i], "unit": "PT"},
                    },  # Установка ширины столбцов в пунктах (PT)
                    "fields": "*",
                }
            },
        )
    # Перемещение на начало таблицы и вставка значений
    ind += 4
    first_ind = ind
    for j in range(begin, strs):
        for i in names_column:
            requests.insert(
                put_index + 1,
                {"insertText": {"text": df[i][j], "location": {"index": ind}}},
            )
            ind += 2
        ind += 1
    last_ind = ind - 1
    # Меняем размер текста в таблице
    requests.insert(
        put_index + 1,
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": first_ind,  # Начальный индекс текста
                    "endIndex": last_ind,  # Конечный индекс текста
                },
                "textStyle": {
                    "fontSize": {
                        "magnitude": 10,  # Размер шрифта в пунктах
                        "unit": "PT",
                    }
                },
                "fields": "*",
            }
        },
    )

    if strs - begin >= 9:
        # Вставка текста для подписи
        requests.insert(
            put_index + 1,
            {
                "insertText": {
                    "location": {
                        "index": ind - 1,
                    },
                    "text": "\n\n" + footer_text,
                }
            },
        )
        # Вставка разрыва страницы перед финальным текстом
        requests.insert(
            put_index + 1, {"insertPageBreak": {"location": {"index": ind - 1}}}
        )
    # Донастройка индексов для вставки текста после таблицы
    put_index += 1
    ind -= 1

# Текст для вставки в конец документа
end_text_to_insert = f"""Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»\n
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.\n
Декан ф-та ВМК МГУ  			       ___________________ Соколов И. А. 


							М.П.

Зам. декана по учебной работе
ф-та ВМК МГУ                                                    ___________________ Федотов М. В.



Председатель профкома ф-та ВМК МГУ          ___________________ Гуров С. И.

							М.П.

Председатель студенческой комиссии 
ф-та ВМК МГУ 				       ___________________ Беппиев Г. И.


Ответственный за ведение БДНС 
ф-та ВМК МГУ  				       ___________________ Гудов Д. О.
"""

requests.insert(
    put_index,
    {
        "insertText": {
            "location": {
                "index": ind,
            },
            "text": end_text_to_insert,
        }
    },
)


# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(
    documentId=new_file["id"], body={"requests": requests}
).execute()

# Работа с документом со справками

# ------------------- Работа с гугл-документом -------------------------------------#
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Аутентификация с использованием учетных данных службы
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=[
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents",
    ],
)

# Создание экземпляра клиента для работы с API Google Drive и Google Docs
drive_service = build("drive", "v3", credentials=credentials)
docs_service = build("docs", "v1", credentials=credentials)

# ________________________________________________________ #


# Создание метаданных для нового файла
file_metadata = {
    "name": "Продление справок",
    "parents": [FOLDER_ID],
    "mimeType": "application/vnd.google-apps.document",
}

# Создание пустого текстового файла
new_file = drive_service.files().create(body=file_metadata).execute()

print("File created: %s" % new_file["id"])

months = [
    "январь",
    "февраль",
    "март",
    "апрель",
    "май",
    "июнь",
    "июль",
    "август",
    "сентябрь",
    "октябрь",
    "ноябрь",
    "декабрь",
]
# Текст, который нужно вставить в документ
text_to_insert = f"""
Справки для БДНС
Факультет ВМК
 "{current_date.day}" {months[current_date.month - 1]} {current_date.year}г.\n """

# Считывание файла с данными в датафрейм, приведение всех данных в строковый тип
df = pd.read_csv(inter_file)
os.remove(inter_file)
df.fillna("   ", inplace=True)
df["X"] = df["X"].astype(str)
df.rename(columns={"X": "№"}, inplace=True)
for i in df.columns:
    df[i] = df[i].astype(str)

# Вставка первоначального текста и первых 20 строк таблицы
df_temp = df
x, y = df_temp.shape
start_idx = len(text_to_insert)
requests = [
    {
        "insertText": {
            "location": {
                "index": 1,
            },
            "text": text_to_insert,
        }
    },
    {
        "updateTextStyle": {
            "textStyle": {"fontSize": {"magnitude": 18, "unit": "PT"}},  # Размер шрифта
            "fields": "fontSize",
            "range": {"startIndex": 1, "endIndex": start_idx},
        }
    },
    {
        "updateParagraphStyle": {
            "paragraphStyle": {"alignment": "CENTER"},  # Выравнивание по центру
            "fields": "alignment",
            "range": {"startIndex": 1, "endIndex": start_idx + 1},
        }
    },
    {
        "insertTable": {
            "rows": int(x) + 1,
            "columns": int(y),
            "location": {"index": start_idx},
        }
    },
]


widths = [40, 120, 145, 120]
for i in range(y):
    requests.insert(
        4,
        {
            "updateTableColumnProperties": {
                "tableStartLocation": {
                    "index": start_idx + 1
                },  # Индекс начальной строки таблицы (начинается с 1)
                "columnIndices": [
                    i
                ],  # Индексы столбцов, для которых нужно установить ширину
                "tableColumnProperties": {
                    "widthType": "FIXED_WIDTH",
                    "width": {"magnitude": widths[i], "unit": "PT"},
                },  # Установка ширины столбцов в пунктах (PT)
                "fields": "*",
            }
        },
    )

# Запоминаем параметры нашего датафрейма
strs, cols = df.shape
names_column = df.columns.values
# Настройка индекса на начало таблицы
ind = len(text_to_insert) + 4

first_ind = ind
# Вставка названий столбцов
for i in names_column:
    requests.insert(4, {"insertText": {"text": i, "location": {"index": ind}}})
    ind += 2
ind += 1
last_ind = ind - 1

requests.insert(
    4,
    {
        "updateTextStyle": {
            "range": {
                "startIndex": first_ind,  # Начальный индекс текста
                "endIndex": last_ind,  # Конечный индекс текста
            },
            "textStyle": {"bold": True},
            "fields": "*",
        }
    },
)


# Вставка значений
for j in range(0, x):
    for i in names_column:
        requests.insert(
            4, {"insertText": {"text": df[i][j], "location": {"index": ind}}}
        )
        ind += 2
    ind += 1
# Отправка запроса на обновление документа
docs_service.documents().batchUpdate(
    documentId=new_file["id"], body={"requests": requests}
).execute()
