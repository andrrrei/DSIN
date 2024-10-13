# __________ #
import pygsheets
import pandas as pd

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


# Считывание ID таблиц начальников курса
# tables = []
# for i in range(6):
#     tables.append(input())


table1 = "1-pZ7CbACZC8a5v2p14Ll8QygtjOQgw9C8fCYAECbzd4"  # 2
# table2 = "1iaL5FvOURtL4SEBfAUv--fIFEdQpb59EioiEse59V7A"  # 5
# table3 = "1LGOd0WiLymp7qcLVxgTrUYIi7G3VUVp8NsElYn9bRNg"  # 3
# table4 = "1v44eeqOp_mc_bnVKbaRsUGWdVnaTfVwZKalakt4FnRY"  # 4
# table5 = "1p8U5muSV43oiyXB222i2uEPIAm9Wc-swPGT0LmBInMU"  # 6
# table6 = '1ZaDMcPsslzJXeWYBCW-vgQnjsO1wrSe43LaC9mx_wG8' # 1

base1 = gc.open_by_key(table1)
# base2 = gc.open_by_key(table2)
# base3 = gc.open_by_key(table3)
# base4 = gc.open_by_key(table4)
# base5 = gc.open_by_key(table5)
# base6 = gc.open_by_key(table6)

df_base1 = base1[0]
df_base1_1 = base1[1]
# df_base2 = base2[0]
# df_base2_1 = base2[1]
# df_base3 = base3[0]
# df_base3_1 = base3[1]
# df_base4 = base4[0]
# df_base4_1 = base4[1]
# df_base5 = base5[0]
# df_base5_1 = base5[1]
# df_base6 = base6[0]
# df_base6_1 = base6[1]

df1 = df_base1.get_as_df()
df1_1 = df_base1_1.get_as_df()
# df2 = df_base2.get_as_df()
# df2_1 = df_base2_1.get_as_df()
# df3 = df_base3.get_as_df()
# df3_1 = df_base3_1.get_as_df()
# df4 = df_base4.get_as_df()
# df4_1 = df_base4_1.get_as_df()
# df5 = df_base5.get_as_df()
# df5_1 = df_base5_1.get_as_df()
# df6 = df_base6.get_as_df()
# df6_1 = df_base6_1.get_as_df()

# Объединение всех датафреймов один
df = pd.concat(
    [
        df1,
        df1_1,
        # df2,
        # df2_1,
        # df3,
        # df3_1,
        # df4,
        # df4_1,
        # df5,
        # df5_1,
        # df6,
        # df6_1,
    ]
)

# Первоначальная подготовка данных: отсеивание отчисленных, сортировка, перестановка столбцов, создание столбца для профсоюзного билета
df2 = df.loc[df["Статус"] != "отчислен"]
df2 = df2.sort_values(["Фамилия"]).reset_index(drop=True)
df2_final = df2.iloc[:, [0, 1, 2, 4, 5, 6]]
df2_final = df2_final[["Фамилия", "Имя", "Отчество", "Курс", "Статус", "Студенческий"]]
df2_final["Номер проф. билета"] = ""

# Первый этап обработки данных из таблиц начальников курса завершён


# Второй этап обработки данных с помощью базы БДНС
base = gc.open_by_key("1Cqa_CERAIpnf3jCPoczB498na8drEMZpDAlUrz9_1cU")
df_base = base[0]
df_base1 = base[1]
df_1 = df_base.get_as_df()
df_2 = df_base1.get_as_df()
df = pd.concat([df_1, df_2])
x, y = df2_final.shape
final_df = df2_final

# С помощью базы данных заполняем недостающие данные
for i in range(0, int(x)):
    if str(final_df.iat[i, 4]) == "закрылся":
        final_df.iat[i, 4] = "да"
    surname = final_df.iat[i, 0]
    name = final_df.iat[i, 1]
    df2 = df.loc[(df["Фамилия"] == surname) & (df["Имя"] == name)]
    if str(df2.iat[0, 8]) == "м":
        final_df.iat[i, 3] = str(df2.iat[0, 9])[0] + "м"
    else:
        final_df.iat[i, 3] = str(df2.iat[0, 9])[0]
    final_df.iat[i, 5] = str(df2.iat[0, 5])
    final_df.iat[i, 6] = str(df2.iat[0, 6])

# Редактируем названия столбцов, добавляем столбец с нумерацией
final_df.rename(
    columns={
        "Статус": "Промежуточная аттестация пройдена",
        "Студенческий": "Номер студ. билета",
    },
    inplace=True,
)
indexes = range(1, len(final_df) + 1)
final_df.insert(loc=0, column="X", value=indexes)
final_df["X"] = final_df["X"].astype(str)

# Второй этап обработки данных завершён, создаём промежуточный CSV-файл с данными (Без промежуточного файла возникает ошибка)
csv_data = final_df.to_csv(r" my_data.csv", index=False)

# ------------------- Работа с гугл-документом -------------------------------------#

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Получение текущей даты
current_date = datetime.now().date()

# Подставьте путь к файлу с вашими учетными данными и ID вашей целевой папки
SERVICE_ACCOUNT_FILE = "credentials.json"

FOLDER_ID = "1dsrZwqADzh3vPCpuhrjuMI0cdgVsQb-V"
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

# Считывание файла с данными в датафрейм, приведение всех данных в строковый тип
df = pd.read_csv(" my_data.csv")
df["X"] = df["X"].astype(str)
df.rename(columns={"X": "\t"}, inplace=True)
for i in df.columns:
    df[i] = df[i].astype(str)
df.fillna(" \t  ", inplace=True)

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
widths = [30, 83, 65, 90, 40, 75, 60, 55]
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
end_text_to_insert = f"""
Список утверждён решением студенческой комиссии профкома ф-та ВМК МГУ от «{current_date.strftime("%d.%m.%y")} г.»
Подтверждаем, что все вышеуказанные студенты обучаются по дневной очной форме обучения за счет средств федерального бюджета РФ.\n
Декан ф-та ВМК МГУ  			       ___________________ Соколов И. А. 


							М.П.

Зам. декана по учебной работе
ф-та ВМК МГУ                                                    ___________________ Федотов М. В.



Председатель профкома ф-та ВМК МГУ          ___________________ Гуров С. И.

							М.П.

Председатель студенческой комиссии 
ф-та ВМК МГУ 				       ___________________ Худяков Д. В.


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
