# Подключение необходимого

from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
import gspread
import pandas as pd
import re

SERVICE_ACCOUNT_FILE = "credentials.json"

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=[
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents",
    ],
)

# Создание таблицы из базы

gc = gspread.authorize(credentials)

ws_bdns = gc.open_by_key("1Cqa_CERAIpnf3jCPoczB498na8drEMZpDAlUrz9_1cU")  # БДНС ВМК
data1 = pd.DataFrame(ws_bdns.get_worksheet(0).get_all_records())  # Бюджет ЧП
data2 = pd.DataFrame(ws_bdns.get_worksheet(1).get_all_records())  # Контракт ЧП
data_bdns = pd.concat([data1, data2], ignore_index=True)  # объединение листов


data_bdns = data_bdns[["Студенческий", "Курс", "Срок действия", "Статус"]]
data_bdns = data_bdns.rename(
    columns={
        "Студенческий": "Номер студенческого",
        "Срок действия": "Истечение документов",
    }
)
data_bdns["Статус"] = data_bdns["Статус"].replace(
    ["", "Ок"], ["В обработке", "Все в порядке"]
)

# Дополнение таблицы формой ответов

ws_ans = gc.open_by_key(
    "1fZhfUDWSGGr6uHQVdMpA1O2KNX32uXpKe8hMMNkoeMM"
)  # Ответы на форму
data = pd.DataFrame(ws_ans.get_worksheet(0).get_all_records())

data = data[
    data["Номер студенческого билета"] != ""
]  # оставляем лишь тех, у кого есть номер студака

list_stud = list(data_bdns["Номер студенческого"])

data = data[
    ~data["Номер студенческого билета"].isin(list_stud)
]  # Этих людей ещё не внесли в базу

data = data[["Номер студенческого билета", "Курс", "Статус"]]
data = data.rename(columns={"Номер студенческого билета": "Номер студенческого"})
data["Статус"] = data["Статус"].replace(
    ["Внести", "Продлить", "", "Ошибка"],
    ["В обработке", "В обработке", "В обработке", "В обработке"],
)

# Объединение датафреймов

data_bdns = pd.concat([data_bdns, data], ignore_index=True).fillna("В обработке")


# Преобразование "Номер студенческого" в int с заменой на стандартное значение в случае неудачи
def safe_cast(val, to_type, default=0):
    try:
        return to_type(val)
    except (ValueError, TypeError):
        print(f"Не удалось преобразовать {val} в {to_type}")
        return default


data_bdns["Номер студенческого"] = data_bdns["Номер студенческого"].map(
    lambda x: safe_cast(x, int)
)
data_bdns = data_bdns.sort_values(by=["Номер студенческого"])

# Проверка истекания даты

search = lambda x: (
    datetime.strptime(x, "%d.%m.%Y") <= datetime.today()
    if re.search(r"\d{2}.\d{2}.\d{4}", x)
    else False
)
data_bdns["tmp"] = data_bdns["Истечение документов"].map(
    search
)  # создаём вспомогательный столбец для проверки истечения даты

data_bdns["Статус"][data_bdns["tmp"]] = "Истек срок действия документов"
del data_bdns["tmp"]

# Загрузка файла на диск

sh = gc.open_by_key(
    "1XYnZHF1nyA4RANVcNYgn1AxKudt0xC8C-XytDrAf8ck"
)  # id итоговой таблицы
worksheet = sh.get_worksheet(0)
worksheet.update([data_bdns.columns.values.tolist()] + data_bdns.values.tolist())
