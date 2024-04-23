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
print(final_df)
# Второй этап обработки данных завершён, создаём промежуточный CSV-файл с данными
csv_data = final_df.to_csv (r' my_data.csv', index= False)


