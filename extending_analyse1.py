# __________ #
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
sheets_service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

gc = pygsheets.authorize(service_file = 'credentials.json')


# Считывание ID таблиц начальников курса
# tables = []
# for i in range(6):
#     tables.append(input())
table1 = '1wtEz3onhjSIZHsPpg0qm8ULOHX8dut_HD2cFCBSuZ8I'
table2 = '1xgv8U5hV9Dt0TkSht4f0X5UaCbXPF0kEnuVB3FQ2-vE'
table3 = '1zlc6TxcR9O0C2sjx2_A2fEXAanfGTCz5Nd_FXWb1kT4'
table4 = '1lf619U3wYA-Ex8xRMniZZDH7CYAi6tAiZk_qsxzUDS8'
table5 = '1Jmnuaxx7c8ZNNKDeLY5ftfcMiukzwCKKHioWXwzZmOQ'
table6 = '1ZaDMcPsslzJXeWYBCW-vgQnjsO1wrSe43LaC9mx_wG8'

base1 = gc.open_by_key(table1)
base2 = gc.open_by_key(table2)
base3 = gc.open_by_key(table3)
base4 = gc.open_by_key(table4)
base5 = gc.open_by_key(table5)
base6 = gc.open_by_key(table6)

df_base1 = base1[0]
df_base1_1 = base1[1]
df_base2 = base2[0]
df_base2_1 = base2[1]
df_base3 = base3[0]
df_base3_1 = base3[1]
df_base4 = base4[0]
df_base4_1 = base4[1]
df_base5 = base5[0]
df_base5_1 = base5[1]
df_base6 = base6[0]
df_base6_1 = base6[1]

df1 = df_base1.get_as_df()
df1_1 = df_base1_1.get_as_df()
df2 = df_base2.get_as_df()
df2_1 = df_base2_1.get_as_df()
df3 = df_base3.get_as_df()
df3_1 = df_base3_1.get_as_df()
df4 = df_base4.get_as_df()
df4_1 = df_base4_1.get_as_df()
df5 = df_base5.get_as_df()
df5_1 = df_base5_1.get_as_df()
df6 = df_base6.get_as_df()
df6_1 = df_base6_1.get_as_df()

# Объединение всех датафреймов один
df = pd.concat([df1, df2, df3, df4, df5, df6, df1_1, df2_1, df3_1, df4_1, df5_1, df6_1])

# Первоначальная подготовка данных: отсеивание отчисленных, сортировка, перестановка столбцов, создание столбца для профсоюзного билета
df2 = df.loc[df['Статус'] != 'отчислен']
df2 = df2.sort_values(['Фамилия']).reset_index(drop=True)
df2_final = df2.iloc[:,[0, 1, 2, 4, 5, 6]]
df2_final = df2_final[['Фамилия', 'Имя', 'Отчество', 'Курс', 'Статус', 'Студенческий']]
df2_final['Номер проф. билета'] = ''

# Первый этап обработки данных из таблиц начальников курса завершён


# Второй этап обработки данных с помощью базы ОПК
base = gc.open_by_key('1KjM9rWLjl4T8Vp4DGWnb4pNRTTtwdh12-qdDld5X9jA')
df_base = base[0]
df_base1 = base[1]
df_base2 = base[2]
df_1 = df_base.get_as_df()
df_2 = df_base1.get_as_df()
df_3 = df_base2.get_as_df()
df = pd.concat([df_1, df_2, df_3])
x, y = df2_final.shape
final_df = df2_final

# С помощью базы данных заполняем недостающие данные
for i in range(0, int(x)):
    if str(final_df.iat[i, 4]) == 'закрылся':
        final_df.iat[i, 4] = 'да'
    surname = final_df.iat[i, 0]
    name = final_df.iat[i, 1]
    df2 = df.loc[(df['Фамилия'] == surname) & (df['Имя'] == name)]
    if str(df2.iat[0, 6]) == "м":
        final_df.iat[i, 3] = str(df2.iat[0, 7])[0] + 'м'
    else:
        final_df.iat[i, 3] = str(df2.iat[0, 7])[0]
    final_df.iat[i, 5] = str(df2.iat[0, 3])
    final_df.iat[i, 6] = str(df2.iat[0, 4])

# Редактируем названия столбцов, добавляем столбец с нумерацией
final_df.rename(columns={'Статус': 'Промежуточная аттестация пройдена', 'Студенческий': 'Номер студ. билета'}, inplace=True)
indexes = range(1, len(final_df) + 1)
final_df.insert(loc=0, column='X', value=indexes)
final_df['X'] = final_df['X'].astype(str)

# Второй этап обработки данных завершён, создаём промежуточный CSV-файл с данными
csv_data = final_df.to_csv (r' my_data.csv', index=False)
