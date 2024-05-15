# __________ #
import gspread
import pygsheets
import pandas as pd
import datetime
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
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
sheets_service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

gc = pygsheets.authorize(service_file='credentials.json')

inter_file = ' my_data.csv'

# Считывание ID таблицы ответов на форму
table1 = '1fZhfUDWSGGr6uHQVdMpA1O2KNX32uXpKe8hMMNkoeMM'

# Считывание даты, с которой нужно брать данные
date_period_str = "01.01.2024"
date_period_obj = datetime.datetime.strptime(date_period_str, '%d.%m.%Y')

base = gc.open_by_key(table1)
df_base = base[0]
df = df_base.get_as_df()

# Отсеиваем только нужный период
df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format='%d.%m.%Y %H:%M:%S')
df = df.loc[df['Отметка времени'] >= date_period_obj]
df2 = df.loc[df['Выберите причину продления'] == 'Истечение срока действия подверждающих документов']

# Считали данные таблицы в датафрейм, вытаскиваем только нужные нам (продление), сортируем
# Если людей на продление нет, завершаем работу
if df2.empty:
    print("Нет людей на продление в этом периоде")
    exit(1)

# Обновление сроков истечение документов в таблице БДНС-Статус
# Подключаемся к базе статусов БДНС
# Установка доступа к Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
status_table = '1OcFvuJ0n2PcyPaBVJIp619j08KE9TWWMrzCzClDCyjM'
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc2 = gspread.authorize(credentials)
status_base = gc2.open_by_key(status_table).sheet1

# Получаем все данные из таблицы в виде списка списков
data = status_base.get_all_values()

# Создаем DataFrame из полученных данных
df = pd.DataFrame(data[1:], columns=data[0])

# Извлекаем номер студенческого и срок действия документов
x, y = df2.shape
for i in range(0, x):
    # Находим индекс строки, в которой нужно изменить значение
    value = df2.iat[i, 3]
    index = df[df['Номер студенческого'] == str(value)].index.tolist()
    # Меняем значение в другом столбце
    if index:
        index = index[
                    0] + 2  # Прибавляем 2, потому что индексация в Google Sheets начинается с 1 и нумерация строк считается с 2
        status_base.update_cell(index, 3, df2.iat[i, 22])
        print("Значение обновлено успешно!")
    else:
        print("Значение для обновления данных не найдено.")

# Создание верного имени файлов
df2 = df2.sort_values(['ФИО']).reset_index(drop=True)
new_df = df2['ФИО'].str.split(expand=True)
new_df.columns = ['Фамилия', 'Имя', 'Отчество']
x, y = new_df.shape
for i in range(0, int(x)):
    new_df.iat[i, 0] = str(new_df.iat[i, 0]) + str(new_df.iat[i, 1])[0] + str(new_df.iat[i, 2])[0]

# Переименование столбцов и подготовка записи датафрейма в CSV-файл
new_df = new_df['Фамилия']
new_df.rename('Название файла')
df2 = df2.iloc[:, [21, 22]]
final_df = pd.concat([new_df, df2], axis=1)
final_df.rename(columns={'Тип справки': 'Документы'}, inplace=True)
indexes = range(1, len(final_df) + 1)
final_df.insert(loc=0, column='X', value=indexes)
final_df['X'] = final_df['X'].astype(str)

# Второй этап обработки данных завершён, создаём промежуточный CSV-файл с данными (Без промежуточного файла возникает ошибка)
csv_data = final_df.to_csv(inter_file, index=False)

# ------------------- Работа с гугл-документом -------------------------------------#
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Получение текущей даты
current_date = datetime.now().date()

# Подставьте путь к файлу с вашими учетными данными и ID вашей целевой папки
SERVICE_ACCOUNT_FILE = 'credentials.json'
#
FOLDER_ID = '14blDt4JalikhJUoenyKLUeT0dxj-sq7G'
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

months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
          'декабрь']
# Текст, который нужно вставить в документ
text_to_insert = f"""
Справки для БДНС
Факультет ВМК
\n
 "{current_date.day}" {months[current_date.month - 1]} {current_date.year}г.\n """

# Считывание файла с данными в датафрейм, приведение всех данных в строковый тип
df = pd.read_csv(inter_file)
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


