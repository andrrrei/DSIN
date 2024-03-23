import pygsheets
import pandas as pd

DATE = 'Отметка времени'
FIO = 'ФИО'
BIRTH = 'Введите свою дату рождения'
STUD = 'Номер студенческого билета'
PROF = 'Номер профсоюзного билета'
PAY = 'Форма обучения'
STAGE = 'Направление'
YEAR = 'Курс'
ACC_NUM = 'Номер счёта (в формате 408ХХ810ХХХХХХХХХХХХ)'
ADRESS = 'Почтовый индекс и место прописки (в формате: индекс, адрес)'
CATEGORY = 'Категория по ОПК'
VK = 'Прикрепите ссылку на Вашу страницу ВК (в формате https://vk.com/lebedev.andrei)'
PHONE = 'Введите Ваш номер телефона в формате 8ХХХХХХХХХХ'
PASSPORT = 'Прикрепите скан паспорта (страницы 2-3 и 4-5) одним файлом в формате pdf, название файла - Ваши ФИО'
FORM = 'Прикрепите анкету на вступление в формате pdf, название файла - Ваши ФИО'
ACCOUNT = 'Прикрепите реквизиты счёта карты МИР Сбербанка России в формате pdf, название файла - Ваши ФИО'
CONFIRM = 'Прикрепите все подтверждающие документы одним pdf файлом, название файла - Ваши ФИО'
VALID = 'Срок действия'

gc = pygsheets.authorize(service_file = '/home/andrei/Python/client_secret.json')

base = gc.open_by_key('1Kzx-W1TXHVhjdL2nl9SfK235hKngvfy_F_1pvtoGTkM')
answers = gc.open_by_key('1fZhfUDWSGGr6uHQVdMpA1O2KNX32uXpKe8hMMNkoeMM')

sh_answers = answers[0]
df_answers = sh_answers.get_as_df()

sh_base1 = base[0]
sh_base2 = base[1]
df_free = sh_base1.get_as_df() #free prof members
df_paid = sh_base2.get_as_df() #paid prof members

def move_to_base(cur_row):
    global df_free
    global df_paid

    s = [''] * 3
    fio = cur_row[FIO].split(' ')
    for i in range(len(fio)):
        s[i] = fio[i]

    link = cur_row[VK]
    link = str(link)
    if link[0] == '@':
        link = link.replace('@', 'https://vk.com/')
    elif 'https://vk.com/' not in link:
        link = 'https://vk.com/' + link
    cur_row[VK] = link
    new_row = [cur_row[DATE].split(' ')[0], 
               cur_row[FIO].split(' ')[0], cur_row[FIO].split(' ')[1], cur_row[FIO].split(' ')[2],
               cur_row[BIRTH],
               int(cur_row[STUD]), int(cur_row[PROF]),
               cur_row[PAY], cur_row[STAGE], cur_row[YEAR], 
               cur_row[ACC_NUM],
               cur_row[ADRESS],
               cur_row[CATEGORY],
               cur_row[VK],
               cur_row[PHONE],
               cur_row[PASSPORT],
               cur_row[FORM],
               cur_row[ACCOUNT],
               cur_row[CONFIRM],
               '', 
               cur_row[VALID],
               'Документы отправлены в ОПК']
    new_row = pd.DataFrame([new_row], columns = df_free.columns)
    if cur_row[PAY] == 'бюдж':
        df_free = pd.concat([df_free, new_row], ignore_index = True)
    elif cur_row[PAY] == 'контр':
        df_paid = pd.concat([df_paid, new_row], ignore_index = True)
    #здесь добавить объединение пдфок


def change_in_base(cur_row):
    global df_free
    global df_paid
    df = df_free if cur_row[PAY] == 'бюдж' else df_paid
    k = 0
    for row in df.itertuples():
        fio = row.Фамилия + ' ' + row.Имя + ' ' + row.Отчество
        if fio == cur_row['ФИО']:
            break
        k += 1
    if cur_row[PASSPORT] != '':
        df.at[k, 'Паспорт'] = cur_row[PASSPORT]
    if cur_row[FORM] != '':
        df.at[k, 'Анкета'] = cur_row[FORM]
    if cur_row[ACCOUNT] != '':
        df.at[k, 'Реквизиты'] = cur_row[ACCOUNT]
    if cur_row[CONFIRM] != '':
        df.at[k, 'Подтверждение'] = cur_row[CONFIRM]
    if cur_row['Новые ФИО'] != '':
        s = cur_row['Новые ФИО'].split(' ')
        df.at[k, 'Фамилия'] = s[0]
        df.at[k, 'Имя'] = s[1]
        df.at[k, 'Отчество'] = s[2] if len(s) == 3 else ''
    df.at[k, 'Статус'] = 'Документы отправлены в ОПК'
    if cur_row[PAY] == 'бюдж':
        df_free = df
    else:
        df_paid = df
    

# df_free['Студенческий'] = df_free['Студенческий'].apply(int)
# df_free['Профсоюзный'] = df_free['Профсоюзный'].apply(int)
# df_free['Студенческий'] = df_free['Студенческий'].apply(str)
# df_free['Профсоюзный'] = df_free['Профсоюзный'].apply(str)

check_range = df_answers.index[df_answers['База данных'] == ''].tolist()
if len(check_range) > 0:
    start = check_range[0]
    finish = check_range[-1]
    
    for i in range(start, finish + 1):
        cur_row = df_answers.iloc[i]
        cur_row['Форма обучения'] = 'бюдж' if cur_row['Форма обучения'] == 'Бюджет' else 'контр'
        cur_row['Направление'] = 'б' if cur_row['Направление'] == 'Бакалавриат' else 'м'
        if cur_row['Статус'] == 'Внести':
            move_to_base(cur_row)
        elif cur_row['Статус'] == 'Продлить':
            change_in_base(cur_row)
        df_answers.at[i, 'База данных'] = 'Обработано'

    sh_base1.set_dataframe(df_free, (0, 0))
    sh_base2.set_dataframe(df_paid, (0, 0)) 
    sh_answers.set_dataframe(df_answers, (0, 0))
else:
    print('Все студенты обработаны')
