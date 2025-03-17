import gspread
import json
import yagmail
import argparse
from typing import List, Dict, Tuple

# Функция для загрузки конфигурации из JSON файла
def load_config(config_path: str) -> Dict:
    with open(config_path) as data:
        config = json.load(data)
    return config

# Функция для загрузки данных из Google Sheets
def load_sheet_data(sheet_name: str, column_name: str) -> List[Tuple[int, str]]:
    gc = gspread.service_account()
    wks = gc.open(sheet_name).sheet1
    col_index = wks.row_values(1).index(column_name) + 1
    emails = wks.col_values(col_index)
    # Возвращаем список мыл (номер строки, email)
    return [(i + 2, email) for i, email in enumerate(emails[1:])]  # +2, так как первая строка - заголовок, а enumerate начинается с 0

# Функция для отправки писем
def send_emails(email_data: List[Tuple[int, str]], config: Dict, sent_emails_path: str, invalid_rows_path: str):
    yag = yagmail.SMTP(config['email'], config['key'], host=config['host'], port=config['port'])
    sent_emails = []
    invalid_rows = []

    # Заголовок и содержание письма берутся из конфигурации
    subject = config.get('subject', 'Без темы')  # По умолчанию "Без темы", если поле отсутствует
    body = config.get('body', '')  # По умолчанию пустое содержание

    for row_number, email in email_data:
        if '@' in email:  # Простая проверка на валидность email
            yag.send(to=email, subject=subject, contents=body)
            sent_emails.append(email)
            print(f'Письмо отправлено на {email} (строка {row_number})')
        else:
            invalid_rows.append(row_number)
            print(f'Неверный адрес в строке {row_number}: {email}')

    # Сохранение отправленных email и номеров строк с неверными адресами
    with open(sent_emails_path, 'w') as file:
        json.dump(sent_emails, file)
    with open(invalid_rows_path, 'w') as file:
        json.dump(invalid_rows, file)

def main():
    # Парсинг аргументов командной строки
    parser = argparse.ArgumentParser(description="Скрипт для рассылки писем.")
    parser.add_argument('--config', type=str, default='config.json', help='Путь к файлу конфигурации (по умолчанию: config.json)')
    parser.add_argument('--sent-emails', type=str, default='sent_emails.json', help='Путь к файлу для сохранения отправленных email (по умолчанию: sent_emails.json)')
    parser.add_argument('--invalid-rows', type=str, default='invalid_rows.json', help='Путь к файлу для сохранения номеров строк с неверными адресами (по умолчанию: invalid_rows.json)')
    args = parser.parse_args()

    # Загрузка конфигурации
    config = load_config(args.config)

    # Загрузка данных из Google Sheets
    email_data = load_sheet_data("TestTable", "Email")  # Замените "Emails" на название вашего столбца

    # Отправка писем
    send_emails(email_data, config, args.sent_emails, args.invalid_rows)

if __name__ == "__main__":
    main()
