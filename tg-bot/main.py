#!/usr/bin/python3.3
import threading, telebot, schedule, main, time
import os, sys
from telebot import types
from telegram.constants import ParseMode
import for_json
import admin

bot = telebot.TeleBot("TOKEN")

@bot.message_handler(commands=['restart'])
@admin.admin_command
def restart_bot(message):
    admin.send_admin_message('bye')
    os.execv(sys.executable, ['python'] + sys.argv)


@bot.message_handler(commands=['start', 'help'])
def start(message):
    text = "Этот бот предназначается только для людей, состоящих в блоке БДНС"
    if for_json.has_access(message.from_user.id):
        text = "Мои команды для тебя:\
            \n/run - запустить какой-либо скрипт\
            \n/file - получить файлы после срабатывания программ\
            \n/debug - получить файлы отладки программ"
        send_message(message.from_user.id, text)

    if admin.is_admin(message):
        text = "И команды только для админа:\
            \n/restart - перезапуск бота.\
            \n/access - просмотр списка доступа\
            \n/access <i>id_пользователя</i> <i>-1/0/1</i> - изменить права доступа"
        admin.send_admin_message(text)

    return


@bot.message_handler(commands=['access'])
def request_access(message):
    if admin.is_admin(message):
        admin.admin_access(message)
        return

    if for_json.return_info(message.from_user.id) != None:
        state = for_json.return_info(message.from_user.id)["state"]
        if state == 1:
            text = "У вас есть доступ!"
        elif state == -1:
            text = "В доступе отклонено!"
        else:
            text = "Ваш запрос на доступ уже отправлен, ожидайте"

        send_message(message.from_user.id, text)
        return

    for_json.save_user(message)
    send_message(message.from_user.id, "Ваш запрос на доступ отправлен")

    text = "Запрос на доступ от: @{} \nid: <code>{}</code>".format(message.from_user.username,
                                                                 message.from_user.id)

    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton(text="✅ Предоставить", callback_data="(access)#1#" + str(message.from_user.id)),
               types.InlineKeyboardButton(text="❌ Отказать", callback_data="(access)#-1#" + str(message.from_user.id)))

    admin.send_admin_message(text, markup)

    return

@bot.callback_query_handler(func=lambda call: "(access)" in call.data)
def give_access(call):
    tmp, state, id = call.data.split("#")
    state = int(state)

    markup = types.InlineKeyboardMarkup()
    text = ""

    if state == 1:
        text = "✅ Доступ @{} предоставлен!".format(for_json.return_info(id)["username"])
        markup.add(types.InlineKeyboardButton(text="Отказать", callback_data="(access)#-1#" + id))
    else:
        text = "❌ В доступе @{} отказано!".format(for_json.return_info(id)["username"])
        markup.add(types.InlineKeyboardButton(text="Предоставить", callback_data="(access)#1#" + id))

    for_json.change_state(id, state)

    bot.edit_message_text(chat_id=call.message.chat.id,
                              message_id=call.message.message_id,
                              text=text,
                              parse_mode=ParseMode.HTML,
                              reply_markup=markup)

    return


@bot.message_handler(commands=['run']) # send list of scripts
def run(message):
    if not for_json.has_access(message.from_user.id):
        send_message(message.from_user.id, "У вас нет на это прав.")
        return

    # print(os.path.dirname(os.path.abspath(__file__)))
    dir = os.getcwd() + "/scripts"
    files = os.listdir(dir)

    markup = types.InlineKeyboardMarkup()
    for i in files:
        if ".py" not in i:
            continue
        markup.add(types.InlineKeyboardButton(text=i, callback_data="(run)#" + i))

    send_message(message.from_user.id, "Возможные программы для запуска:", markup)

    return


@bot.callback_query_handler(func=lambda call: "(run)" in call.data)
def run_program(call):
    # print(os.path.dirname(os.path.abspath(__file__)))
    dir = os.getcwd() + "/scripts/" + call.data.split('#')[1]

    os.system('python3 ' + dir)

    bot.edit_message_text(chat_id=call.message.chat.id,
                              message_id=call.message.message_id, 
                              text='🚀 Был запущен файл <code>{}</code>!'.format(call.data.split('#')[1]), 
                              parse_mode=ParseMode.HTML)

    return




@bot.message_handler(commands=['file', 'debug']) # send list of files
def file(message):
    if not for_json.has_access(message.from_user.id):
        send_message(message.from_user.id, "У вас нет на это прав.")
        return

    dir = ""

    if "debug" in message.text:
        dir = os.getcwd() + "/scripts/debug/"
    else:
        dir = os.getcwd() + "/scripts/files/"

    files = os.listdir(dir)

    send_message(message.from_user.id, "Вот ваши файлы:")

    for path in files:
        path = dir + path
        with open(path, 'rb') as file: 
            send_document(message.from_user.id, file)

    return


@bot.message_handler(commands=['nohup']) # send logs
def nohup(message):
    if not for_json.has_access(message.from_user.id):
        send_message(message.from_user.id, "У вас нет на это прав.")
        return

    dir = os.getcwd() + "/nohup.out"
    
    files = os.listdir(dir)

    send_message(message.from_user.id, "Вот ваш файл логов:")
    send_document(message.from_user.id, file)

    return

def send_message(id, text, markup=None):
    try:
        bot.send_message(chat_id=id, text=text, parse_mode=ParseMode.HTML, disable_web_page_preview=True, reply_markup=markup)
    except Exception as e:
        time.sleep(5)
        try:
            bot.send_message(chat_id=id, text=text, parse_mode=ParseMode.HTML, disable_web_page_preview=True, reply_markup=markup)
        except Exception as e:
            text_error = 'Error from user: @{} <code>{}</code>\n{}'.format(for_json.return_info(id)["username"], id, str(e))
            admin.send_admin_message(text_error)
            print(id, e)

    return

def send_document(id, file, text = '', markup=None):
    try:
        bot.send_document(id, file, caption=text, reply_markup=markup)
    except Exception as e:
        time.sleep(5)
        try:
            bot.send_document(id, file, caption=text, reply_markup=markup)
        except Exception as e:
            text_error = 'Error from user: <code>{}</code>\n{}'.format(id, str(e))
            admin.send_admin_message(text_error)
            print(id, e)

    return


if __name__ == '__main__':
    admin.send_admin_message('Я перезапустился!')
    print("-------------------------")

    schedule.every().day.at("18:00").do(os.system, 'python3 ' + os.getcwd() + "/scripts/stat_base.py") # для обсновления статусов

    threading.Thread(target=bot.infinity_polling, name='bot_infinity_polling', daemon=True).start()
    
    while True:
        try:
            schedule.run_pending()
            time.sleep(5)
        except Exception as e:
            print(e)
            time.sleep(10)
