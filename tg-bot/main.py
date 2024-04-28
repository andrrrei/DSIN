#!/usr/bin/python3.3
import telebot, main, time
import os, sys
from telebot import types
from telegram.constants import ParseMode
import for_json
# from admin import admin_command, send_admin_message, is_admin # позже заменить
import subprocess
import admin

bot = telebot.TeleBot("TOKEN")

@bot.message_handler(commands=['restart'])
@admin.admin_command
def restart_bot(message):
    admin.send_admin_message('bye')
    os.execv(sys.executable, ['python'] + sys.argv)


@bot.message_handler(commands=['start'])
def start(message):
    send_message(message.from_user.id, "Этот бот предназначается только для людей, состоящих в блоке БДНС")

    pass


@bot.message_handler(commands=['access'])
def request_access(message):
    if admin.is_admin(message):
        # send_message(message.from_user.id, "Вы админ!")
        admin.admin_access(message)
        return

    if for_json.return_info(message.from_user.id) != None:
        send_message(message.from_user.id, "Ваш запрос на доступ уже отправлен, ожидайте")
        return

    for_json.save_user(message)
    send_message(message.from_user.id, "Ваш запрос на доступ отправлен")

    text = "Запрос на доступ от: @{} \nid: <code>{}</code>".format(message.from_user.username,
                                                                 message.from_user.id)

    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton(text="✅ Предоставить", callback_data="(access)_1_" + str(message.from_user.id)),
               types.InlineKeyboardButton(text="❌ Отказать", callback_data="(access)_-1_" + str(message.from_user.id)))

    admin.send_admin_message(text, markup)

    return

@bot.callback_query_handler(func=lambda call: "(access)" in call.data)
def give_access(call):
    tmp, state, id = call.data.split("_")
    state = int(state)

    markup = types.InlineKeyboardMarkup()
    text = ""

    if state == 1:
        text = "✅ Доступ @{} предоставлен!".format(for_json.return_info(id)["username"])
        markup.add(types.InlineKeyboardButton(text="Отказать", callback_data="(access)_-1_" + id))
    else:
        text = "❌ В доступе @{} отказано!".format(for_json.return_info(id)["username"])
        markup.add(types.InlineKeyboardButton(text="Предоставить", callback_data="(access)_1_" + id))

    for_json.change_state(id, state)

    bot.edit_message_text(chat_id=call.message.chat.id,
                              message_id=call.message.message_id,
                              text=text,
                              parse_mode=ParseMode.HTML,
                              reply_markup=markup)

    return


def run_program(command):
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    (output, error) = process.communicate()
    return output.decode().strip()

@bot.message_handler(commands=['code'])
def code(message):
    if not for_json.has_access(message.from_user.id):
        send_message(message.from_user.id, "У вас нет на это прав.")
        return

    result = subprocess.run(["scripts/test.py"], capture_output=True)
    print(result.stdout)

    return

    process = subprocess.Popen("./scripts/test.py", stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    (output, error) = process.communicate()
    admin.send_admin_message(output.decode().strip())

# os.startfile('<Путь>') или os.system('"Каталог1"\"Каталог2"\"Нужный файл"')
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

    bot.infinity_polling()

    # threading.Thread(target=bot.infinity_polling, name='bot_infinity_polling', daemon=True).start()
