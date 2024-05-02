import main as bot
import for_json

admin_id = '856850518' # krosh
# admin_id = '397016066'

def admin_command(func):
    def wrapper(*args, **kwargs):
        message = args[0]
        if str(message.from_user.id) != admin_id:
            bot.send_message(message.from_user.id, 'Команда доступна только администратору')
            return

        func(*args, **kwargs)
        return

    return wrapper

def is_admin(message):
    if str(message.from_user.id) == admin_id and message.from_user.id == message.chat.id:
        return True

    return False

def send_admin_message(text, markup=None):
    bot.send_message(admin_id, '🛑 ' + text, markup=markup)

    return

def send_admin_document(file, markup=None):
    bot.send_document(admin_id, file, '🛑 Admin file', markup=markup)

    return

def admin_access(message):
    text = message.text

    if len(text.split(" ")) == 3:          # Команда вида /access id state
        tmp, id, state = text.split(" ")

        if for_json.return_info(id) == None:
            send_admin_message("Такой id не найден!")

        if state not in ["-1", "1", "0"]:
            send_admin_message("Такой state не найден!")

        for_json.change_state(id, state)
        send_admin_message("Успешно!")

    data = for_json.return_json()
    send_admin_message("✅ Пользователи с доступом:\n" + \
        ''.join(["@{} <code>{}</code>\n".format(data[i]["username"], i) if data[i]["state"] == 1 else "" for i in data.keys()]))
    
    send_admin_message("🕒 Пользователи в ожидании:\n" + \
        ''.join(["@{} <code>{}</code>\n".format(data[i]["username"], i) if data[i]["state"] == 0 else "" for i in data.keys()]))
    
    send_admin_message("❌ Пользователи без доступа:\n" + \
        ''.join(["@{} <code>{}</code>\n".format(data[i]["username"], i) if data[i]["state"] == -1 else "" for i in data.keys()]))


    # Прогнаться по списку

    return
    