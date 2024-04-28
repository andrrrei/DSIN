#!/usr/bin/python3.3
import json

path_users = "users.json"

def save_user(message):
    data = {}
    with open(path_users, 'r', encoding='utf-8') as json_file: 
        data = json.load(json_file)

    data[str(message.from_user.id)] = {
        "username": message.from_user.username,
        "first_name": message.from_user.first_name, 
        "last_name": message.from_user.last_name, 
        "state": 0
    }

    with open(path_users, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
                                                        
    return

def change_state(id, state):
    id = str(id)
    state = int(state)
    
    data = {}
    with open(path_users, 'r', encoding='utf-8') as json_file: 
        data = json.load(json_file)
    
    data[id]["state"] = state
    
    with open(path_users, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
    
    return

def has_access(id):
    return return_info(id)["state"] == 1

def return_info(id):
    id = str(id)
    data = {}
    with open(path_users, 'r', encoding='utf-8') as json_file: 
        data = json.load(json_file)
    data = data.get(id)
    
    return data

def return_json():
    data = {}
    with open(path_users, 'r', encoding='utf-8') as json_file: 
        data = json.load(json_file)
        
    return data
