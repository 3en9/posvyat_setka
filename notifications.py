import telebot
import pandas as pd
import time
from datetime import datetime
TOKEN = '7767795518:AAHyp07SVczH6joO3zD2_VbrtRZd24yzlqQ'
tel = telebot.TeleBot(TOKEN)

def get_user_id(user_tag: str):
    user_id = 0
    with open('users.txt') as f:
        data = f.read().split('\n')
        for i in data:
            try:
                tag, tag_id = i.split()
                tag_id = int(tag_id)
                if tag == user_tag:
                    user_id = tag_id
                    break
            except:
                break
    return user_id

df = pd.read_excel(r'СЕТКА.xlsx', sheet_name = 'ЛИЧНЫЕ СЕТКИ')
n = 109 #количество значимых строк(кол-во оргов + первые 2)
arr = df.to_numpy()
name_columns = arr[0][8::]
arr = arr[1:n-1]
print(name_columns)
while True:
    try:
        cur_time = datetime.now()
        cur_hour = cur_time.hour
        cur_minute = cur_time.minute
        cur_second = cur_time.second
        if cur_second == 30:
            for i in range(len(name_columns)):
                start_time = name_columns[i].split('-')
                # print(start_time)
                next_hour, next_minute = list(map(int, start_time[0].split(':')))
                if next_hour == cur_hour and next_minute == cur_minute:
                    for user in arr:
                        try:
                            tg = user[1].replace('@', '')
                            action = user[i+8]
                            tel.send_message(get_user_id(tg), f'Текущая точка:\n{action}')
                        except:
                            continue
                if next_minute-cur_minute == 5 and next_hour == cur_hour or next_hour == cur_hour+1 and cur_minute==55:
                    for user in arr:
                        try:
                            tg = user[1].replace('@', '')
                            action = user[i + 8]
                            tel.send_message(get_user_id(tg), f'Через 5 минут начинается точка:\n{action}')
                        except:
                            continue
            time.sleep(10)
    except:
        pass
