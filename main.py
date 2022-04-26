import os
import sqlite3
import numpy as np
from openpyxl import load_workbook
import httplib2
import pandas as pd
import vk_api
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
from vk_api.longpoll import VkLongPoll
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import TOKEN_bot

post_mail = [[]]
money = 0
wb = pd.read_excel('CDEK-offices_ru (1).xlsx', sheet_name="Россия")
book = load_workbook("product_in_stock.xlsx", data_only=True)
sheet_xlsl = book["Лист1"]
oformlenie = "Нажмите кнопку <сделать заказ> чтобы начать оформление заказа"
city_user = ""
product_in_stock_wb = pd.read_excel("product_in_stock.xlsx", sheet_name="Лист1")
true_city = 0


vk_session = vk_api.VkApi(token=TOKEN_bot.TOKEN)
session_api = vk_session.get_api()
longpool = VkLongPoll(vk_session)


db = sqlite3.connect("info.db")
sql = db.cursor()
sql.execute("""CREATE TABLE IF NOT EXISTS users(
    userId  BIGINT,
    act TEXT,
    fio TEXT,
    date_of_birth TEXT,
    telephone TEXT,
    emal TEXT,
    pos_produc TEXT,
    city TEXT,
    post TEXT)""")
db.commit()
user_act = "0"


def get_service_sacc():
    creds_json = os.path.dirname(__file__) + "/pythonbotvk-3bcdc4b45418.json"
    scopes = ['https://www.googleapis.com/auth/spreadsheets']

    creds_service = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scopes).authorize(httplib2.Http())
    return build('sheets', 'v4', http=creds_service)


sheet = get_service_sacc().spreadsheets()
sheet_id = "1PbxNtvA6Kt6F3-IoRhBscrNNRmHaAyg8jiFnhXr-yPM"


def send_message(user_id, message, keyboard=None):
    button = {"user_id": user_id,
              "message": message,
              "random_id": 0,
              }

    if keyboard is not None:
        button["keyboard"] = keyboard.get_keyboard()
    else:
        button = button
    vk_session.method("messages.send", button)


def Fix_msg(msg):
    msg = "'" + msg + "'"
    return msg


def product_stock(id):
    global product_in_stock, not_in_stock
    not_in_stock = 0
    money = 0
    for value in sql.execute("SELECT * FROM users"):
        product_in_stock = value[6]
    colum_a = sheet_xlsl["A"]
    for i in range(1, len(colum_a)):
        if int(colum_a[i].value) == int(product_in_stock):
            money = sheet_xlsl['C' + str(i + 1)].value
            if sheet_xlsl['B' + str(i + 1)].value == 0:
                send_message(id, "Извините, но этот товар закончился. Выбирите другой")
                not_in_stock = 1
            else:
                send_message(id, "Введите ваш город")
    print(money)
    return money, not_in_stock


def post_mail_func(post_mail, id):
    global city_user, true_city
    for value in sql.execute("SELECT * FROM users"):
        city_user = value[7].capitalize()
    send_message(id, "Одну минтку, ищем для вас пункты выдачи")
    for i, row in wb.iterrows():
        if row['Город'] == city_user:
            true_city = 1
            post_mail[0].append(wb['Адрес'][i])
    if true_city == 0:
        send_message(id, "Извините, но мы не осуществляем доставку в ваш город")
    post_mail = np.asarray(post_mail).reshape(-1, 1)
    post_mail_colum = post_mail.tolist()
    if true_city == 1:
        rangeAll = '{0}!A1:ZZ'.format("Лист1")
        body = {}
        resultClear = sheet.values().clear(
            spreadsheetId=sheet_id,
            range=rangeAll,
            body=body).execute()
        resp = sheet.values().update(
            spreadsheetId=sheet_id,
            range='лист1!A1',
            valueInputOption="RAW",
            body={'values': post_mail_colum}).execute()
    send_message(id,
                 "https://docs.google.com/spreadsheets/d/1PbxNtvA6Kt6F3-IoRhBscrNNRmHaAyg8jiFnhXr-yPM/edit?usp=sharing")
    send_message(id, "Выберите пункт выдачи из предложенного списка")
    return true_city




def main():
    global city_user, resp, true_city, post_mail, product_in_stock, money, not_in_stock, recipient
    for event in longpool.listen():
        if event.type == vk_api.longpoll.VkEventType.MESSAGE_NEW and event.to_me:
            msg = event.text.lower()
            id = event.user_id
            sql.execute(f" SELECT userId FROM users WHERE userId = '{id}'")
            if sql.fetchone() is None:
                sql.execute("INSERT INTO users VALUES (?,?,?,?,?,?,?,?,?)",
                            (id, "newUser", "0", "0", "0", "0", "0", "0", "0"))
                db.commit()
                send_message(id, "Привет для начало оформления заказа напиши старт")
            else:
                user_act = sql.execute(f"SELECT act FROM users WHERE userId ='{id}'").fetchone()[0]
                if user_act == "newUser" and msg == "старт":
                    keyboard = VkKeyboard(one_time=True)
                    keyboard.add_button("Сделать заказ", color=VkKeyboardColor.PRIMARY)
                    send_message(id, oformlenie, keyboard)

                elif user_act == "newUser" and msg == "сделать заказ":
                    send_message(id, "Начало оформления заказа")
                    send_message(id, "Напишите свое ФИО")
                    sql.execute(f"UPDATE users SET act = 'Get_fio' WHERE userId = {id}")
                    db.commit()
                elif user_act == "Get_fio":
                    sql.execute(f"UPDATE users SET fio ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_date_of_birth' WHERE userId = {id}")
                    db.commit()
                    send_message(id, "Введите вашу дату рождения")
                elif user_act == "Get_date_of_birth":
                    sql.execute(f"UPDATE users SET date_of_birth ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_telephone' WHERE userId = {id}")
                    db.commit()
                    send_message(id, "Введите ваш номер телефона")
                elif user_act == "Get_telephone":
                    sql.execute(f"UPDATE users SET telephone ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_emal' WHERE userId = {id}")
                    db.commit()
                    send_message(id, "Введите вашу электронную почту")
                elif user_act == "Get_emal":
                    sql.execute(f"UPDATE users SET emal ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_pos_produc' WHERE userId = {id}")
                    db.commit()
                    send_message(id,
                                 "Введите позицию товара, который хотите заказть(позицию можно посмотреть на страничке сообщества)")
                elif user_act == "Get_pos_produc":
                    sql.execute(f"UPDATE users SET pos_produc ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_city' WHERE userId = {id}")
                    db.commit()
                    money, not_in_stock = product_stock(id)
                elif user_act == "Get_city" and not_in_stock == 1:
                    sql.execute(f"UPDATE users SET pos_produc ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_city' WHERE userId = {id}")
                    db.commit()
                    money, not_in_stock = product_stock(id)
                    not_in_stock = 0
                elif user_act == "Get_city" and not_in_stock == 0:
                    sql.execute(f"UPDATE users SET city ={Fix_msg(msg)} WHERE userId = {id}")
                    sql.execute(f"UPDATE users SET act = 'Get_post' WHERE userId = {id}")
                    db.commit()
                    post_mail_func(post_mail, id)
                elif user_act == "Get_post":
                    sql.execute(f"UPDATE users SET post ={Fix_msg(msg)} WHERE userId ={id}")
                    sql.execute(f"UPDATE users SET act = 'REG' WHERE userId ={id}")
                    db.commit()
                    send_message(id, f'Сумма: {money}\n Ссылка:')
                    send_message(id, "Спасибо за заказ")
                    send_message(id, "Если хотите еще раз заказать напишите <заказ>")
                elif user_act == "REG":
                    send_message(id, "Напишите номер позиции нового заказа")
                    sql.execute(f"UPDATE users SET act = 'Get_pos_produc' WHERE userId = {id}")
                    db.commit()


main()


