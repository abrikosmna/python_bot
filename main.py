import os
import sqlite3
import httplib2
import vk_api
import pandas as pd
import smtplib

from vk_api.keyboard import VkKeyboard, VkKeyboardColor
from vk_api.longpoll import VkLongPoll
from openpyxl import load_workbook
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from platform import python_version

import TOKEN_bot

post_mail = [[]]
money = 0
true_city = 0
product_in_stock = None
not_in_stock = None
recipient = None
city_user = ""
user_act = "0"
oformlenie = "Нажмите кнопку <сделать заказ> чтобы начать оформление заказа"
sheet_id = "1PbxNtvA6Kt6F3-IoRhBscrNNRmHaAyg8jiFnhXr-yPM"

wb = pd.read_excel('CDEK-offices_ru (1).xlsx', sheet_name="Россия")
book = load_workbook("product_in_stock.xlsx", data_only=True)
sheet_xlsl = book["Лист1"]
product_in_stock_wb = pd.read_excel("product_in_stock.xlsx", sheet_name="Лист1")

vk_session = vk_api.VkApi(token=TOKEN_bot.TOKEN)
session_api = vk_session.get_api()
longpool = VkLongPoll(vk_session)

db = sqlite3.connect("info.db")
sql = db.cursor()
sql.execute("""CREATE TABLE IF NOT EXISTS users(
    userId          BIGINT,
    act             VARCHAR (255),
    fio             VARCHAR (255),
    date_of_birth   VARCHAR (255),
    telephone       VARCHAR (255),
    emal            VARCHAR (255),
    pos_produc      VARCHAR (255),
    city            VARCHAR (255),
    post            VARCHAR (255))
    """)
db.commit()


def get_service_sacc():
    creds_json = os.path.dirname(__file__) + "/pythonbotvk-3bcdc4b45418.json"
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds_service = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scopes).authorize(httplib2.Http())
    return build('sheets', 'v4', http=creds_service)


sheet = get_service_sacc().spreadsheets()


def send_message(user_id, message, keyboard=None):
    button = {
        "user_id": user_id,
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


def product_stock(ID: int):
    global product_in_stock, not_in_stock
    not_in_stock = 0
    money_ = 0
    for value in sql.execute("SELECT * FROM users"):
        product_in_stock = value[6]
    colum_a = sheet_xlsl["A"]
    for i in range(1, len(colum_a)):
        if int(colum_a[i].value) == int(product_in_stock):
            money_ = sheet_xlsl['C' + str(i + 1)].value
            if sheet_xlsl['B' + str(i + 1)].value == 0:
                send_message(ID, "Извините, но этот товар закончился. Выверите другой")
                not_in_stock = 1
            else:
                send_message(ID, "Введите ваш город")
    print(money_)
    return money_, not_in_stock


def post_mail_func(email, ID: int):
    global city_user, true_city
    for value in sql.execute("SELECT * FROM users"):
        city_user = value[7].capitalize()
    send_message(ID, "Одну минтку, ищем для вас пункты выдачи")
    for i, row in wb.iterrows():
        if row['Город'] == city_user:
            true_city = 1
            email[0].append(wb['Адрес'][i])
    if true_city == 0:
        send_message(ID, "Извините, но мы не осуществляем доставку в ваш город")
    if true_city == 1:
        send_message(
            ID,
            "https://docs.google.com/spreadsheets/d/1PbxNtvA6Kt6F3-IoRhBscrNNRmHaAyg8jiFnhXr-yPM/edit?usp=sharing"
        )
        send_message(ID, "Выберите пункт выдачи из предложенного списка")
    return true_city


def send_email(money_):
    global recipient
    for value in sql.execute("SELECT * FROM users"):
        recipient = value[5].capitalize()
    user = TOKEN_bot.user
    password = TOKEN_bot.password
    server = 'smtp.gmail.com'
    str(recipient)
    sender = user
    subject = 'Заказ принят'
    text = f'Спасибо за заказ <br>' \
           f' сумма заказа {money_} <br>' \
           f' Ждем вас еще в нашем магазине'
    html = '<html><head></head><body><p>' + text + '</p></body></html>'
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = 'Python Bot <' + sender + '>'
    msg['To'] = ', '.join(recipient)
    msg['Reply-To'] = sender
    msg['Return-Path'] = sender
    msg['X-Mailer'] = 'Python/' + (python_version())
    part_text = MIMEText(text, 'plain')
    part_html = MIMEText(html, 'html')
    msg.attach(part_text)
    msg.attach(part_html)
    mail = smtplib.SMTP_SSL(server)
    mail.login(user, password)
    mail.sendmail(sender, recipient, msg.as_string())
    mail.quit()


def main():
    global city_user, true_city, post_mail, product_in_stock, money, not_in_stock, recipient
    for event in longpool.listen():
        if event.type == vk_api.longpoll.VkEventType.MESSAGE_NEW and event.to_me:
            msg = event.text.lower()
            ID = event.user_id
            sql.execute(f" SELECT userId FROM users WHERE userId = '{ID}'")
            if sql.fetchone() is None:
                sql.execute("INSERT INTO users VALUES (?,?,?,?,?,?,?,?,?)",
                            (ID, "newUser", "0", "0", "0", "0", "0", "0", "0"))
                db.commit()
                send_message(ID, "Привет для начало оформления заказа напиши старт")
            else:
                user_action = sql.execute(f"SELECT act FROM users WHERE userId ='{ID}'").fetchone()[0]
                if user_action == "newUser" and msg == "старт":
                    keyboard = VkKeyboard(one_time=True)
                    keyboard.add_button("Сделать заказ", color=VkKeyboardColor.PRIMARY)
                    send_message(ID, oformlenie, keyboard)

                elif user_action == "newUser" and msg == "сделать заказ":
                    send_message(ID, "Начало оформления заказа")
                    send_message(ID, "Напишите свое ФИО")
                    sql.execute(f"UPDATE users SET act = 'Get_fio' WHERE userId = {ID}")
                    db.commit()
                elif user_action == "Get_fio":
                    sql.execute(f"UPDATE users SET fio ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_date_of_birth' WHERE userId = {ID}")
                    db.commit()
                    send_message(ID, "Введите вашу дату рождения")
                elif user_action == "Get_date_of_birth":
                    sql.execute(f"UPDATE users SET date_of_birth ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_telephone' WHERE userId = {ID}")
                    db.commit()
                    send_message(ID, "Введите ваш номер телефона")
                elif user_action == "Get_telephone":
                    sql.execute(f"UPDATE users SET telephone ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_emal' WHERE userId = {ID}")
                    db.commit()
                    send_message(ID, "Введите вашу электронную почту")
                elif user_action == "Get_emal":
                    sql.execute(f"UPDATE users SET emal ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_pos_produc' WHERE userId = {ID}")
                    db.commit()
                    send_message(
                        ID,
                        "Введите позицию товара, который хотите заказaть"
                        "(позицию можно посмотреть на страничке сообщества)"
                    )
                elif user_action == "Get_pos_produc":
                    sql.execute(f"UPDATE users SET pos_produc ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_city' WHERE userId = {ID}")
                    db.commit()
                    money, not_in_stock = product_stock(ID)
                elif user_action == "Get_city" and not_in_stock == 1:
                    sql.execute(f"UPDATE users SET pos_produc ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_city' WHERE userId = {ID}")
                    db.commit()
                    money, not_in_stock = product_stock(ID)
                    not_in_stock = 0
                elif user_action == "Get_city" and not_in_stock == 0:
                    sql.execute(f"UPDATE users SET city ={Fix_msg(msg)} WHERE userId = {ID}")
                    sql.execute(f"UPDATE users SET act = 'Get_post' WHERE userId = {ID}")
                    db.commit()
                    post_mail_func(post_mail, ID)
                elif user_action == "Get_post":
                    sql.execute(f"UPDATE users SET post ={Fix_msg(msg)} WHERE userId ={ID}")
                    sql.execute(f"UPDATE users SET act = 'REG' WHERE userId ={ID}")
                    db.commit()
                    send_message(ID, f'Сумма: {money}\n Ссылка:')
                    send_message(ID, "Спасибо за заказ")
                    send_email(money)
                    send_message(ID, "Если хотите еще раз заказать напишите <заказ>")
                elif user_action == "REG":
                    send_message(id, "Напишите номер позиции нового заказа")
                    sql.execute(f"UPDATE users SET act = 'Get_pos_produc' WHERE userId = {ID}")
                    db.commit()


if __name__ == "__main__":
    main()
