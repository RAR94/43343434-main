import telebot
from telebot import types
import os.path
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import phonenumbers
import datetime
import time
from threading import Thread

t = 0
name_list = 'Вох1 (копия) (копия) (копия)'
bot = telebot.TeleBot("5822620441:AAFd_htjU-_x_7oHXADdL8kRbYCgtFRiXUk")


class GoogleSheet:
    SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = None

    def __init__(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print('flow')
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('sheets', 'v4', credentials=creds)

    def updateRangeValues(self, range, values):
        data = [{
            'range': range,
            'values': values
        }]
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        result = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                                  body=body).execute()
        global edit
        edit = ('{0} cells updated.'.format(result.get('totalUpdatedCells')))

def delete():
    day1 = str(day).split()
    dat = day1[1].split('.')
    tim1 = tim.split(':')
    while True:
        now = datetime.datetime.now() - datetime.timedelta(minutes=5)
        d = datetime.datetime(int(dat[2]), int(dat[1]), int(dat[0]))
        if int(d.year) <= int(now.year):
            if int(d.month) <= int(now.month):
                if int(d.day) <= int(now.day):
                    if int(tim1[0]) <= int(int(now.hour)):
                        if int(tim1[1]) <= int(int(now.minute)):
                            dele = telebot.types.ReplyKeyboardRemove()
                            gs = GoogleSheet()
                            range = f'{name_list}!{t}:{chr(ord("@") + k)}{t}'
                            values = [
                                [""]
                            ]
                            gs.updateRangeValues(range, values)
                            break
        else:
            time.sleep(30)

@bot.message_handler(commands=['start'])
def start(message):
    dele = telebot.types.ReplyKeyboardRemove()
    bot.send_message(message.chat.id,
                     f'Виртуальный помощник приветствует Вас\n✅ Наш виртуальный помощник подберет для вас удобное время. \n✅ Выберите то что вам подходит из предложенных вариантов.\n✅ По окончанию диалога, вы будете записаны.',
                     reply_markup=dele)
    bot.send_message(message.chat.id, 'Пожалуйста введите ваше имя')



@bot.message_handler(content_types=['text'])
def get_name(message):
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    dele = telebot.types.ReplyKeyboardRemove()
    global name
    name = message.text
    bot.send_message(message.chat.id, 'Введите ваш телефон', reply_markup=dele)

    bot.register_next_step_handler(message, get_phone)


@bot.message_handler(content_types=['text'])
def get_phone(message):
    global phone
    phone = message.text
    markup = types.ReplyKeyboardMarkup()
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    try:
        my_number = phonenumbers.parse(message.text, "RU")
        phonenumbers.is_valid_number(my_number)
        SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SAMPLE_RANGE_NAME = 'Процедуры!A1:B'
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        if not values:
            return
        for row in values:
            markup.add(types.KeyboardButton('%s' % row[0]))
        bot.send_message(message.chat.id, f'Выберите процедуру',
                         reply_markup=markup)
        bot.register_next_step_handler(message, text_save3)
    except phonenumbers.phonenumberutil.NumberParseException:
        bot.send_message(message.from_user.id, 'Введен некорректный номер\nВведите ваш телефон')
        bot.register_next_step_handler(message, get_phone)


@bot.message_handler(content_types=['text'])
def text_save3(message):
    global process
    process = message.text
    dele = telebot.types.ReplyKeyboardRemove()
    bot.send_message(message.chat.id, 'Проверяю наличие свободного времени...', reply_markup=dele)
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    markup = types.ReplyKeyboardMarkup()
    SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    SAMPLE_RANGE_NAME = f'{name_list}!B3:ZZ'
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        days = values[0]
        if not values:
            return
        for i in range(30):
            try:
                markup.add(types.KeyboardButton(days[i]))
            except IndexError:
                break
    except HttpError as err:
        bot.send_message(message.chat.id, err)
    bot.send_message(message.chat.id, f'Свободные дни. Выберите для записи',
                     reply_markup=markup)
    bot.register_next_step_handler(message, text_save4)
    global back2
    back2 = 1


@bot.message_handler(content_types=['text'])
def text_save4(message):
    global day
    global k
    day = message.text
    bot.send_message(message.chat.id, f'Вы выбрали {message.text}. Проверяю наличие свободного времени для записи.')
    SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    SAMPLE_RANGE_NAME = f'{name_list}!B3:ZZ'
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        days =values[0]
        if not values:
            return
        for i in range(30):
            try:
                if days[i] == message.text:
                    k = i + 1
            except IndexError:
                break
    except HttpError as err:
        bot.send_message(message.chat.id, err)
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)
    SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    SAMPLE_RANGE_NAME = f'{name_list}!4:100'
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        b = []
        d = []
        for row in values:
            try:
                d.append(row[0])
                if row[k] != '':
                    b.append(row[0])
            except IndexError:
                continue
        for i in range(len(b)):
            d.remove(b[i])
        for i in range(len(d)):
            markup.add(types.KeyboardButton(d[i]))
    except HttpError as err:
        bot.send_message(message.chat.id, err)
    markup.add(types.KeyboardButton('Нaзaд'))
    bot.send_message(message.chat.id, 'Выберите время для записи', reply_markup=markup)
    global back
    back = 1
    bot.register_next_step_handler(message, text_save5)



@bot.message_handler(content_types=['text'])
def text_save5(message):
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    if message.text == 'Нaзaд' and back == 1:
        dele = telebot.types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, 'Проверяю наличие свободного времени...', reply_markup=dele)
        if message.text == '/start':
            bot.register_next_step_handler(message, start)
        markup = types.ReplyKeyboardMarkup()
        SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        SAMPLE_RANGE_NAME = f'{name_list}!B3:ZZ'
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                        range=SAMPLE_RANGE_NAME).execute()
            values = result.get('values', [])

            if not values:
                return
            for i in range(30):
                for row in values:
                    try:
                        markup.add(types.KeyboardButton('%s' % row[i]))

                    except IndexError:
                        break
        except HttpError as err:
            bot.send_message(message.chat.id, err)
        bot.send_message(message.chat.id, f'Свободные дни. Выберите для записи',
                         reply_markup=markup)
        bot.register_next_step_handler(message, text_save4)
        global back2
        back2 = 1
    else:
        global tim
        global t

        tim = message.text
        SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        SAMPLE_RANGE_NAME = f'{name_list}!4:100'
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                        range=SAMPLE_RANGE_NAME).execute()
            values = result.get('values', [])
            q = 0
            if not values:
                return
            for row in values:
                q += 1
                try:
                    if ('%s' % row[0]) == message.text:
                        t = q + 3
                except IndexError:
                    break
        except HttpError as err:
            bot.send_message(message.chat.id, err)
        markup = types.ReplyKeyboardMarkup()
        markup.add(types.KeyboardButton('Подтвердить запись'))
        markup.add(types.KeyboardButton('Изменить врeмя'))
        bot.send_message(message.chat.id, f'Вы выбрали {day} в {tim}', reply_markup=markup)
        bot.register_next_step_handler(message, text_save6)
        global back1

        back1 = 2


def text_save6(message):
    global t
    global k
    if message.text == '/start':
        bot.register_next_step_handler(message, start)
    if message.text == 'Подтвердить запись':
        gs = GoogleSheet()
        k+=1
        range = f'{name_list}!{t}:{chr(ord("@")+k)}{t}'
        values = [
            [f"{name}\n{phone}\n{process}\n{day} {tim}"]
        ]
        gs.updateRangeValues(range, values)
        markup = types.ReplyKeyboardMarkup()
        markup.add(types.KeyboardButton('Удалить запись'))
        markup.add(types.KeyboardButton('Новая запись'))
        bot.send_message(message.chat.id,
                         f'Спасибо, что выбрали нас! Ждем Вас в {day} в {tim} Телефон для связи.',
                         reply_markup=markup)
        day1 = str(day).split()
        dat = day1[1].split('.')
        tim1 = tim.split(':')
        th = Thread(target=delete)
        th.start()
        bot.register_next_step_handler(message, text_save7)

    else:
        if message.text == 'Изменить врeмя':
            dele = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id, 'Проверяю наличие свободного времени...', reply_markup=dele)
            markup = types.ReplyKeyboardMarkup()
            SAMPLE_SPREADSHEET_ID = '1bJzkrUXcIgHZo-gGOPfYg9NhJnqGoOLHdtUG_n1vBQE'
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

            creds = None
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            SAMPLE_RANGE_NAME = f'{name_list}!B3:ZZ'
            try:
                service = build('sheets', 'v4', credentials=creds)
                sheet = service.spreadsheets()
                result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                            range=SAMPLE_RANGE_NAME).execute()
                values = result.get('values', [])
                p = 0
                if not values:
                    return
                while p != 30:
                    for row in values:
                        try:
                            markup.add(types.KeyboardButton('%s' % row[p]))

                        except IndexError:
                            break
                    p+=1
            except HttpError as err:
                bot.send_message(message.chat.id, err)
            bot.send_message(message.chat.id, f'Свободные дни. Выберите для записи',
                             reply_markup=markup)
            bot.register_next_step_handler(message, text_save4)
            global back2
            back2 = 1
        else:
            if message.text == 'Нaзaд' and back1 == 2:
                bot.register_next_step_handler(message, text_save4)
            else:
                bot.send_message(message.chat.id, 'Я вас не понимаю')
                bot.register_next_step_handler(message, text_save6)





@bot.message_handler(content_types=['text'])
def text_save7(message):
    if message.text == "Удалить запись":
        global t
        global k
        dele = telebot.types.ReplyKeyboardRemove()
        gs = GoogleSheet()
        range = f'{name_list}!{t}:{chr(ord("@") + k)}{t}'
        values = [
            [""]
        ]
        gs.updateRangeValues(range, values)
        bot.send_message(message.chat.id, 'Запись удалена. Чтобы сделать новую запись напишите команду /start',
                         reply_markup=dele)
    else:
        if message.text == 'Новая запись':
            dele = telebot.types.ReplyKeyboardRemove()
            bot.send_message(message.chat.id,
                             'Пожалуйста введите ваше имя', reply_markup=dele)
            bot.register_next_step_handler(message, get_name)
        else:
            bot.send_message(message.chat.id, 'Я вас не понимаю. Чтобы сделать новую запись напишите команду /start')
            bot.register_next_step_handler(message, text_save7)


bot.polling(none_stop=True)
