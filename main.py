import openpyxl
import telebot
from openpyxl import Workbook
from telebot import types

tokenbot = telebot.TeleBot('5567917612:AAFqPqeKOSVuDYWRepZVaZ5hS9aUKpACmCc')

start_text = '''Здравствуйте.
Кнопки записывают текст в соответсвующие ячейки.  
После выбора поля - введите текст одним сообщением.
Будьте кратки и нформативны, либо нажмтите /help что бы увидеть пример заполнения. 
Что бы начать нажмите /go '''

help_text = '''Пример заполнения таблицы.

(A) Ситуация 
Муж повысил голос из-за того, что я долго 
собиралась на работу и ему пришлось 
меня ждать.

(B) Мысли
Он меня не любит.
Он меня ударит.

(С) Эмоция и её сила 
Обида – 80
Страх - 65

(D) Ощущения в теле
Ком в горле, напряжение лица. Боль в 
солнечном сплетении, напряжение 
в плечах.

(E) Поведение
Молча ускорилась.
Молчала всю дорогу до работы.
'''

#   Создание файла xlsx
book: Workbook = openpyxl.Workbook()
sheet = book.active

'''Номера ячеек по умолчанию'''
column_num = 1
row_num = 3


#   Приветствие и запуск
@tokenbot.message_handler(commands=['start'])
def start(start):
    tokenbot.reply_to(start, start_text)


@tokenbot.message_handler(commands=['help'])
def start(help):
    tokenbot.reply_to(help, help_text)


#   Кнопки управления
@tokenbot.message_handler(commands=['go'])
def button(go):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)
    chapter = types.KeyboardButton('Hовая запись')
    a = types.KeyboardButton('(A) Ситуация')
    b = types.KeyboardButton('(B) Мысли')
    c = types.KeyboardButton('(С) Эмоция и её сила')
    d = types.KeyboardButton('(D) Ощущения в теле')
    e = types.KeyboardButton('(E) Поведение')
    markup.add(chapter, a, b, c, d, e)
    tokenbot.send_message(go.chat.id, 'Кнопки сформированы\nНажмите кнопку «Ситауция» и расскажите что произошло',
                          reply_markup=markup)


@tokenbot.message_handler(content_types=['text'])
def tap_button(tap):
    global column_num
    global row_num
    if tap.text == 'Hовая запись':
        tokenbot.register_next_step_handler(tap, user_mess)
        tokenbot.send_message(tap.chat.id, 'Теперь кнопку «Ситуация» нажимать не надо - вводите текст')
        column_num = 1
        row_num += 1
    if tap.text == '(A) Ситуация':
        tokenbot.register_next_step_handler(tap, user_mess)
    if tap.text == '(B) Мысли':
        tokenbot.register_next_step_handler(tap, user_mess)
        column_num += 1
    if tap.text == '(С) Эмоция и её сила':
        tokenbot.register_next_step_handler(tap, user_mess)
        column_num += 1
    if tap.text == '(D) Ощущения в теле':
        tokenbot.register_next_step_handler(tap, user_mess)
        column_num += 1
    if tap.text == '(E) Поведение':
        tokenbot.register_next_step_handler(tap, user_mess)
        column_num += 1


def user_mess(ms):
    sheet.cell(column=column_num, row=row_num).value = ms.text
    # sheet['A4'] = ms.text

    book.save('my_book.xlsx')
    book.close()


tokenbot.polling(none_stop=True)
