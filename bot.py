import telebot, logging, time, yandex_parser_full as yandex_parser_full, os
from telebot import types
token = '5643122315:AAHTc3MOGgFiwXznLctdNPVRZ6QDO_iDSSs'

logging.basicConfig(format='%(asctime)s %(levelname)s - %(message)s',
                    level=logging.INFO, filename='bot.log')

bot = telebot.TeleBot(token)


@bot.message_handler(commands=['start'])
def start(message):
    print('start')
    bot.send_message(message.chat.id, 'Здравствуйте, чтобы получить файл введите ссылку на товар с Яндекс.Маркета')

@bot.message_handler(content_types=['text'])
def message(message):
    if 'market.yandex' in message.text:
        bot.send_message(message.chat.id, 'Запускаю парсинг')
        result = yandex_parser_full.start_parser(message.chat.id, message.text)
        bot.send_message(message.chat.id, f'Результаты парсинга отзывов:{result}')
    else:
        bot.send_message(message.chat.id, 'Попробуйте еще раз! Введите корректную ссылку с Яндекс.Маркета')


while True:
    try:
        bot.polling(none_stop=True)  
    except:
        time.sleep(5)  