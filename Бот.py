import telebot
import apiclient.discovery
import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery

# Цвета, используемые в оформлении таблицы
colorLightBlue = {
                'red' : 207/255,
                'green' : 226/255,
                'blue' : 243/255 
                }

colorWhite = {
            'red' : 1,
            'green' : 1,
            'blue' : 1
            }

# Авторизуемся
def authorization():
    global service, httpAuth

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    # Достаём ключ из скачанного json файла
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
            '/Users/maksimkosnikov/Downloads/wide-origin-343614-8f3995911d78.json',
             scope)

    # Авторизуемся в системе
    httpAuth = credentials.authorize(httplib2.Http()) 
    service = discovery.build('sheets', 'v4', credentials=credentials)

    makeSpreadsheet()

# Создаём таблицу
def makeSpreadsheet():
    global spreadsheetId, link
    body = {
        'properties': {'title': 'Прогноз вашего запуска', 'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист номер один',
                                   'gridProperties': {'rowCount': 100, 'columnCount': 9}}}]}
    
    
    spreadsheet = service.spreadsheets().create(
        body = body
        ).execute()

    spreadsheetId = spreadsheet['spreadsheetId'] # Сохраняем идентификатор файла
    link = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId

    access()

# Открываем доступ к созданной таблице
def access():
    # Выбираем работу с Google Drive и 3 версию API
    driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth)

    body = {'type': 'user', 'role': 'writer', 'emailAddress': answers[4]}

    # Открываем доступ на редактирование
    access = driveService.permissions().create(
        fileId = spreadsheetId,
        body = body,
        fields = 'id',
        ).execute()
    
    copySheet()

# Копируем шаблон в таблицу
def copySheet():
    body = {
        'destination_spreadsheet_id': spreadsheetId
    }

    result = service.spreadsheets().sheets().copyTo(
        # ID книги, откуда берётся лист для копирования
        spreadsheetId = '1smNUWjYa6tAQYPQaMeSZLLdCEzTlgVXHVaLHRZ8terQ',
        # ID листа, который копируется в новую книгу
        sheetId = 269971183,
        body=body).execute()

    deleteFirstSheet()

# Удаляем первый лист
def deleteFirstSheet():
    body = {
        'requests': [
            {
            'deleteSheet': {
                'sheetId': 0
                }
            }
        ]
    }

    result = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId, body=body).execute()

    rename()

# Переименовываем скопированный лист
def rename():
    global sheet_id

    sheet_metadata = service.spreadsheets().get(
        spreadsheetId=spreadsheetId).execute()     
    sheets = sheet_metadata.get('sheets', '')
    sheet_id = sheets[0].get('properties', {}).get('sheetId', 0)

    body = {
        'requests': {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': sheet_id,
                    'title': 'Консультация (для эксперта/фрилансера)'
                },
                'fields': 'title',
            }
        }
    }

    result = service.spreadsheets().batchUpdate(
        spreadsheetId = spreadsheetId, body=body).execute()

    updateAdBudget()

# Добавляем значения бюджета рекламы
def updateAdBudget():
    body = {
        'values': [[answers[1]]]
    }

    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheetId,
        range = 'Консультация (для эксперта/фрилансера)!D6',
        valueInputOption = 'RAW', body=body).execute()

    makeCVArray()


percent = 100
numberOfPayments = 0


# Генерируем значения для этапов воронки
def makeCVArray():
    global values, percent, numberOfPayments
    values = []
    k = 1 - percent / int(answers[3]) / 100

    for i in range(1, int(answers[3])+1): #создаём поля с CV
        percent = percent * k
        if i == 1:
            values.append(['CV' + str(i) + '(стоимость подписчика)', '', '',
                          str(answers[2])])
            values.append(['Кол-во переходов на 1 этап воронки', '', '',
                          '=D' + str(6 + len(values) - 1) + '/D' + str(6 + len(values))])
        elif i == int(answers[3]) - 1: 
            values.append(['CV' + str(i), '', '', str(int(percent)) + '%'])
            values.append(['Кол-во переходов на КЭВ', '', '',
                          '=D' + str(6 + len(values) - 1) + '*D' + str(6 + len(values))])
        elif i == int(answers[3]):
            values.append(['CV' + str(i) + ' (конверсия в оплату)', '', '',
                          str(int(percent)) + '%'])
            values.append(['Кол-во оплат', '', '', 
                          '=D' + str(6 + len(values) - 1) + '*D' + str(6 + len(values))])
            numberOfPayments = 6 + len(values)
        else: 
            values.append(['CV' + str(i), '', '', str(int(percent)) + '%'])
            values.append(['Кол-во переходов на ' + str(i) + ' этап воронки', '', '', 
                          '=D' + str(6 + len(values) - 1) + '*D' + str(6 + len(values))])

    addCV()

# Добавляем сгенерированные значения этапов воронки в таблицу 
def addCV():
    body = {'values': values }

    resp = service.spreadsheets().values().append(
        spreadsheetId=spreadsheetId,
        range = 'Консультация (для эксперта/фрилансера)!A7',
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body = body).execute()

    updatePercentFormat()

# Обновляем формат строк на процентный
def updatePercentFormat():
    for i in list(range(0, int(answers[3])+2, 2)):
        body = {
            'requests': [
                {
                    'repeatCell': {
                        'range': {'sheetId': sheet_id,
                            'startRowIndex': 8 + i,
                            'endRowIndex': 9 + i,
                            'startColumnIndex': 2,
                            'endColumnIndex' : 4
                            },
                        'cell': {
                            'userEnteredFormat': {
                                'numberFormat': {
                                    'type': 'PERCENT'
                                }
                            },
                        },
                        'fields': 'userEnteredFormat.numberFormat'
                    }
                }
            ]
        }

        resp = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheetId, body=body).execute()

    frame()

# Генерируем стиль и размер границы для этапов воронки
def frame():
    global style, frame_range

    style = {'style': 'SOLID',          # Сплошная линия
            'width': 1,         # Шириной 1 пиксель
            'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}}     # Чёрный цвет

    frame_range = {'sheetId': sheet_id,
                'startRowIndex': 4,
                'endRowIndex': numberOfPayments,
                'startColumnIndex': 0,
                'endColumnIndex': 4}

    updateFrame()

# Добавляем границу в лист
def updateFrame():
    global style, frame_range

    body = {
        'requests': [
            {'updateBorders': {'range': frame_range,
                   'bottom': style,
                   'top': style,
                   'left': style,
                   'right': style,
                   'innerHorizontal': style,
                   'innerVertical': style
                   }                   
            }
        ]
    }

    resp = service.spreadsheets().batchUpdate(
        spreadsheetId = spreadsheetId, body=body).execute()

    deleteLightYellow()

# Удаляем случайно созданные жёлтые полоски в этапах воронки
def deleteLightYellow():
    global colorWhite

    for i in list(range(0, int(answers[3])+6, 2)):
        body = {
            'requests': [
                {
                    'repeatCell': {
                        'range': {'sheetId': sheet_id,
                            'startRowIndex': 7 + i,
                            'endRowIndex': 8 + i,
                            'startColumnIndex': 2,
                            'endColumnIndex' : 4
                            },
                        'cell': {
                            'userEnteredFormat': {
                                'backgroundColor': colorWhite,
                                'numberFormat': {
                                    'type': 'NUMBER',
                                    'pattern': '#'
                                },     
                            },
                        },
                        'fields': 'userEnteredFormat(backgroundColor, numberFormat)'
                    }
                }
            ]
        }

        resp = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheetId, body=body).execute()

    createLightBlue()

# Добавляем синие полоски в этапы воронки
def createLightBlue():
    for i in list(range(0, int(answers[3])*2, 2)):
        body = {
            'requests': [
                {
                    'repeatCell': {
                        'range': {'sheetId': sheet_id,
                            'startRowIndex': 6 + i,
                            'endRowIndex': 7 + i,
                            'startColumnIndex': 0,
                            'endColumnIndex' : 1
                            },
                        'cell': {
                            'userEnteredFormat': {
                                'backgroundColor': colorLightBlue,
                            },
                        },
                        'fields': 'userEnteredFormat.backgroundColor'
                    }
                }
            ]
        }

        resp = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheetId, body=body).execute()

    updateProductStandart()

# Добавляем среднюю цену продукта
def updateProductStandart():
    body = {'values': [[str(answers[0])]]}

    resp = service.spreadsheets().values().update(
        spreadsheetId=spreadsheetId,
        range = 'Консультация (для эксперта/фрилансера)!D' + str(numberOfPayments + 21),
        valueInputOption = 'USER_ENTERED',
        body = body).execute()

    getProfit()

# Получаем значение прибыли из таблицы
def getProfit():
    global profit

    profit = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range='Консультация (для эксперта/фрилансера)!D2').execute()


bot = telebot.TeleBot('5133517742:AAEuEzGu1isutmdFEUOGPpMvJRzwRMLi2L8')

help = 'Ответами на все вопросы, кроме почты, ' \
       'являются числа без запятых, точек, пробелов и иных символов. ' \
       'Почта должна оканчивать на gmail.com, ' \
       'иначе бот будет работать некорректно, ' \
       'и все данные придётся вводить заново'


# Создаем команды /start и /help
@bot.message_handler(content_types=['text'])
def firstQuestion(message):
    if message.text == '/start':
        global answers
        answers = []

        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта.')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, firstQuestion)
    else:
        send = bot.send_message(message.chat.id, 'Введите /start или /help')
        bot.register_next_step_handler(send, firstQuestion)


# Получаем ответ на первый вопрос и задаём второй
@bot.message_handler(content_types=['text'])
def secondQuestion(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта.')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, secondQuestion)
    else: 
        firstAnswer = message.text

        try:
            number = int(message.text)
            answers.append(firstAnswer)
            send = bot.send_message(message.chat.id, 'Введите бюджет на рекламу.')
            bot.register_next_step_handler(send, thirdQuestion)
        except:
            send = bot.send_message(message.chat.id,
                                    'Ошибка\nВведите среднюю цену продукта.')
            bot.register_next_step_handler(send, secondQuestion)

# Получаем ответ на второй вопрос и задаём третий
@bot.message_handler(content_types=['text'])
def thirdQuestion(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта.')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, thirdQuestion)
    else:
        secondAnswer = message.text
            
        try:
            number = int(message.text)
            answers.append(secondAnswer)
            send = bot.send_message(message.chat.id,
                                    'Введите среднюю стоимость клиента за вход в воронку.')
            bot.register_next_step_handler(send, fourthQuestion)
        except:
            send = bot.send_message(message.chat.id, 'Ошибка\nВведите бюджет на рекламу.')
            bot.register_next_step_handler(send, thirdQuestion)

# Получаем ответ на третий вопрос и задаём четвёртый
@bot.message_handler(content_types=['text'])
def fourthQuestion(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта.')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, fourthQuestion)
    else:
        thirdAnswer = message.text

        try:
            number = int(message.text)
            answers.append(thirdAnswer)
            send = bot.send_message(message.chat.id, 'Введите количество этапов воронки.')
            bot.register_next_step_handler(send, fifthQuestion)
        except:
            send = bot.send_message(message.chat.id,
                                    'Ошибка\nВведите среднюю стоимость клиента за вход в воронку.')
            bot.register_next_step_handler(send, fourthQuestion)
        

# Получаем ответ на четвёртый вопрос и задаём пятый
@bot.message_handler(content_types=['text'])
def fifthQuestion(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта.')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, fifthQuestion)
    else:
        fourthAnswer = message.text

        if message.text.isdigit() and int(message.text) >= 3:
            answers.append(fourthAnswer)
            send = bot.send_message(message.chat.id, 'Введите вашу google почту.')
            bot.register_next_step_handler(send, end)
        else:
            send = bot.send_message(message.chat.id, 
                'Ошибка\nКорректно введите количество этапов воронки.')
            bot.register_next_step_handler(send, fifthQuestion)

# Получаем ответ на последний вопрос и выводим результат
@bot.message_handler(content_types=['text'])
def end(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, end)
    else:
        fifthAnswer = message.text

        answers.append(fifthAnswer)
        print(answers)
        bot.send_message(message.chat.id, 'Пожалуйста, подождите. Создаём таблицу.'
                         'Это займёт 1-2 минуты')
        bot.register_next_step_handler(message, after_end)
        authorization()
        bot.send_message(message.chat.id, str(profit['values'][0][0]) + 
                         ' - это финансовый потенциал вашего запуска: примерное значение выручки,'
                         'которую можно получить.\n\nВ каждой нише и в каждой ситуации все индивидуально,'
                         'поэтому для более детального результата вы можете скачать таблицу с декомпозицией'
                         'запуска, внести туда свои данные и цифры из вашего конкретного случая\n\n'
                         'Ссылка на скачивание таблицы:\n' + link)

# Проверяем, написал ли что-либо пользователь после вывода ботом результата
@bot.message_handler(content_types=['text'])
def after_end(message):
    if message.text == '/start':
        send = bot.send_message(message.chat.id, 'Введите среднюю цену продукта')
        bot.register_next_step_handler(send, secondQuestion)
    elif message.text == '/help':
        send = bot.send_message(message.chat.id, help)
        bot.register_next_step_handler(send, after_end)
    else:
        send = bot.send_message(message.chat.id, 
                                'Похоже, ты начал заново вводить данные, не буду мешать. Введи /start или /help')
        bot.register_next_step_handler(send, after_end)


bot.polling()