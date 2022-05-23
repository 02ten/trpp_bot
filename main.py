"""
This is main package
"""

import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType, VkMessageFlag
from vk_api.utils import get_random_id
import requests
from bs4 import BeautifulSoup
import xlrd
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
import re
import datetime
from datetime import datetime, timedelta

months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября',
          'декабря', ]
weeks = ['понедельник', 'вторник', 'среду', 'четверг', 'пятницу', 'субботу', ]

start_year = 21
keyboard = VkKeyboard(one_time=True)  # Клавиатура бота
keyboard.add_button('На сегодня', color=VkKeyboardColor.POSITIVE)
keyboard.add_button('На завтра', color=VkKeyboardColor.NEGATIVE)
keyboard.add_line()  # переход на вторую строку
keyboard.add_button('На эту неделю')
keyboard.add_button('На следующую неделю')
keyboard.add_line()
keyboard.add_button('Какая неделя')
keyboard.add_button('Какая группа')


def parsing():
    """Скачивает расписание с сайта МИРЭА
    """
    page = requests.get("https://www.mirea.ru/schedule/")
    soup = BeautifulSoup(page.text, "html.parser")
    result = soup.find("div", {"class": "rasspisanie"}). \
        find(string="Институт информационных технологий"). \
        find_parent("div"). \
        find_parent("div"). \
        findAll("a", {"class": "uk-link-toggle"})
    j = 0
    for x in result:
        s = (result[j].find("div", {"class": "uk-link-heading"}).get_text())
        if x.find("div", {"class": "uk-link-heading"}).get_text() == s:
            f = open(str(j) + '.xlsx', "wb")
            resp = requests.get(x.get("href"))
            f.write(resp.content)
        j += 1



def add_to_raspisanie(sheet, start, stop, i):

    raspisanie = []
    for j in range(start, stop, 2):
        raspisanie.append(
            sheet.cell(j, i).value + '\n' + sheet.cell(j, i + 1).value + '\n' + sheet.cell(j, i + 2).value)
    return raspisanie



def parsing_exel_by_day(group, date):
    """Извлечение данных из файла с расписанием по дню
    :param group: группа для которого нужно расписание
    :param date: дата на какое нам необходимо расписание
    :return: возвращает расписание
    """
    course = int(group[1].split('-')[2])
    course = start_year - course
    today_week = (get_week().days // 7 + 1) % 2
    book = xlrd.open_workbook(str(course) + '.xlsx')
    sheet = book.sheet_by_index(0)
    num_cols = sheet.ncols
    group[1] = group[1].upper().rstrip()
    raspisanie = []
    for i in range(num_cols):
        if (sheet.cell(1, i).value == group[1]):
            if (today_week == 1):
                if (date.weekday() == 0):
                    raspisanie = add_to_raspisanie(sheet, 3, 14, i)
                if (date.weekday() == 1):
                    raspisanie = add_to_raspisanie(sheet, 15, 26, i)
                if (date.weekday() == 2):
                    raspisanie = add_to_raspisanie(sheet, 27, 38, i)
                if (date.weekday() == 3):
                    raspisanie = add_to_raspisanie(sheet, 39, 50, i)
                if (date.weekday() == 4):
                    raspisanie = add_to_raspisanie(sheet, 51, 62, i)
                if (date.weekday() == 5):
                    raspisanie = add_to_raspisanie(sheet, 63, 74, i)
            else:
                if (date.weekday() == 0):
                    raspisanie = add_to_raspisanie(sheet, 4, 15, i)
                if (date.weekday() == 1):
                    raspisanie = add_to_raspisanie(sheet, 16, 27, i)
                if (date.weekday() == 2):
                    raspisanie = add_to_raspisanie(sheet, 28, 39, i)
                if (date.weekday() == 3):
                    raspisanie = add_to_raspisanie(sheet, 40, 51, i)
                if (date.weekday() == 4):
                    raspisanie = add_to_raspisanie(sheet, 52, 63, i)
                if (date.weekday() == 5):
                    raspisanie = add_to_raspisanie(sheet, 64, 75, i)
    print(raspisanie)
    return raspisanie


def parsing_exel_by_week_day(group, weekday):
    """Извлечение данных из файла с расписанием по дню недели
    :param group: какой группе нужно расписание
    :param weekday: на какой день недели нужно расписание
    :return:
    """
    course = int(group[1].split('-')[2])
    course = start_year - course
    book = xlrd.open_workbook(str(course) + '.xlsx')
    sheet = book.sheet_by_index(0)
    num_cols = sheet.ncols
    group[1] = group[1].upper().rstrip()

    for i in range(num_cols):
        if (sheet.cell(1, i).value == group[1]):
            if weekday == 0:
                raspisanie = add_to_raspisanie(sheet, 3, 14, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 4, 15, i)
                return raspisanie
            if weekday == 1:
                raspisanie = add_to_raspisanie(sheet, 15, 26, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 16, 27, i)
                return raspisanie
            if weekday == 2:
                raspisanie = add_to_raspisanie(sheet, 27, 38, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 28, 39, i)
                return raspisanie
            if weekday == 3:
                raspisanie = add_to_raspisanie(sheet, 39, 50, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 40, 51, i)
                return raspisanie
            if weekday == 4:
                raspisanie = add_to_raspisanie(sheet, 51, 62, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 52, 63, i)
                return raspisanie
            if weekday == 5:
                raspisanie = add_to_raspisanie(sheet, 63, 74, i)
                raspisanie = raspisanie + add_to_raspisanie(sheet, 64, 75, i)
                return raspisanie



def parsing_exel_by_week(group, week):
    """Извлечение данных из файла с расписанием по неделе
    :param group: группа для которого необходимо расписание
    :param week: на какую неделю нужно расписание
    """
    course = int(group[1].split('-')[2])
    course = start_year - course
    if (week == 0):
        today_week = (get_week().days // 7 + 1) % 2
    elif (week == 1):
        today_week = (get_week().days // 7 + 2) % 2
    book = xlrd.open_workbook(str(course) + '.xlsx')
    sheet = book.sheet_by_index(0)
    num_cols = sheet.ncols
    group[1] = group[1].upper().rstrip()
    raspisanie = []
    for i in range(num_cols):
        if (sheet.cell(1, i).value == group[1]):
            if (today_week == 1):
                raspisanie = add_to_raspisanie(sheet, 3, 74, i)
            else:
                raspisanie = add_to_raspisanie(sheet, 4, 75, i)
    return raspisanie


def add_group(number, group):
    """Добавляет группу пользователю
    :param number: пользователь которому нужно заменить группу
    :param group: группа на которую нужно заменить
    """
    with open('users.txt') as f:
        lines = f.readlines()
    index = lines.index(group[0])
    lines[index] = lines[index].split('\n')[0]
    lines[index] += ' ' + number + '\n'
    with open('users.txt', 'w') as f:
        for line in lines:
            f.write(line)


def replace_group(number, group):
    """Заменяет группу пользователю
    :param number: пользователь которому нужно заменить группу
    :param group: группа на которую нужно заменить
    """
    with open('users.txt') as f:
        lines = f.readlines()
    s = ' '.join(group)
    index = lines.index(s)
    temp = lines[index].split(' ')
    temp[1] = number + '\n'
    temp = ' '.join(temp)
    lines[index] = temp
    with open('users.txt', 'w') as f:
        for line in lines:
            f.write(line)


def get_week():
    """Вспомогательная функция для подсчета текущей недели
    """
    start_date = datetime(2022, 2, 7)
    todayWeek = (datetime.now() - start_date)
    return todayWeek


def get_group(user_id):
    """Вспомогательная функция для получения группы пользователя
    :param user_id: текущий пользователь
    """
    f = open('users.txt', 'r')
    for line in f:
        if str(user_id) in line:
            group = line.split(' ')
    f.close()
    return group


def get_flag(user_id):
    flag = False
    f = open('users.txt', 'r')
    for line in f:
        if str(user_id) in line:
            flag = True
    f.close()
    return flag


def print_raspisanie_by_week(vk, event, week, raspisanie):
    """Выводит расписание боту по неделе
    :param vk: Использует vk_bot для отправки сообщений
    :param event: Распозноет собитие которое произошло
    :param week: Неделя, на которое нам необходимо расписание
    :param raspisanie: Список, в котором передается расписание
    """
    current_week = (get_week().days + week) // 7 + 1
    vk.messages.send(
        user_id=event.user_id,
        random_id=get_random_id(),
        message='Расписание на ' + str(current_week + week) + ' неделю'
    )
    print(raspisanie)
    result = ''
    for i in range(6):
        if (len(raspisanie) != 0):
            result = result + 'Расписание на ' + weeks[i] + '. ' + str(current_week + week) + ' недели\n'
            s = 0
            for j in range(i * 6, 6 * (i + 1)):
                s += 1
                x = raspisanie[j].splitlines()
                x1 = ''
                if len(raspisanie[j]) == 2:
                    x1 = '-'
                else:
                    x1 = ', '.join(x)
                result = result + str(s) + ')' + x1 + '\n'

        else:
            vk.messages.send(
                user_id=event.user_id,
                random_id=get_random_id(),
                message='Сегодня выходной\n'
            )
        result = result + '\n'
    vk.messages.send(
        user_id=event.user_id,
        random_id=get_random_id(),
        keyboard=keyboard.get_keyboard(),
        message=result
    )


def print_raspisanie_by_day(vk, event, date, raspisanie):
    """Выводит расписание боту по дню
    :param vk: Использует vk_bot для отправки сообщений
    :param event: Распозноет собитие которое произошло
    :param week: Неделя, на которое нам необходимо расписание
    :param raspisanie: Список, в котором передается расписание
    """
    if (len(raspisanie) != 0):
        vk.messages.send(
            user_id=event.user_id,
            random_id=get_random_id(),
            message='Расписание на ' + str(date.day) + ' ' + months[date.month - 1] + '\n'
        )
        result = ''
        for i in range(6):
            x = raspisanie[i].splitlines()
            x1 = ''
            if (len(raspisanie[i]) != 2):
                x1 = ', '.join(x)
            else:
                x1 = '-'
            result = result + str(i + 1) + ')' + x1 + '\n'
        vk.messages.send(
            user_id=event.user_id,
            random_id=get_random_id(),
            keyboard=keyboard.get_keyboard(),
            message=result
        )
    else:
        vk.messages.send(
            user_id=event.user_id,
            random_id=get_random_id(),
            keyboard=keyboard.get_keyboard(),
            message='Сегодня выходной\n'
        )


def print_raspisanie_by_week_day(vk, event, weekday, raspisanie):
    """Выводит расписание боту по дню недели
    :param vk: Использует vk_bot для отправки сообщений
    :param event: Распозноет собитие которое произошло
    :param week: Неделя, на которое нам необходимо расписание
    :param raspisanie: Список, в котором передается расписание
    """
    flag = True
    result = ''
    print(len(raspisanie))
    print(weekday)
    for i in range(2):
        if flag:
            if (weekday == 2 or weekday == 4 or weekday == 5):
                result = result + 'Расписание на нечетную ' + weeks[weekday] + '\n'
            else:
                result = result + 'Расписание на нечетный ' + weeks[weekday] + '\n'
            flag = False
        else:
            if (weekday == 2 or weekday == 4 or weekday == 5):
                result = result + '\nРасписание на четную ' + weeks[weekday] + '\n'
            else:
                result = result + '\nРасписание на четный ' + weeks[weekday] + '\n'
        s = 0
        for j in range(i * 6, 6 * (i + 1)):
            s += 1
            x = raspisanie[j].splitlines()
            x1 = ''
            if len(raspisanie[j]) == 2:
                x1 = '-'
            else:
                x1 = ', '.join(x)
            result = result + str(s) + ')' + x1 + '\n'
    vk.messages.send(
        user_id=event.user_id,
        random_id=get_random_id(),
        keyboard=keyboard.get_keyboard(),
        message=result
    )


def bot():
    """Основная функция для бота
    """
    vk_session = vk_api.VkApi(
        token='2fd4d0ee9475deb74c870e3c055a2e8381d4eb396ef2875473d6f1776cac7ab4d1c6e831bd06af01fafbe')
    vk = vk_session.get_api()

    longpoll = VkLongPoll(vk_session)
    start = VkKeyboard(one_time=True)
    start.add_button("Начать", color=VkKeyboardColor.POSITIVE)

    diff = False

    for event in longpoll.listen():
        if event.type == VkEventType.MESSAGE_NEW:
            if event.to_me:

                flag = get_flag(event.user_id)

                if (flag):

                    if event.type == VkEventType.MESSAGE_NEW and len(get_group(event.user_id)) == 1:
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            message='Сначала необходимо ввести номер группы, чтобы взаимодействовать с расписанием'
                                    'Введите номер группы в формате XXXX-00-00\n'
                        )
                        group = get_group(event.user_id)
                        number = event.text
                        result = re.match(r'\w{4}-\d{2}-\d{2}', event.text)
                        if (result):
                            vk.messages.send(
                                user_id=event.user_id,
                                random_id=get_random_id(),
                                message='Я запомнил, что ты из группы ' + number
                            )
                            add_group(number, group)
                        else:
                            vk.messages.send(
                                user_id=event.user_id,
                                random_id=get_random_id(),
                                message='Группа введена неккоректно, попробуй еще раз'
                            )

                    elif event.type == VkEventType.MESSAGE_NEW and re.match(r'\w{4}-\d{2}-\d{2}', event.text):
                        replace_group(event.text, get_group(event.user_id))


                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 2 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[1]):
                        diff = True
                        diff_group = event.text.split(' ')[1] + '\n'
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=keyboard.get_keyboard(),
                            message='Показать расписание ' + diff_group
                        )

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == "бот":
                        group = get_group(event.user_id)
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=keyboard.get_keyboard(),
                            message='Показать расписание ' + group[1]
                        )

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'понедельник':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 0, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'вторник':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 1, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'среда':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 2, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'четверг':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 3, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'пятница':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 4, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and len(event.text.split(' ')) == 3 and \
                            event.text.split(' ')[0].lower() == 'бот' and re.match(r'\w{4}-\d{2}-\d{2}',
                                                                                   event.text.split(' ')[2]) and \
                            event.text.split(' ')[1].lower() == 'суббота':
                        group = get_group(event.user_id)
                        group[1] = event.text.split(' ')[2]
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 5, raspisanie)


                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот понедельник':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 0)
                        print_raspisanie_by_week_day(vk, event, 0, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот вторник':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 1)
                        print_raspisanie_by_week_day(vk, event, 1, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот среда':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 2)
                        print_raspisanie_by_week_day(vk, event, 2, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот четверг':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 3)
                        print_raspisanie_by_week_day(vk, event, 3, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот пятница':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 4)
                        print_raspisanie_by_week_day(vk, event, 4, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'бот суббота':
                        group = get_group(event.user_id)
                        raspisanie = parsing_exel_by_week_day(group, 5)
                        print_raspisanie_by_week_day(vk, event, 5, raspisanie)


                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'какая группа':
                        group = get_group(event.user_id)
                        if (diff):
                            group[1] = diff_group
                            diff = False
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=keyboard.get_keyboard(),
                            message='Твоя группа ' + group[1]
                        )

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'какая неделя':
                        todayWeek = get_week()
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=keyboard.get_keyboard(),
                            message='Сейчас идет неделя ' + str(todayWeek.days // 7 + 1)
                        )

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'на сегодня':
                        group = get_group(event.user_id)
                        if (diff):
                            group[1] = diff_group
                            diff = False
                        raspisanie = parsing_exel_by_day(group, datetime.today())
                        print_raspisanie_by_day(vk, event, datetime.today(), raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'на завтра':
                        current_date = datetime.today() + timedelta(1)
                        group = get_group(event.user_id)
                        if (diff):
                            group[1] = diff_group
                            diff = False
                        raspisanie = parsing_exel_by_day(group, current_date)
                        print_raspisanie_by_day(vk, event, current_date, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'на эту неделю':
                        group = get_group(event.user_id)
                        if (diff):
                            group[1] = diff_group
                            diff = False
                        raspisanie = parsing_exel_by_week(group, 0)
                        print_raspisanie_by_week(vk, event, 0, raspisanie)

                    elif event.type == VkEventType.MESSAGE_NEW and event.text.lower() == 'на следующую неделю':
                        group = get_group(event.user_id)
                        if (diff):
                            group[1] = diff_group
                            diff = False
                        raspisanie = parsing_exel_by_week(group, 1)
                        print_raspisanie_by_week(vk, event, 1, raspisanie)

                    else:
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=keyboard.get_keyboard(),
                            message='Неизвестная мне команда'

                        )


                else:
                    vk.messages.send(
                        user_id=event.user_id,
                        random_id=get_random_id(),
                        keyboard=start.get_keyboard(),
                        message='Привет, ' + vk.users.get(user_id=event.user_id)[0]['first_name']

                    )
                    if event.type == VkEventType.MESSAGE_NEW and event.text.lower() == "начать":
                        f = open('users.txt', 'a')
                        f.write(str(event.user_id) + '\n')
                        f.close()
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            message='Чтобы вызвать клавиатуру, напишите "Бот",но сначало нужно, чтобы я запомнил группу, '
                                    'напишите ее в формате XXXX-00-00. '
                                    'Если хотите посмотреть погоду, напишите "Погода"'
                        )
                    else:
                        vk.messages.send(
                            user_id=event.user_id,
                            random_id=get_random_id(),
                            keyboard=start.get_keyboard(),
                            message='Для началы работы с ботом, напиши или нажми кнопку "начать"'

                        )


parsing()
bot()
