import os
import xlrd
import re
import xlwings as xw
import tkinter as tk

"""
Модуль для поиска информации в exel файле счета
Обратить внимание, что иногда распознается с ошибками
Таблица состоит из 10 колонок
"""

def window_start():
    user = input('Добавить счет 1 \n Добавить акт по номеру счета 2')
    if int(user) == 1:
        added_account()
    else:
        added_act()

    # """
    # Окно для выбора заносим счет или добавляем акт
    # """
    # win_user = tk.Tk()
    # win_user.title("Начало")
    # win_user.geometry('600x400+200+100')
    # tk.Label(text=
    #          'Добавить счет? \n'
    #          'Добавить акт по номеру счета',
    #          width=200, height=10, font='14'
    #          ).pack()
    # tk.Label(text='', height=0).pack()
    # button1 = tk.Button(win_user, text='Добавить счет', width=0, height=0, font='12', command=added_account)
    # button2 = tk.Button(win_user, text='Добавить акт по номеру счета', width=0, height=0, font='12', command=added_act)
    # button3 = tk.Button(win_user, text='Закрыть', width=0, height=0, font='12', command=win_user.destroy)
    # button1.pack(side='left')
    # button2.pack(side='right')
    # button3.pack(side='top')
    # win_user.mainloop()
    # return


def window_check_price(account, price):
    """
    Функция выводит окно с распознаной ценой и номером счета
    Args:
        account: номер счета
        price: цена

    Returns:
        Введеное значение пользователем
    """
    win_user_check = tk.Tk()
    win_user_check.title("Проверка")
    win_user_check.geometry('600x400+200+100')
    tk.Label(text=
             'Провертье цену, возможно ошибка при распознание \n СЧЕТ: {0} \n Распознаная цена: {1}\n '
             '\n введите цену если она не соответсвует цене в документе'
             .format(account, price),
             width=200, height=10, font='14'
             ).pack()
    entry_text = tk.StringVar()
    entry = tk.Entry(win_user_check, width=30, textvariable=entry_text)
    entry.insert(0, price)
    entry.pack()
    tk.Label(text='', height=0).pack()
    button1 = tk.Button(win_user_check, text='ok', width=0, height=0, font='12', command=win_user_check.destroy)
    button1.pack()
    win_user_check.mainloop()
    return entry_text.get()

# Поиск номера счета
def get_number_account(item_one_info_in_account_f_Exel):
    number_account = re.findall('(\d+)', item_one_info_in_account_f_Exel)
    # print(number_account[0])
    return number_account[0]


# Поиск даты счета
# Продумать как преобразовать 10 января 2018 в 10.01.2018 ---- СДЕЛАЛ 15082018
''' 15.08.
1) Оптимизировать (сделать код более компактным за счет list compression)
2) Сделать отдельную фунцию на обработку даты (добавить различные варианты для обработки месяца)
3) Продумать, как собирать данные с ошибочным распознанием и применить их при дальнейшем использование программы 
пример: распознал июл1 надо чтобы мы сохранили этот вариант и при следующем таком распознание мы получили 07, а не вариант с возможностью занести инф. самим
'''


# 11082018---------------1603
def get_date_account(item_one_info_in_account_f_Exel):
    date_account = re.findall('(\d\d\s\w\w+\s\d\d\d\d)', item_one_info_in_account_f_Exel)
    if not date_account:
        return False
    for date_account_one in date_account:
        date_account = date_account_one
    date_account = date_account.split(' ')
    if date_account[1] == 'января':
        date_account[1] = '01'
    elif date_account[1] == 'февраля':
        date_account[1] = '02'
    elif date_account[1] == 'марта':
        date_account[1] = '03'
    elif date_account[1] == 'апреля':
        date_account[1] = '04'
    elif date_account[1] == 'мая':
        date_account[1] = '05'
    elif date_account[1] == 'июня':
        date_account[1] = '06'
    elif date_account[1] == 'июля':
        date_account[1] = '07'
    elif date_account[1] == 'августа':
        date_account[1] = '08'
    elif date_account[1] == 'сентября':
        date_account[1] = '09'
    elif date_account[1] == 'октября':
        date_account[1] = '10'
    elif date_account[1] == 'ноября':
        date_account[1] = '11'
    elif date_account[1] == 'декабря':
        date_account[1] = '12'
    else:
        date_account[1] = input(
            'На экране указан нераспознаный месяц, если вы не можете его индефецировать то введите две цифры, '
            'соответсвующие этому месяцу, пример: 02')
    date_account = (str(date_account[0]) + '.' + date_account[1] + '.' + str(date_account[2]))
    return date_account


# Функция возращает список из (номер счета, дата, прибор, сумма с НДС)
def seach_need_info_in_account_f(info_in_account_f_Exel):
    i = 0
    n = 0
    list_name_account_date_nds = []
    all_list_name_account_date_nds = []
    # print(info_in_account_f_Exel)
    for item_info_in_account_f_Exel in info_in_account_f_Exel:
        # print(item_info_in_account_f_Exel)
        for item_one_info_in_account_f_Exel in item_info_in_account_f_Exel:
            if re.search(r'СЧЕТ', item_one_info_in_account_f_Exel):
                i += 1
                number_account = get_number_account(item_one_info_in_account_f_Exel)
                # print('Номер: ' + number_account)
                list_name_account_date_nds.append(number_account)
                date_account = get_date_account(item_one_info_in_account_f_Exel)
                list_name_account_date_nds.append(date_account)
                # print('Дата:' + date_account)
            if re.search(r'ШТ', item_one_info_in_account_f_Exel):
                n += 1
                list_name_account_date_nds.append(item_info_in_account_f_Exel[0])
                list_name_account_date_nds.append(item_info_in_account_f_Exel[7])
                # print(item_info_in_account_f_Exel)
                # print('Прибор: ' + item_info_in_account_f_Exel[0])
                # print('Сумма с НДС: ' + item_info_in_account_f_Exel[7])
    all_list_name_account_date_nds.append(list_name_account_date_nds)
    print('Количество счетов: ' + str(i))
    print('Количество позиций поверки: ' + str(n))
    if n == i:
        print('Отлично в распознаном файле Exel ошибок нет')
    else:
        print('Ошибка просим обратить внимание на распознантый файл в название не правильно распозналось слово СЧЕТ '
              'или в таблицы не верно распознались ШТ')
    return all_list_name_account_date_nds


# Функция открытия основной таблицы Уралтест.xlsx
def open_main_task():
    name_main_task = os.getcwd() + '\\Уралтест.xlsx'
    open_main_task = xlrd.open_workbook(name_main_task)
    sheet_main_task = open_main_task.sheet_by_index(name_sheet)
    # получаем значение первой ячейки A1
    # val = sheet_f_account.row_values(0)[0]
    # print(val)
    # получаем список значений из всех записей
    info_in_main_task_Exel = [sheet_main_task.row_values(rownum) for rownum in range(sheet_main_task.nrows)]
    print(info_in_main_task_Exel)


# Вылавливаем ошибку для нахождения последней строки и вычетаем 1 для продолжения работы программы
# Находим пустую строку в тексте
def get_empty_line_in_table(name_sheet):
    try:
        rb = xlrd.open_workbook('Уралтест.xlsx')
        # выбираем активный лист
        sheet = rb.sheet_by_index(name_sheet)
        for i in range(0, 1000000000000):
            sheet.row_values(i)[0]
    except:
        i = i + 1
        return i


"""Функция генерирует список all_list_name_account_date_nds
по след маске [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
"""


def get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds):
    n = 0
    j = 4
    sort_all_list_name_account_date_nds = []
    count_insert_list = len(all_list_name_account_date_nds[0]) / 4
    # print(count_insert_list)
    # print(len(all_list_name_account_date_nds[0]))
    for i in range(0, int(count_insert_list)):
        sort_all_list_name_account_date_nds.append(all_list_name_account_date_nds[0][n:j])
        n = n + 4
        j = j + 4
    return (sort_all_list_name_account_date_nds)


"""Вставляем значение в файл Уралтест
Счет на оплату
(A)прибор
(E)номер
(F)дата
(G)Сумма с НДС
список вида [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
"""


# Продумать закрытие и сохранение занесеных данных
# 1508 Продумать чтобы данные заносились в ячейки и фильтр их сортировал верно
def add_info_in_main_f(empty_line_in_table, sort_all_list_name_account_date_nds):
    wb = xw.Book('Уралтест.xlsx')
    for i in range(0, len(sort_all_list_name_account_date_nds)):
        xw.Range('A' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][2]
        xw.Range('E' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][0]
        xw.Range('F' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][1]
        '''Добавить модуль тнике для графического сообщения пользователю о не верно распозаной цене'''
        sort_all_list_name_account_date_nds[i][3] = sort_all_list_name_account_date_nds[i][3].replace(' ', '')
        check_price = sort_all_list_name_account_date_nds[i][3].split(',')
        if len(check_price) == 2:
            xw.Range('G' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][3]
        else:
            entry = window_check_price(sort_all_list_name_account_date_nds[i][0], sort_all_list_name_account_date_nds[i][3])
            # user_check = input(
            #     'Провертье цену, возможно ошибка при распознание СЧЕТ:{0} Распознаная цена:{1} '
            #     'и введите цену если она не соответсвует цене в документе'
            #     .format(sort_all_list_name_account_date_nds[i][0], sort_all_list_name_account_date_nds[i][3])
            # )
            if entry:
                xw.Range('G' + str(empty_line_in_table)).value = entry
            else:
                xw.Range('G' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][3]
        empty_line_in_table = empty_line_in_table + 1
    wb.save()
    wb.close()





# ---------------------------------------------------------------------------------------------------------------------


def get_act_number(item_one_info_in_act_f_Exel):
    """
    Ищем номер акта
    Args:
        item_one_info_in_act_f_Exel: строка с номером акта

    Returns:
        номер акта

    """
    number_act = re.findall('(\w\w\w\w[-]\d\d\d\d\d\d)', item_one_info_in_act_f_Exel)
    # print(number_account[0])
    return number_act[0]


# Функция возращает список из (номер счета, дата, прибор, сумма с НДС)
def seach_need_info_in_act_f(info_in_act_f_exel):
    i = 0
    n = 0
    z = 0
    list_name_act_date_nds = []
    all_list_name_act_date_nds = []
    for item_info_in_act_f_exel in info_in_act_f_exel:
        for item_one_info_in_act_f_Exel in item_info_in_act_f_exel:
            if not re.search(r'РА №', str(item_one_info_in_act_f_Exel)) and \
                    (re.search(r'ЕК00', str(item_one_info_in_act_f_Exel)) or
                    re.search(r'№ ЕК', str(item_one_info_in_act_f_Exel))):
                date_account = get_date_account(item_one_info_in_act_f_Exel)
                if date_account:
                    list_name_act_date_nds.append(date_account)
                    act_number = get_act_number(item_one_info_in_act_f_Exel)
                    list_name_act_date_nds.append(act_number)
                    z += 1
            if re.search(r'ту №', str(item_one_info_in_act_f_Exel)):
                i += 1
                number_account = get_number_account(item_one_info_in_act_f_Exel)
                # print('Номер счета: ' + number_account)
                list_name_act_date_nds.append(number_account)
            if re.search(r'шт', str(item_one_info_in_act_f_Exel)) or re.search(r'ШТ', str(item_one_info_in_act_f_Exel)):
                if date_account and item_info_in_act_f_exel[3] != 'шт':
                    n += 1
                    list_name_act_date_nds.append(item_info_in_act_f_exel[0])
                    list_name_act_date_nds.append(item_info_in_act_f_exel[3])
                    all_list_name_act_date_nds.append(list_name_act_date_nds)
                    list_name_act_date_nds = []
    print('Количество актов: ' + str(i))
    print('Количество позиций поверки: ' + str(n))
    if n == i:
        print('Отлично в распознаном файле Exel ошибок нет')
    else:
        print('Ошибка просим обратить внимание на распознантый файл в название не правильно распозналось слово счету № '
              'или в таблицы не верно распознались шт')
    return all_list_name_act_date_nds


# Продумать закрытие и сохранение занесеных данных
# 1508 Продумать чтобы данные заносились в ячейки и фильтр их сортировал верно
def add_act_in_main_f(empty_line_in_table, all_list_name_act_date_nds):
    """
    Добавляем номер акта, сумму без НДС, дату акта, по совпадени с номером счета
    Args:
        empty_line_in_table:
        all_list_name_act_date_nds:

    Returns:

    """

    #[[дата, номер акта, номер счета, прибор, сумма без НДС]...]
    # Можно добавить проверку по названию прибора
    i = 0
    wb = xw.Book('Уралтест.xlsx')
    open_f_act = xlrd.open_workbook('Уралтест.xlsx')
    sheet_f_act = open_f_act.sheet_by_index(0)
    info_in_act = [sheet_f_act.row_values(rownum) for rownum in range(sheet_f_act.nrows)]
    for item_all_list_name_act_date_nds in all_list_name_act_date_nds:
        for item_info_in_act_f_exel in info_in_act:
            i += 1
            if not isinstance(item_info_in_act_f_exel[4], str):
                if int(item_info_in_act_f_exel[4]) == int(item_all_list_name_act_date_nds[2]):
                    xw.Range('K' + str(i)).value = item_all_list_name_act_date_nds[1]
                    xw.Range('L' + str(i)).value = item_all_list_name_act_date_nds[0]
                    xw.Range('M' + str(i)).value = item_all_list_name_act_date_nds[4]
    wb.save()
    wb.close()


def added_account():
    """
    Добавляем счет
    Returns:

    """
    # Выбранна cтраница Уралтест в основной таблице
    name_sheet = 0
    name_f_account_info_two_part = os.listdir(path="./transform_accounts")
    # Формируем путь нахождения файла с счетами
    for item in name_f_account_info_two_part:
        name_f_account_info = os.getcwd() + '\\transform_accounts\\' + item
        open_f_account = xlrd.open_workbook(name_f_account_info)
        sheet_f_account = open_f_account.sheet_by_index(0)
        # получаем список значений из всех записей
        info_in_account_f_exel = [sheet_f_account.row_values(rownum) for rownum in range(sheet_f_account.nrows)]
        # print(info_in_account_f_Exel)

        # Функция выводит список в списке с дангыми номер, дата, прибор, сумма с НДС
        all_list_name_account_date_nds = seach_need_info_in_account_f(info_in_account_f_exel)
        # print(all_list_name_account_date_nds)

        # Функция возращает список с даными с листа
        # open_main_task()

        # Возращает номер последней строки таблицы
        empty_line_in_table = get_empty_line_in_table(name_sheet)
        # print(empty_line_in_table)

        # Функция возращает список вида [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
        sort_all_list_name_account_date_nds = get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds)
        # print(sort_all_list_name_account_date_nds)

        # Функция добавления информации в таблицу
        add_info_in_main_f(empty_line_in_table, sort_all_list_name_account_date_nds)


def added_act():
    name_sheet = 0
    name_f_act_info_two_part = os.listdir(path="./transorm_act")
    # Формируем путь нахождения файла с счетами
    for item in name_f_act_info_two_part:
        name_f_act_info = os.getcwd() + '\\transorm_act\\' + item
        open_f_act = xlrd.open_workbook(name_f_act_info)
        sheet_f_act = open_f_act.sheet_by_index(0)
        # получаем список значений из всех записей
        info_in_act_f_exel = [sheet_f_act.row_values(rownum) for rownum in range(sheet_f_act.nrows)]
        # print(info_in_act_f_exel)
        # Функция выводит список в списке с данными:
        # [[дата, номер акта, номер счета, прибор, сумма без НДС] ...]
        all_list_name_act_date_nds = seach_need_info_in_act_f(info_in_act_f_exel)
        # print(all_list_name_act_date_nds)

        # Возращает номер последней строки таблицы
        empty_line_in_table = get_empty_line_in_table(name_sheet)

        # Функция добавления информации в таблицу
        add_act_in_main_f(empty_line_in_table, all_list_name_act_date_nds)


window_start()
