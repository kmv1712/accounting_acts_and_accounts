import os
import xlrd
import re
import xlwings as xw


def console_start():
    """
    Спрашивает о пользователя, что добавить акт или счет
    Консольное решение интерфейса
    Returns: вызывает фугкцию по добавление счета или акта

    """
    print('1 Добавить счет  \n2 Добавить акт по номеру счета')
    user = input('Введите число:')
    if int(user) == 1:
        added_account()
    else:
        added_act()


def get_number_account(item_one_info_in_account_f_exсel):
    """
    Ищем номер счета
    Args:
        item_one_info_in_account_f_exсel: строка найденная по подстроке для счета 'СЧЕТ' или для акта 'ту №'

    Returns: номер счета или акта

    """
    number_account = re.findall('(\d+)', item_one_info_in_account_f_exсel)
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
def get_date_account(item_one_info_in_account_f_excel):
    """
    Ищем дату в
    и преобразуем из формата 01.01.1992 в 01 января 1992
    Args:
        item_one_info_in_account_f_excel: строка найденая по подстроке

    Returns: Дату нужного вида

    """
    date_account = re.findall('(\d\d\s\w\w+\s\d\d\d\d)', item_one_info_in_account_f_excel)
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
def seach_need_info_in_account_f(info_in_account_f_excel):
    i = 0
    n = 0
    list_name_account_date_nds = []
    all_list_name_account_date_nds = []
    for item_info_in_account_f_excel in info_in_account_f_excel:
        for item_one_info_in_account_f_excel in item_info_in_account_f_excel:
            if re.search(r'СЧЕТ', item_one_info_in_account_f_excel):
                i += 1
                number_account = get_number_account(item_one_info_in_account_f_excel)
                # print('Номер: ' + number_account)
                list_name_account_date_nds.append(number_account)
                date_account = get_date_account(item_one_info_in_account_f_excel)
                list_name_account_date_nds.append(date_account)
                # print('Дата:' + date_account)
            if re.search(r'ШТ', item_one_info_in_account_f_excel):
                n += 1
                list_name_account_date_nds.append(item_info_in_account_f_excel[0])
                list_name_account_date_nds.append(item_info_in_account_f_excel[7])
                # print(item_info_in_account_f_excel)
                # print('Прибор: ' + item_info_in_account_f_excel[0])
                # print('Сумма с НДС: ' + item_info_in_account_f_excel[7])
    all_list_name_account_date_nds.append(list_name_account_date_nds)
    print('Количество счетов: ' + str(i))
    print('Количество позиций поверки: ' + str(n))
    if n == i:
        print('Отлично в распознаном файле Exel ошибок нет')
    else:
        print('Ошибка просим обратить внимание на распознантый файл в название не правильно распозналось слово СЧЕТ '
              'или в таблицы не верно распознались ШТ')
    return all_list_name_account_date_nds


# Вылавливаем ошибку для нахождения последней строки и вычетаем 1 для продолжения работы программы
# Находим пустую строку в тексте
def get_empty_line_in_table(name_sheet):
    try:
        rb = xlrd.open_workbook('Уралтест.xlsx')
        # выбираем активный лист
        sheet = rb.sheet_by_index(name_sheet)
        for i in range(0, 1000000000000):
            sheet.row_values(i)[0]
    except IndexError:
        i = i + 1
        return int(i)


"""Функция генерирует список all_list_name_account_date_nds
по след маске [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
"""


def get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds):
    n = 0
    j = 4
    sort_all_list_name_account_date_nds = []
    count_insert_list = len(all_list_name_account_date_nds[0]) / 4
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

            print(
                '\nПровертье цену (возможно ошибка при распознание)\nСЧЕТ: {0}\nРаспознаная цена: {1}\n'
                'введите цену если она не соответсвует цене в документе или нажмите Enter'
                    .format(sort_all_list_name_account_date_nds[i][0], sort_all_list_name_account_date_nds[i][3])
            )
            user_check = input()

            if user_check:
                xw.Range('G' + str(empty_line_in_table)).value = user_check
            else:
                xw.Range('G' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][3]
        empty_line_in_table = empty_line_in_table + 1
        wb.save()
        wb.close()


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

    # [[дата, номер акта, номер счета, прибор, сумма без НДС]...]
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


def get_all_num_accounts_in_table(empty_line_in_table):
    """
    Функция возвращает все номера счетов в таблице
    empty_line_in_table: последняя строка в табл.

    Returns(list): список номеров счетов

    """
    all_num_accounts_in_table = []
    wb = xw.Book('Уралтест.xlsx')
    open_f_act = xlrd.open_workbook('Уралтест.xlsx')
    sheet_f_act = open_f_act.sheet_by_index(0)
    info_in_act = [sheet_f_act.row_values(rownum) for rownum in range(sheet_f_act.nrows)]
    for i in range(3, empty_line_in_table):
        num_account_main_table = sheet_f_act.row_values(i)[4]
        if num_account_main_table:
            all_num_accounts_in_table.append(int(sheet_f_act.row_values(i)[4]))
    wb.close()
    return all_num_accounts_in_table


def check_acconts_main_table(sort_all_list_name_account_date_nds, all_num_accounts_in_table):
    """
    Функия проверяет есть ли счета с таким же номером в основной таблице
    Args:
        sort_all_list_name_account_date_nds:  Номера счетов которые будут добавлены. вместе с названием НДС и.т.д.
        all_num_accounts_in_table: список всех номеров счетов

    Returns: список с номерами счетов которых еще нет в таблице остальные не добавляет и сообщает пользователю

    """
    list_double_account = []
    list_unique = []
    key_in_list_sort_all_list = []
    # Надо проверить на работо способность
    # Формируем списки с уникальными и повторяющимися значениями
    # list_double_account спимок с характеристиками дублей
    # key_in_list_sort_all_list список только с номерами счетов дублей
    for attr_account in sort_all_list_name_account_date_nds:
        if attr_account[0] in all_num_accounts_in_table:
            list_double_account.append(attr_account)
            key_in_list_sort_all_list.append(attr_account[0])
        else:
            list_unique.append(attr_account[0])

    if len(list_double_account):
        pass
    else:
        print("Номера счетов которые вы добавляете совпадают с существующими в таблице\n"
              "Номера счетов| Сумма | Наименование:\n")
        for item in list_double_account:
            print('item{0:10}|{1:10}|{2:20}'.format(item[0], item[4], item[3]))
        print('Добавить счета в таблицу? Варианты:'
              '1 Добавить все счета\n'
              '2 Добавить только номера конкретных счетов\n'
              '3 Не добавлять дублирующиеся номера счетов, добавить только уникальные\n')
        user_input = input('Введите число')
        if user_input == 3:
            return list_unique
        elif user_input == 2:
            pass
        elif user_input == 1:
            return  list_unique + key_in_list_sort_all_list


    """
    Остановился на том что решил проверять весь список
    если добавляемй счет совпадает с имеющимся добавляем его номер в список и его позицию в общем списке
    чтобы удалить его или добавить в зависимости от того как решит пользователь
    """
    pass


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
        info_in_account_f_excel = [sheet_f_account.row_values(rownum) for rownum in range(sheet_f_account.nrows)]
        # print(info_in_account_f_Exel)

        # Функция выводит список в списке с данными номер, дата, прибор, сумма с НДС
        all_list_name_account_date_nds = seach_need_info_in_account_f(info_in_account_f_excel)

        # Возращает номер последней строки таблицы
        empty_line_in_table = get_empty_line_in_table(name_sheet)

        # Функция возращает список вида [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
        sort_all_list_name_account_date_nds = get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds)

        # Функция возвращает список номеров всех счтетов в таблице Уралтест
        all_num_accounts_in_table = get_all_num_accounts_in_table(empty_line_in_table)

        # Функция проверяет нет ли приборов с таким же счетом
        sort_all_list_name_account_date_nds = check_acconts_main_table(sort_all_list_name_account_date_nds,
                                                                       all_num_accounts_in_table)

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


console_start()
