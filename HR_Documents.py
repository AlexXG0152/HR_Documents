# -*- coding: utf-8 -*-
import os
import sys
from mailmerge import MailMerge
from time import strftime, localtime
from datetime import datetime, timedelta
import locale
import requests
from bs4 import BeautifulSoup
import tempfile
import pymorphy2
import csv


locale.setlocale(locale.LC_ALL, '')


def exception_handler(func):
    """ Exception handler for user input in functions """
    def inner_function(*args, **kwargs):
        while True:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                print(f'Проверьте введенные данные! \n{e}')
                continue
    return inner_function


def resource_path(relative_path):
    """
    Get the absolute path to the resource, works for dev and for PyInstaller
    This is need, because I'm create -onefile application and add inside blanks
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)


@exception_handler
def create_request_to_database(arg):
    """
    Creating request data for one employee if arg=1
    or all department (for find position later) if any another arg
    """
    if arg == 1:
        tnom = int(input('Введите табельный номер работника: '))
        payload = {'StrFIO': tnom}
    else:
        cex = int(input('Введите номер нужного подразделения: '))
        payload = {'StrCex': cex}

    return arg, payload


@exception_handler
def send_request_to_db(payload):
    """
    Send request do server, clear answer and return data.
    This server - its something like my API.
    I'm don't asking IT department about access directly to DB and just
    use this public server with HR data.
    Answer come to me in constant structure HTML page.
    Because of this i'm must clean server answer.
    """
    print('Нужно подожадать пару секунд...')
    try:
        r = requests.post('http://1.1.1.1/hrbase/hrbase.fwx', data=payload)
        r.encoding = 'windows-1251'

        soup = BeautifulSoup(r.text, 'lxml')

        titles = []
        for td in soup.find_all('td'):
            if td.get('title') is not None:
                titles.append(td.get('title'))

        if len(titles) == 0:
            print('Нет данных!')
            return 0, 0

        return titles, soup

    except requests.exceptions.HTTPError as errh:
        print('Http Error:', errh)
        main()
    except requests.exceptions.ConnectionError as errc:
        print('Error Connecting: ', errc)
        main()
    except requests.exceptions.Timeout as errt:
        print('Timeout Error: ', errt)
        main()
    except requests.exceptions.RequestException as err:
        print('OOps: Something Else', err)
        main()


@exception_handler
def parsing_server_answer(arg, payload):
    """
    Process answer data from request.
    I'm don't asking IT department about access directly to DB and just
    use this public server with HR data.
    Answer come to me in constant structure HTML page.
    Because of this i'm must clean server answer.
    """
    titles, soup = send_request_to_db(payload)  # titles - depatrment, soup - all_data

    # department name for case without workers data (vacancy)
    cex = str(titles).split('] ', 1)[1].split(', ', 1)[0].replace("'", "")

    # case 1 - user looking for one worker data.
    if arg == 1:
        # working with 3rd table on server answer HTML page
        for tr in soup.find_all('table')[3].find_all('tr')[1:]:
            """
            here func find and compare ID worker with answer data,
            because if user ask ID '132' server will show all users
            contain '132' in ID number (1322, 7132, etc...)
            """
            if int(tr.text[0:6]) == int(payload['StrFIO']):
                data = tr.find_all('td')
                cex = data[2].get('title').split('] ', 1)[1].split(', ', 1)[0].replace("'", "")  # depatrment
                professia = data[3].get('title').split(") ", 2)[1].replace("']", "")  # profession
                tab_n = data[0].text  # worker ID number
                fio = data[1].text[1:]  # worker's famaly and name

                # working with enpty request (fired workers)
                while True:
                    if len(fio) > 0 and '  ' in fio:
                        fio = fio.replace('  ', '')
                        print(f'\n{tab_n}\n{fio}\n{cex}\n{professia}\n')
                    else:
                        break
                return cex, professia, tab_n, fio

    # case 2 - user looking for department / profession data.
    else:
        profs = []
        for k, i in enumerate(titles, 1):
            if k % 2 == 0:
                i = str(i)
                professia = i.split(') ', 3)[1].replace("']", "")
                profs.append(professia)

        # print list of professions in this department
        for k, v in enumerate(sorted(set(profs)), 1):
            print(f'№ {k} === {v}')

        # choise from list of professions in this department
        select = int(input('Укажите необходимый номер профессии по порядку: '))
        professia = sorted(set(profs))[select - 1]

        return cex, professia


@exception_handler
def create_list_of_holidays(date_start, modified_end_date):
    """
    Here app creating list of holidays every time, when user choise vacation
    document. Holidays create for every year from start date to end date
    if vacation on border between 2 years
    """
    # strong holidays in every year
    hdays_strong = [
        '01.01.', '02.01.', '07.01.', '08.03.', '01.05.', '09.05.',
        '03.07.', '07.11.', '25.12.'
    ]
    # soft orthodox confession holiday in every year from 2020 to 2030
    hdays_soft = [
        '28.04.2020', '11.05.2021', '03.05.2022', '25.04.2023',
        '14.05.2024', '29.04.2025', '21.04.2026', '11.05.2027',
        '25.04.2028', '17.04.2029', '07.05.2030'
    ]
    # find years from dates
    year_start = date_start[-4:]
    year_end = modified_end_date[-4:]
    # create empty list for holydays
    prazd = hdays_soft
    # loop for add to list of holidays between start and end dates
    for i in hdays_strong:
        if year_end != year_start:
            # prazd.extend((i + year_start, i + year_end))
            prazd.append(i + year_start)
            prazd.append(i + year_end)
        else:
            prazd.append(i + year_start)

    return prazd


@exception_handler
def check_holidays(date_start, modified_end_date, holidays):
    """
    Here app check if holidays in dates of vacation or not.
    If Yes - add days to vacation, if Not - end date unchangeable
    """
    # first end date for check loop because end date move +1 for every weekend
    date_1 = datetime.strptime(date_start, '%d.%m.%Y')  # start date
    # second end date (after add holidays)
    date_2 = datetime.strptime(modified_end_date, '%d.%m.%Y')
    # third end date for finish date after adding holidays in vacations
    date_3 = datetime.strptime(modified_end_date, '%d.%m.%Y')

    # counter for days in vacation
    x = 0

    # loop for check dates in created holidays list
    for i in holidays:
        if date_1 <= datetime.strptime(i, '%d.%m.%Y') <= date_2:
            print(i)
            x += 1
            date_2 = date_2 + timedelta(days=1)
    print(x)

    # adding counter to first end date
    date_end = date_3 + timedelta(days=x)
    date_end = datetime.strftime(date_end, '%d.%m.%Y')

    return date_end


@exception_handler
def date_input(times):
    """
    Input date (or two if user know end date) without checking inputed data.
    I'm not checking data here because im dont need any checks.
    User can input anything what  he want and replace this in given blank.
    """
    date_start = str(input('Дата начала (ОБРАЗЕЦ - "01.02.2021"): '))
    print(date_start)
    if times == 1:
        date_end = 0
    else:
        print('Если на один день, просто нажмите Enter.')
        date_end = str(input('Дата окончания (ОБРАЗЕЦ - "17.02.2021"): '))
    print(date_end)

    return date_start, date_end


@exception_handler
def input_date_for_vacation(arg):
    """ Input dates and check if dates not in holiday list (for vacations) """
    while True:
        # input and check user input format data.
        date = datetime.strptime(
            input(f'Введите дату {arg} (ОБРАЗЕЦ - "01.02.2021"): '), '%d.%m.%Y'
            ).strftime('%d.%m.%Y')
        # check if day start not in holidays list
        holidays_list = create_list_of_holidays(date, date)
        if date not in holidays_list:
            return date


@exception_handler
def count_dates_for_vacation():
    """ Working with dates and counting end date of vacations """
    date_start = input_date_for_vacation('НАЧАЛА')
    print(f'Дата НАЧАЛА отпуска: {date_start}')

    message = '''
Дата окончания (нажмите цифру 1)
Количество дней (нажмите цифру 2)
Заполнят в отделе персонала (нажмите цифру 3)
Ваш выбор: '''

    date_or_days = str(input(message))
    while date_or_days not in ['1', '2', '3']:
        date_or_days = str(input(message))

    # first variant - user know end date of his vacation
    if date_or_days == '1':
        date_end = input_date_for_vacation('ОКОНЧАНИЯ')
        print(f'Дата ОКОНЧАНИЯ отпуска: {date_end}')
        days = '_____'
        return date_start, date_end, days

    # second variant - user know how much days he want
    elif date_or_days == '2':
        # change format of start date
        date_1 = datetime.strptime(date_start, '%d.%m.%Y')
        days = int(input('Количество дней (например 14): '))
        # add days inputed by user
        end_date_kd = date_1 + timedelta(days - 1)
        # change format of end date
        modified_end_date = end_date_kd.strftime('%d.%m.%Y')
        # information message for first data (before check holidays list)
        # print(f'INFO - modified_end_date = {modified_end_date}')
        # create list of holidays between two dates (start and modified)
        holidays = create_list_of_holidays(date_start, modified_end_date)
        # check end date if dates in holidays
        date_end = check_holidays(date_start, modified_end_date, holidays)
        print(f'Дата окончания отпуска: {date_end}')
        return date_start, date_end, days

    # third variant - user know just start date of his vacation. Another info fill HR on paper
    else:
        days = '_____'
        date_end = '______________'
        return date_start, date_end, days


@exception_handler
def format_date_for_template(date_start, date_end):
    """
    Formating dates with rules of Russian language for filling Word template
    Case for one date and for two date
    """
    if date_end != '':
        date_start = f'c {date_start}'
        date_end = f'по {date_end}'
    else:
        date_start = f'на {date_start}'
        date_end = ''
    return date_start, date_end


@exception_handler
def vacation_finance_help(date_start):
    """
    Finance help for vacation. Paid one time in year.
    Fill in right place in vacation Word blank.
    """
    year = date_start[-4:]  # because finance help paid only once in year on start date
    finance_help = str(input('Материальная помощь? Да/Нет '))
    if finance_help.lower() in ['Да', 'да', 'Д', 'д', 'Yes', 'Y', 'y', '1']:
        finance_help = f'''Прошу выплатить единовременную материальную помощь к отпуску на \
                оздоровление за {year} год. В {year} году материальную помощь не получал.'''
    else:
        finance_help = '_' * 204
    return finance_help


@exception_handler
def replace_wrong_cex_names(gde):
    """
    Because server sometimes (for long department names) not showing right names
    app replacing answer to right and full department names
    """
    if gde.startswith('Цех эксплуатации и обслуживания'):
        gde = 'Цех эксплуатации и обслуживания автомобилей'
    elif gde.startswith('Управление производством'):
        gde = 'Управление производством автомобилей'
    elif gde.startswith('Отдел текущей эксплуатации'):
        gde = 'Отдел текущей эксплуатации транспорта'
    elif gde.startswith('Управление автоматизации'):
        gde = 'Управление автоматизации и автомобильных технологий'
    elif gde.startswith('Цех контрольно-измерительных'):
        gde = 'Цех контрольно-измерительных приборов и автоматики'
    elif gde.startswith('Цех отопления'):
        gde = 'Цех отопления и вентиляции'
    elif gde.startswith('Цех перемотки'):
        gde = 'Цех перемотки, упаковки и отправки'
    return gde


@exception_handler
def site_in_department(gde_2, file):
    """ Read CSV files with data needed for fill templates. """
    list_from_csv = []
    # read file and find: 
    # 1. By user department number input
    # 2. Append all data in csv to list
    with open(f"{file}.csv", newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=';', quotechar='|')
        if gde_2 != "0":
            print('\nУчастки в подразделении: ')
            for row in spamreader:
                if gde_2 in row[1]:
                    list_from_csv.append(row[1:])
        else:
            for row in spamreader:
                list_from_csv.append(row)
                
    # create enumerating loop for choise number in list and easy user choise and add to Word template
    for k, v in enumerate((list_from_csv)[1:], 1):
        if file == 'employer':
            print(f"{k} = {', '.join(v[0:3])}")
        else:
            print(f"{k} = {', '.join(v[1:3])}")

    # return right part by number
    if file == 'type_mol':
        liability = []
        select = True
        print('\nВведите 17 для формирования договора\n')
        while select != 17:
            select = int(input('Укажите необходимый вид ответственности по порядку: '))
            if select == 12:
                liability.append('__' * 177)
                continue
            elif select == 17:
                continue
            elif select not in range(1, 13):
                print('Проверьте введенные данные!')
            liability.append(list_from_csv[select][1])
        return ('. '.join(map(str, liability)))  #unpack list for easier fill word template
    
    else:
        select = int(input('\nУкажите необходимый номер по порядку: '))
        part_list_from_csv = (list_from_csv)[select]
        return part_list_from_csv


@exception_handler
def initials(fio):
    """ Get short version of Name / Father Name for some type of document (#2) """
    fio_initials = fio.split()

    if len(fio_initials) >= 3:
        fio = f'{fio_initials[0]} {fio_initials[1][0]}.{fio_initials[2][0]}.'
    elif len(fio_initials) == 2:
        fio = f'{fio_initials[0]} {fio_initials[1][0]}.'
    else:
        fio = f'{fio_initials[0]}'

    return fio


@exception_handler
def sklonenie(cex, professia, fio, padez):
    """
    Changing cex, professia, fio with rules of Russian language.
    Last two letters taken from surname and checking entry into dictionary.
    If there is, it changes, if not, a special library works.
    Further,  name of  workshop and profession are divided into parts.
    Most often, only first word in title is declined, so there are two parts.
    Then everything that was declined is going back and again checking for name
    of structural divisions and professions.
    """
    morph = pymorphy2.MorphAnalyzer()

    surname = {'ич': 'ича', 'ко': 'ко', 'ий': 'ого', 'ов': 'ова',
               'ик': 'ика', 'ая': 'ую', 'ва': 'ву', 'ук': 'ука',
               'ин': 'ина', 'ев': 'ева', 'на': 'ну', 'юк': 'юка',
               'ль': 'ль', 'ок': 'ок', 'ак': 'ака', 'ец': 'еца',
               'ан': 'ана', 'ло': 'ло', 'ун': 'уна', 'ач': 'ача',
               'да': 'ду', 'ей': 'ея', 'ёв': 'ёва', 'та': 'ты',
               'ер': 'ера', 'ый': 'ого', 'ис': 'иса', 'ка': 'ки',
               'як': 'яка', 'ух': 'уха'}

    exception = {' цех': ' цеха', 'цех ': 'цеха ', 'отдел ': 'отдела ',
                 ' отдел': ' отделa', ' лаборатория': ' лаборатории',
                 'управление ': 'управления ', 'слесарящую': 'слесаря',
                 'инженер-электроники': 'инженера-электроникa',
                 'логистики': 'логистика', 'металлизатор': 'металлизатора',
                 'центральную': 'центральной',
                 'распределитель': 'распределителя', 'техники': 'техника',
                 'отдела логистика': 'отдела логистики'}

    # split cex and profession because in 90% cases changed only first word
    cex_1, cex_2 = cex.split()[0], cex.split()[1:]
    professia_1, professia_2 = professia.split()[0], professia.split()[1:]

    # split FIO for easier change in cases what save in dict surname
    f_split, io_split = fio.split()[0], fio.split()[1:]

    changed_data_list = []  # create list for save changed data
    f = 0  # for save unchanged famaly

    # check last 2 letter in famaly and add if case in dict
    for k, v in surname.items():
        if k in f_split[-2:]:
            f = f'{f_split[:-2]}{v}'
            changed_data_list.append(f)
    if f == 0:  # if not - save unchanged famaly
        changed_data_list.append(f_split)

    # combine for change all another info in the same time
    data_for_work = [i for i in io_split] + [professia_1] + [cex_1]

    # working and change data for work list
    # [I, O, professia (first word), cex (first word)]
    for i in data_for_work:
        part = morph.parse(i)[0]
        accs = part.inflect({padez})  # vinit = accs // rodit = gent
        changed_data_list.append(accs.word)

    # concatenate full changed data
    fio_skl = f'{changed_data_list[0]} {changed_data_list[1]} {changed_data_list[2]}'.title()
    professia_skl = f'{changed_data_list[3]} {" ".join(map(str, professia_2))}'
    cex_skl = f'{changed_data_list[4]} {" ".join(map(str, cex_2))}'

    # fix fails and exceptions in changed data [professia, cex]
    for k, v in exception.items():
        if k in cex_skl or k in professia_skl:
            cex_skl = cex_skl.replace(k, v)
            professia_skl = professia_skl.replace(k, v)

    # fix fail with letter 'a' in some cases
    professia_skl = professia_skl.replace('цехааа', 'цеха').replace('цехаа', 'цеха').replace('отделaа', '')

    return cex_skl, professia_skl, fio_skl


@exception_handler
def dismissal_reason():
    """ Data for fill reason in dismissal blank """
    dismissal_reason = {1: 'по соглашению сторон',
                        2: 'в связи с истечением срока трудового договора',
                        3: 'досрочно, по требованию работника, в связи с выходом на пенсию'}

    # enumerate print for easy choice
    for k, v in dismissal_reason.items():
        print(f'{k} --- {v}')

    reason = int(input('Введите номер причины увольнения: '))

    return dismissal_reason[reason]


@exception_handler
def data_for_liability_contract():  # MOLs
    """Data for filling out a liability agreement""" 
    type_mol = site_in_department("0", "type_mol")
    employer = site_in_department("0", "employer")

    return employer, type_mol


@exception_handler
def fill_and_show_ready_template(
        choise, fio, gde, kem, tab, date_start, date_end, days,
        modified_end_date, finance_help, gde_2, kem_2, tab_2, fio_2, razr):
    """ Filling choisen template and open ready Word doc """
    # template file and start year
    template_1 = resource_path(f'{choise}.docx')
    # fill document
    document_2 = MailMerge(template_1)
    document_2.merge(
        FIO=str(fio),
        gde=str(gde.lower()),
        kem=str(kem),
        tab=str(tab),
        date_start=str(date_start),
        date_end=str(date_end),
        days=str(days),
        modified_end_date=str(modified_end_date),
        finance_help=str(finance_help),
        razr=str(razr),
        FIO2=str(fio_2),
        gde_2=str(gde_2.lower()),
        kem_2=str(kem_2),
        tab_2=str(tab_2))

    # create filename
    dateform = strftime('%d.%m.%Y %H.%M.%S', localtime())
    filename = f'{choise} {fio} [{dateform}].docx'

    # use template directory for run file and show user result of fill
    with tempfile.TemporaryDirectory() as tmpdirname:
        document_2.write(tmpdirname + filename)
        os.startfile(tmpdirname + filename)


@exception_handler
def main():
    """ USER CHOISE TO WHAT TO DO AND WHAT TO FILL """
    try:
        choise = int(input("""
Выберите номер документа для формирования:
1. Записка об отпуске.
2. Заявление за свой счет.
3. Докладная о переводе.
4. Докладная об исполнении.
5. Заявление о переводе.
6. Заявление об увольнении.
7. Договор МОЛ (индивидуальный).
8. Выход.

НЕ ЗАБУДЬТЕ ИСПРАВИТЬ ПАДЕЖИ В СФОРМИРОВАННОМ ДОКУМЕНТЕ!

Ваш выбор -  """))

        if choise == 1:
            print('\nЗаписка об отпуске')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            date_start, date_end, days, = count_dates_for_vacation()
            finance_help = vacation_finance_help(date_start)
            modified_end_date = gde_2 = kem_2 = tab_2 = fio_2 = razr = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 2:
            print('\nЗаявление за свой счет')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            date_start, date_end = date_input(2)
            date_start, date_end = format_date_for_template(date_start, date_end)
            fio = initials(fio)  # just create short variant of FIO
            gde_2 = kem_2 = tab_2 = fio_2 = razr = ''
            days = modified_end_date = finance_help = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 3:
            print('\nДокладная о временном переводе')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            _, kem, fio = sklonenie(gde, kem, fio, 'accs')
            date_start, date_end = date_input(2)
            date_start, date_end = format_date_for_template(date_start, date_end)
            vak_ili_chel = int(input('Вакансия (нажмите цифру 1) или вместо другого работника (нажмите цифру 2): '))
            if vak_ili_chel == 1:  # vacancy
                arg, payload = create_request_to_database(arg=2)
                gde_2, kem_2 = parsing_server_answer(arg, payload)
                if gde_2 == 0:
                    return 0
                gde_2 = replace_wrong_cex_names(gde_2)
                tab_2, fio_2 = '', ''
            else:  # list of workers place in departent
                arg, payload = create_request_to_database(arg=1)
                gde_2, kem_2, tab_2, fio_2 = parsing_server_answer(arg, payload)
                tab_2 = f"таб.№ {tab_2})"
                if gde_2 == 0:
                    return 0
                gde_2 = replace_wrong_cex_names(gde_2)
                fio_2 = initials(fio_2)
            razr = str(input('Укажите, НА какой разряд БУДЕТ перевод рабочим! '))
            if razr not in ['0', '1', '2', '3', '4', '5', '6', '7', '8']:
                razr = ''
            else:
                razr = f'{razr} разряда'
            days = modified_end_date = finance_help = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help,
                         gde_2, kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 4:
            print('\nДокладная об исполнении обязанностей с доплатой')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            gde, kem, fio = sklonenie(gde, kem, fio, 'accs')
            date_start, date_end = date_input(2)
            date_start, date_end = format_date_for_template(date_start, date_end)
            arg, payload = create_request_to_database(arg=1)
            gde_2, kem_2, tab_2, fio_2 = parsing_server_answer(arg, payload)
            if gde_2 == 0:
                return 0
            gde_2 = replace_wrong_cex_names(gde_2)
            gde_2, kem_2, fio_2 = sklonenie(gde_2, kem_2, fio_2, 'accs')
            razr, days, modified_end_date, finance_help = '', '', '', ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 5:
            print('\nЗаявление о переводе')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            date_start, date_end = date_input(1)
            arg, payload = create_request_to_database(arg=2)
            gde_2, kem_2 = parsing_server_answer(arg, payload)
            if gde_2 == 0:
                return 0
            gde_2 = replace_wrong_cex_names(gde_2)
            razr = str(input('Укажите, НА какой разряд БУДЕТ перевод рабочим! '))
            if razr not in ['0', '1', '2', '3', '4', '5', '6', '7', '8']:
                razr = ''
            else:
                razr = f'{razr} разряда'
            tab_2 = site_in_department(gde_2, "sector")[1]  # site_in_department name
            days = modified_end_date = finance_help = date_end = fio_2 = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 6:
            print('\nЗаявление об увольнении')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            date_dis = str(input('Дату увольнения знаете (нажмите 1) или заполните вручную после печати (нажмите 2): '))
            if date_dis == '1':
                date_start, date_end = date_input(1)
            else:
                date_start = '____________________'
            tab_2 = dismissal_reason()  # dismissal_reason name
            days = modified_end_date = finance_help = ''
            date_end = fio_2 = gde_2 = kem_2 = razr = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 7:
            print('\nДоговор МОЛ (индивидуальный)')
            arg, payload = create_request_to_database(arg=1)
            gde, kem, tab, fio = parsing_server_answer(arg, payload)
            if gde == 0:
                return 0
            gde = replace_wrong_cex_names(gde)
            gde, _, _ = sklonenie(gde, kem, fio, 'accs')
            razr, finance_help, tab_2,  = site_in_department("0", "employer")  # mols (data for a liability contract)
            kem_2 = site_in_department("0", "type_mol")
            if razr == 'Иванов И.И.':
                finance_help = 'заместитель генерального директора ОАО «Рога и копыта»'
            mesto_raboty = 0
            try:
                # mesto_raboty = 0
                while mesto_raboty not in ['1', '2']:
                    mesto_raboty = str(input('По своей работе (нажмите 1) или вместо кого-то (нажмите 2): '))
                    if mesto_raboty == '1': 
                        date_start, date_end = gde, kem
                    elif mesto_raboty == '2':
                        arg, payload = create_request_to_database(arg=2)
                        date_start, date_end = parsing_server_answer(arg, payload) 
                        date_start = replace_wrong_cex_names(date_start)
                        date_start, date_end, _ = sklonenie(date_start, date_end, fio, 'accs')
                        date_end = f'и.о. {date_end}'
            except Exception:
                print('Проверьте введенные данные!')
            days = modified_end_date = gde_2 = fio_2 = ''
            arguments = [choise, fio, gde, kem, tab, date_start, date_end,
                         days, modified_end_date, finance_help, gde_2,
                         kem_2, tab_2, fio_2, razr]
            fill_and_show_ready_template(*arguments)

        elif choise == 8:
            print('\nВыход!')
            return None

        else:
            main()
        return choise

    except Exception as e:
        print(f'Проверьте введенные данные! \n{e}')
        main()


if __name__ == '__main__':
    end = ''
    while end.lower() != 'y':
        main()
        end = input('Enter "y"/"Y" to continue ')
