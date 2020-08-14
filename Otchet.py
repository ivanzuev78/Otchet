import openpyxl as opx
import datetime


worker = {}
themes = []
date_start = datetime.datetime(2020, 1, 1)
date_end = datetime.datetime(2020, 12, 15)


def path_to_excel(name: str) -> None:
    """
    Находит путь к файлу эксель
    :return:
    """

    pass


def read_excel(ws, col: int = 5) -> list:
    """
    Функция для чтения excel файла
    :param ws: Лист страницы Excel, с которого надо получить данные
    :param col: количество считаных столбцов
    :return: Список массивов с элементами ячеек
    """
    sv_tabl = []  # Список массивов, который вернем

    for row in ws:  # Проходим по всем строкам на листе
        current_row = []  # Память для текущего ряда

        # Проходим по всем ячейкам в строке. Нумеруем, что бы не считывать все колонки
        for index, cell in enumerate(row):
            if index == col:
                break
            current_row.append(cell.value)  # Добавляем значение ячейки в массив
        sv_tabl.append(current_row)  # Добавляем массив в список
    return sv_tabl


def input_name(name: str) -> str:
    """
    :param name: str (ФамилияИ.О. в любом формате)
    :return: str (Фамилия И.О. с пробелом после фамилии)
    """

    integ = ''  # Возвращаемая строка
    probel_check = False  # Флаг для постановки пробела после фамилии
    prev = ''  # Предыдущий символ при проходе ФИО
    for index, i in enumerate(name):  # Проходим по всем элементам
        if i == '.' and not probel_check:  # Если встречаем точку
            integ += ' '  # Вставляем пробел перед добавлением предыдущего символа
            probel_check = True   # Меняем флаг
        if i != ' ':  # Если символ не пробел
            integ += prev  # Добавляем предыдущий
            prev = i  # Запоминаем текущий
    integ += prev  # Добавляем последний в конце
    return integ  # Возвращаем ФИО


def svod_tabl_count(tabl: list):
    """
    Элементы массива (номер колонки - 1):
    0) Маркировка плёнки
    1) Тема
    2) Нанесение
    3) Дата
    4) Ответственный
    Добавляет в общий отчет данные из сводной таблицы
    :param tabl: Полученная таблица
    :return: А хз, еще не решил
    """
    global worker, themes

    for row in tabl:  # Берем строку
        if type(row[3]) is type(date_start):  # Смотрим, является ли клетка с датой датой, а не заглавной строкой
            if date_start <= row[3] <= date_end:  # Смотрим, находится ли дата в нужном диапазоне
                row[4] = input_name(row[4])  # Форматируем ФИО по шаблону
                if row[4] not in worker:  # Если нет работника в словаре, то создаем
                    worker[row[4]] = {}
                if row[1] not in worker[row[4]]:  # Если нет темы у работника, то создаем
                    worker[row[4]][row[1]] = {'образцы': [], 'плёнки': [], 'отчёты': []}
                worker[row[4]][row[1]]['плёнки'].append(row[0])  # Добавляем номер пленки работнику в тему
                if row[1] not in themes and row[1]:
                    themes.append(row[1])

    return None


def report_count(tabl):
    """
    0) Номер отчёта
    1) Название отчёта
    2) Ответственный 1
    3) Ответственный 2
    4) Ответственный 3
    5) Дата
    6) Название темы
    :param tabl:
    :return:
    """
    for row in tabl:  # Берем строку
        if type(row[5]) is type(date_start):  # Смотрим, является ли клетка с датой датой, а не заглавной строкой
            if date_start <= row[5] <= date_end:  # Смотрим, находится ли дата в нужном диапазоне
                row[2] = input_name(row[2])  # Форматируем ФИО по шаблону
                if row[3]:
                    row[3] = input_name(row[3])
                if row[4]:
                    row[4] = input_name(row[4])

                if row[2] not in worker:  # Если нет работника в словаре, то создаем
                    worker[row[2]] = {}
                if row[6] not in worker[row[2]]:  # Если нет темы у работника, то создаем
                    worker[row[2]][row[6]] = {'образцы': [], 'плёнки': [], 'отчёты': []}
                if row[0]:
                    worker[row[2]][row[6]]['отчёты'].append(row[0])  # Добавляем отчёт
                if row[6] not in themes and row[6]:   # Добавляем тему в список тем, если её там еще нет
                    themes.append(row[6])


def production_count(tabl: list, name: str):
    """
    Элементы массива (номер колонки - 1):
    0) Маркировка состава
    1) Дата
    2) Тема
    :param tabl: Полученная таблица
    :param name: Имя сотрудника
    :return: None
    """
    global worker

    for row in tabl:  # Берем строку
        if type(row[1]) is type(date_start):  # Смотрим, является ли клетка с датой датой, а не заглавной строкой
            if date_start <= row[1] <= date_end:
                if name not in worker:  # Если нет работника в словаре, то создаем
                    worker[name] = {}
                if row[2]:  # Если клетка с темой не пустая
                    if row[2] not in worker[name]:  # Если нет темы у работника, то создаем
                        worker[name][row[2]] = {'образцы': [], 'плёнки': [], 'отчёты': []}
                    worker[name][row[2]]['образцы'].append(row[0])  # Добавляем номер пленки работнику в тему
                if row[2] not in themes and row[2]:
                    themes.append(row[2])
    return None


def conductor(svod, production, report):
    """
    Функция читает все файлы и прогоняет по ним обсчитыающие функции
    :param svod: Имя файла "Сводная таблица"
    :param production: Имя файла "Перечень продукции ОВНТ"
    :param report: Имя файла "Общий перечень отчётов"
    :return: None
    """

    svod_wb = opx.load_workbook(filename=svod)
    svod_ws = svod_wb.active
    svod_tabl_count(read_excel(svod_ws))

    production_wb = opx.load_workbook(filename=production)
    prod_names = production_wb.sheetnames
    good_names = {'Катя': 'Шаповал Е.С.', 'Женя': 'Гусева Е.Н.', 'Дима': 'Пихуров Д.В.', 'Вова': 'Васильев В.А.',
                  'Иван': 'Зуев И.А.', 'Артур': 'Калимуллин А.В.'}
    for name in prod_names:
        if name not in worker:
            if name not in good_names:
                continue
        production_count(read_excel(production_wb[name], col=3), good_names[name])

    report_wb = opx.load_workbook(filename=report)
    report_ws = report_wb.active
    report_count(read_excel(report_ws, col=7))


def short_show(massiv):
    """
    Функция позволяет сворачивать штучные образцы в дефисные группы
    :param massiv: Массив с образцами или плёнками поштучно
    :return: Массив с образцами или плёнками, свёрнуто дефисными группами
    """
    fine = []
    current_prefix = ''
    current_numb = -1
    first_numb = -1
    for word in massiv:
        word_prefix = ''
        word_numb = ''
        for b in word:
            if b.isdigit():
                word_numb += b
            else:
                word_prefix += b
        if word_prefix == current_prefix and current_numb == int(word_numb) - 1:
            current_numb = int(word_numb)
            if word == massiv[-1]:
                if current_numb != first_numb:
                    fine.append(f'{current_prefix}{first_numb}-{current_numb}')
                else:
                    fine.append(f'{current_prefix}{first_numb}')
        else:
            if current_prefix:
                if current_numb != first_numb:
                    fine.append(f'{current_prefix}{first_numb}-{current_numb}')
                else:
                    fine.append(f'{current_prefix}{first_numb}')
            current_prefix = word_prefix
            first_numb = int(word_numb)
            current_numb = first_numb
            if word == massiv[-1] and current_prefix:
                if current_numb != first_numb:
                    fine.append(f'{current_prefix}{first_numb}-{current_numb}')
                else:
                    fine.append(f'{current_prefix}{first_numb}')
    return fine


if __name__ == '__main__':
    name = 'Зуев И.А.'
    # svod_tabl_count(read_excel('Сводная таблица.xlsm'))
    # production_count(read_excel('Общий перечень продукции ОВНТ2.xlsx', col=3, sheet=5), input_name(name))
    # report_count(read_excel('Общий перечень отчётов.xlsm', col=7))

    # for name in opx.load_workbook('Общий перечень продукции ОВНТ2.xlsx').

    conductor('Сводная таблица.xlsm', 'Общий перечень продукции ОВНТ2.xlsx', 'Общий перечень отчётов.xlsm')
    # for name in worker:
    #     print(name)
    #     for i in worker[name]:
    #         print(i, ": ", worker[input_name(name)][i])
    #     print('----------------------------------------------------------------------------------------------')
    # print(themes)
    print(worker['Пихуров Д.В.']['ППУ']['образцы'])
    print(short_show(worker['Пихуров Д.В.']['ППУ']['образцы']))
