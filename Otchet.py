import openpyxl as opx
import datetime
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment
import itertools
from time import sleep

worker = {}
themes = []
date_start = datetime.datetime(2000, 1, 1)
date_end = datetime.datetime(2020, 12, 15)
tema_report = {}

thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))


def change_date(start, end):
    global date_start, date_end
    date_start = datetime.datetime(start[0], start[1], start[2])
    date_end = datetime.datetime(end[0], end[1], end[2])


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
                if row[4]:
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


def conductor(svod='Сводная таблица.xlsm',
              production='Общий перечень продукции ОВНТ.xlsx', report='Общий перечень отчётов.xlsm'):
    """
    Функция читает все файлы и прогоняет по ним обсчитыающие функции
    :param svod: Имя файла "Сводная таблица"
    :param production: Имя файла "Перечень продукции ОВНТ"
    :param report: Имя файла "Общий перечень отчётов"
    :return: None
    """
    process_bar = 5
    global worker, themes, tema_report
    worker = {}
    themes = []
    tema_report = {}
    yield process_bar, 'Сводная таблица'
    svod_wb = opx.load_workbook(filename=svod)
    svod_ws = svod_wb.active
    svod_tabl_count(read_excel(svod_ws))
    process_bar += 15

    production_wb = opx.load_workbook(filename=production)
    prod_names = production_wb.sheetnames
    good_names = {'Катя': 'Шаповал Е.С.', 'Женя': 'Гусева Е.Н.', 'Дима': 'Пихуров Д.В.', 'Вова': 'Васильев В.А.',
                  'Иван': 'Зуев И.А.', 'Артур': 'Калимуллин А.В.'}
    for name in prod_names:
        process_bar += 5
        yield process_bar, f'Список образцов: {name}'
        if name not in worker:
            if name not in good_names:
                continue
        production_count(read_excel(production_wb[name], col=3), good_names[name])

    yield process_bar, f'Список отчётов'
    report_wb = opx.load_workbook(filename=report)
    report_ws = report_wb.active
    report_count(read_excel(report_ws, col=7))
    yield 100, 'Готово!'


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
        if word_prefix == current_prefix and int(current_numb) == int(word_numb) - 1:
            current_numb = word_numb
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
            first_numb = word_numb
            current_numb = first_numb
            if word == massiv[-1] and current_prefix:
                if current_numb != first_numb:
                    fine.append(f'{current_prefix}{first_numb}-{current_numb}')
                else:
                    fine.append(f'{current_prefix}{first_numb}')
    return fine


def worker_into_thems():
    global tema_report
    for tema in themes:  # Проходим по всем темам
        for name in worker:  # Проходим по всем сотрудникам
            if tema in worker[name]:  # Смотрим есть ли тема у сотрудника
                if tema not in tema_report:  # Проверяем есть ли тема в отчётном словаре
                    tema_report[tema] = {}
                if name not in tema_report[tema]:  # Проверяем есть ли сотрудник в этой теме
                    tema_report[tema][name] = {'образцы': [], 'плёнки': [], 'отчёты': []}
                for i in worker[name][tema]['образцы']:
                    tema_report[tema][name]['образцы'].append(i)
                for i in worker[name][tema]['плёнки']:
                    tema_report[tema][name]['плёнки'].append(i)
                for i in worker[name][tema]['отчёты']:
                    tema_report[tema][name]['отчёты'].append(i)


def short_show_report(massiv):
    massiv_to_return = []
    for otchet in massiv:
        word = ''
        for b in otchet:
            if b.isdigit():
                word += b
            else:
                break
        massiv_to_return.append(word)

    return massiv_to_return


def make_excel():
    wb = opx.Workbook()
    ws_title = ['Выпущенные опытно-промышленные партии',
                'Изготовленные лабораторные образцы и компоненты',
                'Написанные отчеты', 'Выпущенные рецептуры и ТК', 'Проведенные промышленные нанесения']
    ws = wb.active
    ws_count_row = []
    for tema in tema_report:
        ws_title.insert(0, tema)
        ws.append(ws_title)
        ws_count_row.append(True)
        for name in tema_report[tema]:
            sum_obraz = len(tema_report[tema][name]['образцы']) + len(tema_report[tema][name]['плёнки'])
            sum_otchet = len(tema_report[tema][name]['отчёты'])
            current_list_to_append = [name]
            cell = ''
            for i in (short_show(tema_report[tema][name]['образцы']) +
                      short_show(tema_report[tema][name]['плёнки'])):
                cell += f'{i}, '
            if cell:
                cell = cell[:-2]
                cell += f'\n\nИтого: {sum_obraz}'
            current_list_to_append.append('')
            current_list_to_append.append(cell)
            cell = ''
            for i in short_show_report(tema_report[tema][name]['отчёты']):
                cell += f'{i}, '
            if cell:
                cell = cell[:-2]
                cell += f'\n\nИтого: {sum_otchet}'
            current_list_to_append.append(cell)
            for i in range(2):
                current_list_to_append.append(' ')
            ws.append(current_list_to_append)
            ws_count_row.append(True)
        ws_title.pop(0)
        for _ in range(2):
            ws.append([' '])
            ws_count_row.append(False)
        for index, row, row_numb in zip(ws_count_row, ws, itertools.count(1, 1)):
            if index:
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(wrapText=True)
            else:
                ws.merge_cells(f'A{row_numb}:F{row_numb}')

    ws.page_setup.paperSize = '9'

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 17
    ws.column_dimensions['F'].width = 18
    wb.save('Otchet.xlsx')


def make_excel_noname():
    ws_title = ['Темы', 'Выпущенные опытно-промышленные партии',
                'Изготовленные лабораторные образцы и компоненты',
                'Написанные отчеты', 'Выпущенные рецептуры и ТК', 'Проведенные промышленные нанесения']

    wb = opx.Workbook()
    ws = wb.active
    ws_count_row = []
    ws.append(ws_title)
    ws_count_row.append(True)
    for tema in tema_report:
        current_list_to_append = [tema]
        sum_obraz = 0
        sum_otchet = 0
        cell = ''
        cell2 = ''
        for name in tema_report[tema]:
            sum_obraz += len(tema_report[tema][name]['образцы']) + len(tema_report[tema][name]['плёнки'])
            sum_otchet += len(tema_report[tema][name]['отчёты'])
            for i in (short_show(tema_report[tema][name]['образцы']) +
                      short_show(tema_report[tema][name]['плёнки'])):
                cell += f'{i}, '
            for i in short_show_report(tema_report[tema][name]['отчёты']):
                cell2 += f'{i}, '

        if cell:
            cell = cell[:-2]
            cell += f'\n\nИтого: {sum_obraz}'
        if cell2:
            cell2 = cell2[:-2]
            cell2 += f'\n\nИтого: {sum_otchet}'

        current_list_to_append.append('')
        current_list_to_append.append(cell)
        current_list_to_append.append(cell2)
        for i in range(2):
            current_list_to_append.append('')
        ws.append(current_list_to_append)
        ws_count_row.append(True)


        # for _ in range(2):
        #     ws.append([' '])
        #     ws_count_row.append(False)
        for index, row, row_numb in zip(ws_count_row, ws, itertools.count(1, 1)):
            if index:
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(wrapText=True)
            else:
                ws.merge_cells(f'A{row_numb}:F{row_numb}')

    ws.page_setup.paperSize = '9'

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 17
    ws.column_dimensions['F'].width = 18
    wb.save('Otchet_noname.xlsx')




def import_data():
    conductor('Сводная таблица.xlsm', 'Общий перечень продукции ОВНТ2.xlsx', 'Общий перечень отчётов.xlsm')
    return


if __name__ == '__main__':

    conductor('Сводная таблица.xlsm', 'Общий перечень продукции ОВНТ.xlsx', 'Общий перечень отчётов.xlsm')

    # for name in worker:
    #     print(name)
    #     for tema in worker[name]:
    #         print(tema, ' : ', worker[name][tema]['отчёты'])
    #     print('----------------------------------------------------------------------------------------------')
    # print(themes)

    worker_into_thems()
    make_excel()
