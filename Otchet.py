import openpyxl as opx

worker = {}

def path_to_excel(name: str) -> None:
    """
    Находит путь к файлу эксель
    :return:
    """

    pass


def read_excel(name: str, col: int = 5) -> list:
    """
    Функция для чтения excel файла
    :param name: имя файла
    :param col: количество считаных столбцов
    :return: Список массивов с элементами ячеек
    """
    sv_tabl = []  # Список массивов, который вернем
    wb = opx.load_workbook(filename=name)  # Открываем файл
    ws = wb.active  # Запоминаем активный лист. По умолчанию первый.

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
    global worker

    for row in tabl:  # Берем строку
        row[4] = input_name(row[4])  # Форматируем ФИО по шаблону
        if row[4] not in worker:  # Если нет работника в словаре, то создаем
            worker[row[4]] = {}
        if row[1] not in worker[row[4]]:  # Если нет темы у работника, то создаем
            worker[row[4]][row[1]] = {'образцы': [], 'плёнки': [], 'отчёты': []}
        worker[row[4]][row[1]]['плёнки'].append(row[0])  # Добавляем номер пленки работнику в тему
    return None

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
        if row[name] not in worker:  # Если нет работника в словаре, то создаем
            worker[name] = {}
        if row[2] not in worker[name]:  # Если нет темы у работника, то создаем
            worker[name][row[2]] = {'образцы': [], 'плёнки': [], 'отчёты': []}
        worker[name][row[2]]['образцы'].append(row[0])  # Добавляем номер пленки работнику в тему
    return None



if __name__ == '__main__':
    svod_tabl_count(read_excel('Сводная таблица.xlsm'))
    print(worker[input_name(" Шаповал Е.С.")])


