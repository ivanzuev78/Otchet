from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import QDate
import sys
import Otchet
from time import localtime, struct_time
import os
import pickle

class Ui(QtWidgets.QMainWindow, uic.loadUiType("main_window.ui")[0]):
    kvartal_dict = {0: {'start': [1, 1], 'end': [3, 31]},
                    1: {'start': [4, 1], 'end': [6, 30]},
                    2: {'start': [7, 1], 'end': [9, 30]},
                    3: {'start': [10, 1], 'end': [12, 31]}}

    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        self.update_bot.clicked.connect(self.update_it)
        self.make_xl_bot.clicked.connect(self.make_xl)
        self.make_xl_bot_2.clicked.connect(self.make_xl_noname)
        self.settings.clicked.connect(self.settings_window)
        self.read_data_but.clicked.connect(self.read_data)
        self.try_bot.clicked.connect(self.try_but)
        self.path_program = os.getcwd()
        self.unknown_themes = []
        self.theme_list = {}
        self.known_themes = []
        self.load_settings()
        self.set_kvartal()


    def settings_window(self):
        self.window = SecondWindow(self)
        self.setEnabled(False)
        self.window.show()

    def try_but(self):
        for name in Otchet.worker:
            print(name)
            for i in Otchet.worker[name]:
                print(i, '   ', Otchet.worker[name][i])

            print('------------------------------')
        print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')

        for name in Otchet.replaced_worker:
            print(name)
            for i in Otchet.replaced_worker[name]:
                print(i, '   ', Otchet.replaced_worker[name][i])
        print('------------------------------')
        print(self.theme_list)
        pass

    def read_data(self):
        for i in Otchet.reader():
            self.progressBar.setValue(i)
        self.current_action.setText('Считано!')

    def update_it(self):
        self.update_date()
        Otchet.conductor()
        Otchet.replace_names_thems(self.theme_list, self.known_themes)
        Otchet.worker_into_thems_2()
        tema_report = Otchet.tema_report
        self.unknown_themes = Otchet.unknown_thems

        self.show_data.setRowCount(len(tema_report))
        for row, tema in enumerate(tema_report):
            count = {'отчёты': 0, 'плёнки': 0, 'образцы': 0}
            for name in tema_report[tema]:
                count['отчёты'] += len(tema_report[tema][name]['отчёты'])
                count['плёнки'] += len(tema_report[tema][name]['плёнки'])
                count['образцы'] += len(tema_report[tema][name]['образцы'])
            self.show_data.setItem(row, 0, QTableWidgetItem(f"{tema}"))
            self.show_data.setItem(row, 1, QTableWidgetItem(f"{count['образцы']}"))
            self.show_data.setItem(row, 2, QTableWidgetItem(f"{count['плёнки']}"))
            self.show_data.setItem(row, 3, QTableWidgetItem(f"{count['отчёты']}"))
            self.progressBar.setValue(100)

    def update_date(self):
        year = ''
        for i in self.year.dateTime().toString()[len(self.year.dateTime().toString()) - 4::]:
            year += i
        year = int(year)
        Otchet.change_date([year] + self.kvartal_dict[self.kvartal.currentIndex()]['start'],
                           [year] + self.kvartal_dict[self.kvartal.currentIndex()]['end'])

    def make_xl(self):
        if not Otchet.themes:
            self.update_it()
        path = Otchet.make_excel()

        # try:
        self.current_action.setText('Отчёт сформирован!')
        if self.open_check.isChecked():
            os.startfile(path)
        # except:
        #     self.current_action.setText('Отчёт ой!')



    def make_xl_noname(self):
        if Otchet.themes:
            self.update_it()
        path = Otchet.make_excel_noname()
        self.current_action.setText('Отчёт сформирован!')
        if self.open_check_2.isChecked():
            os.startfile(path)

    def load_settings(self):
        if os.path.exists('settings.otchet'):
            with open('settings.otchet', 'rb') as f:
                self.theme_list = pickle.load(f)
            for global_tema in self.theme_list:
                self.known_themes += self.theme_list[global_tema]

    def set_kvartal(self):
        if localtime()[1] > 1:
            self.year.setDate(QDate(localtime()[0], 1, 1))
        else:
            self.year.setDate(QDate(localtime()[0] - 1, 1, 1))

        data1 = struct_time((localtime()[0], 2, 1, 1, 1, 1, 1, 1, 1))
        data2 = struct_time((localtime()[0], 5, 1, 1, 1, 1, 1, 1, 1))
        data3 = struct_time((localtime()[0], 8, 1, 1, 1, 1, 1, 1, 1))
        data4 = struct_time((localtime()[0], 11, 1, 1, 1, 1, 1, 1, 1))

        if data1 <= localtime() < data2:
            self.kvartal.setCurrentIndex(0)
        elif data2 <= localtime() < data3:
            self.kvartal.setCurrentIndex(1)
        elif data3 <= localtime() < data4:
            self.kvartal.setCurrentIndex(2)
        else:
            self.kvartal.setCurrentIndex(3)


class SecondWindow(QtWidgets.QMainWindow, uic.loadUiType("settings.ui")[0]):
    def __init__(self, main_wind):
        self.main_window = main_wind
        super(SecondWindow, self).__init__()
        self.setupUi(self)
        self.close_but.clicked.connect(self.close)
        self.tema_navigator()
        self.tema_add_But.clicked.connect(self.add_global_tema)
        self.official_tems.itemSelectionChanged.connect(self.update_not_official_tems)
        self.move_tema_but.clicked.connect(self.move_tema)
        # self.save_but.clicked.connect(self.try_it)
        self.remove_tema_but.clicked.connect(self.remove_tema)
        self.del_tema_but.clicked.connect(self.remove_global_tema)

    def try_it(self):
        print(self.main_window.theme_list)

    def closeEvent(self, event):
        self.main_window.setEnabled(True)
        self.save_file()
        event.accept()

    def tema_navigator(self):
        for tema in self.main_window.theme_list:
            for tema_to_find in self.main_window.unknown_themes.copy():
                if str(tema_to_find) in self.main_window.theme_list[tema]:
                    self.main_window.unknown_themes.pop(self.main_window.unknown_themes.index(tema_to_find))
            self.official_tems.addItem(tema)

        # for tema_to_find in self.main_window.unknown_themes.copy():
        #     for tema in self.main_window.theme_list:
        #         if str(tema_to_find) in self.main_window.theme_list[tema]:
        #             self.main_window.unknown_themes.pop(self.main_window.unknown_themes.index(tema_to_find))


        while self.unkown_tema.takeItem(0):
            pass
        for tema in self.main_window.unknown_themes:
            self.unkown_tema.addItem(str(tema))

        # for tema in self.main_window.theme_list:
        #     self.official_tems.addItem(tema)

    def add_global_tema(self):

        if str(self.tema_add_line.text()) not in self.main_window.theme_list and self.tema_add_line.text():
            self.main_window.theme_list[str(self.tema_add_line.text())] = []
            self.official_tems.addItem(self.tema_add_line.text())
        self.tema_add_line.setText('')

    def remove_global_tema(self):
        tema_to_del = self.official_tems.takeItem(self.official_tems.currentRow()).text()
        for tema in self.main_window.theme_list[tema_to_del]:
            self.unkown_tema.addItem(str(tema))
        self.main_window.unknown_themes += self.main_window.theme_list[tema_to_del]
        del self.main_window.theme_list[tema_to_del]

    def move_tema(self):
        try:
            if self.official_tems.selectedItems()[0].text():
                tema_fly = self.unkown_tema.takeItem(self.unkown_tema.currentRow()).text()
                for ind, tem in enumerate(self.main_window.unknown_themes):
                    if str(tem) == tema_fly:
                        self.main_window.unknown_themes.pop(ind)
                        break
                self.main_window.theme_list[self.official_tems.selectedItems()[0].text()].append(tema_fly)
                self.update_not_official_tems()
        except:
            pass

    def remove_tema(self):
        try:
            if self.not_official_tems.selectedItems()[0].text():
                tema_fly = self.not_official_tems.takeItem(self.not_official_tems.currentRow()).text()
                self.main_window.theme_list[self.official_tems.selectedItems()[0].text()].pop(
                    self.main_window.theme_list[self.official_tems.selectedItems()[0].text()].index(tema_fly))

                self.main_window.unknown_themes.append(tema_fly)
                self.unkown_tema.addItem(tema_fly)
        except:
            pass

    def update_not_official_tems(self):
        while self.not_official_tems.takeItem(0):
            pass
        if self.official_tems.selectedItems():
            for tema in self.main_window.theme_list[self.official_tems.selectedItems()[0].text()]:
                self.not_official_tems.addItem(tema)

    def save_file(self):
        with open('settings.otchet', 'wb') as f:
            pickle.dump(self.main_window.theme_list, f)

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
