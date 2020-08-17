from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import QDate
import sys
import Otchet
from time import localtime

Form, _ = uic.loadUiType("main_window.ui")


class Ui(QtWidgets.QMainWindow, Form):
    kvartal_dict = {0: {'start': [1, 1], 'end': [3, 31]},
                    1: {'start': [1, 4], 'end': [6, 30]},
                    2: {'start': [1, 7], 'end': [9, 30]},
                    3: {'start': [1, 10], 'end': [12, 31]}}

    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        self.update_bot.clicked.connect(self.update_it)
        self.make_xl_bot.clicked.connect(self.make_xl)
        self.make_xl_bot_2.clicked.connect(self.make_xl_noname)
        self.try_bot.clicked.connect(self.try_but)

        if localtime()[1] > 1:
            self.year.setDate(QDate(localtime()[0], 1, 1))
        else:
            self.year.setDate(QDate(localtime()[0] - 1, 1, 1))

    def try_but(self):
        pass

    def update_it(self):
        self.update_date()
        for PB in Otchet.conductor():
            self.progressBar.setValue(PB[0])
            self.current_action.setText(PB[1])
        Otchet.worker_into_thems()
        themes = Otchet.tema_report

        self.show_data.setRowCount(len(themes))
        for row, tema in enumerate(themes):
            count = {'отчёты': 0, 'плёнки': 0, 'образцы': 0}
            for name in themes[tema]:
                count['отчёты'] += len(themes[tema][name]['отчёты'])
                count['плёнки'] += len(themes[tema][name]['плёнки'])
                count['образцы'] += len(themes[tema][name]['образцы'])
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
        self.kvartal.currentIndex()
        Otchet.change_date([year] + self.kvartal_dict[self.kvartal.currentIndex()]['start'],
                           [year] + self.kvartal_dict[self.kvartal.currentIndex()]['end'])

    def make_xl(self):
        if Otchet.themes:
            Otchet.make_excel()
            self.current_action.setText('Отчёт сформирован!')
        else:
            self.update_it()
            Otchet.make_excel()
            self.current_action.setText('Отчёт сформирован!')

    def make_xl_noname(self):
        if Otchet.themes:
            Otchet.make_excel_noname()
            self.current_action.setText('Отчёт сформирован!')
        else:
            self.update_it()
            Otchet.make_excel_noname()
            self.current_action.setText('Отчёт сформирован!')

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
