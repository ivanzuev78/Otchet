from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem
import PyQt5
import sys
import Otchet


Form, _ = uic.loadUiType("main_window.ui")

class Ui(QtWidgets.QMainWindow, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        self.update_bot.clicked.connect(self.update_it)
        self.make_xl_bot.clicked.connect(self.make_xl)

    def update_it(self):
        Otchet.conductor()
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

    def make_xl(self):
        if Otchet.themes:
            print('1')
            Otchet.make_excel()
            print('2')
        else:
            print('else')
            self.update_it()
            Otchet.make_excel()
            print('else ok')

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
