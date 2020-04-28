from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTableView


class PrintView(QTableView):
    def __init__(self):
        super(PrintView, self).__init__()

        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

    def print_(self, printer):
        self.resize(printer.width(), printer.height())
        self.render(printer)
