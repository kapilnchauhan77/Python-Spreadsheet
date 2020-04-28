from PyQt5.QtCore import QDate, QPoint, Qt, QSizeF
from PyQt5.QtGui import QColor, QIcon, QKeySequence, QPainter, QPixmap, QPalette, QImage
from PyQt5.QtWidgets import (QAction, QActionGroup, QApplication, QColorDialog,
                             QComboBox, QDialog, QFontDialog, QGroupBox, QFormLayout,
                             QHBoxLayout, QLabel, QPushButton, QTableWidget, QInputDialog,
                             QLineEdit, QMainWindow, QMessageBox, QToolBar, QDialogButtonBox,
                             QTableWidgetItem, QVBoxLayout, QWidget, QFileDialog, QSizePolicy, QCheckBox)
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog, QPrintDialog
import sqlite3
from printview import PrintView
import webcolors
from spreadsheetdelegate import SpreadSheetDelegate
from spreadsheetitem import SpreadSheetItem
from util import decode_pos, encode_pos
from PyQt5 import QtGui
import csv
import xlrd
import numpy as np
from functools import partial
import smtplib
from email.message import EmailMessage
import shutil
import os
import imghdr
import requests
from PyPDF2 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
# import keyboard


# URL = 'https://www.way2sms.com/api/v1/sendCampaign'


# def sendPostRequest(reqUrl, apiKey, secretKey, useType, phoneNo, senderId, textMessage, number_to_send, msg):

#     req_params = {

#         'apikey': 'DF5Z7DF4ZOQMLT2ZLKSQG3Q1HVVY63YV',

#         'secret': 'O9PUY8CXTS85S7KE',

#         'usetype': 'stage',

#         'phone': number_to_send,

#         'message': msg,

#         'senderid': 'Spreadsheet'

#     }

#     return requests.post(reqUrl, req_params)

URL = "http://sms.tozzutechnology.com/rest/services/sendSMS/sendGroupSms"


def sendPostRequest(number, msg):

    querystring = {"AUTH_KEY": "de5b88a38524556f59cbcbc1f54fa8", "message": msg,
                   "senderId": "honeym", "routeId": "1", "mobileNos": number, "smsContentType": "english"}

    headers = {
        'Cache-Control': "no-cache"
    }
    a = requests.request("GET", URL, headers=headers, params=querystring)
    print(a.text)
    return a


images_ = []
booleanVal = False
Rows_Sorted_Ascendingly = True
Rows_Sorted_Descendingly = False
to_sort_rows = False
Cols_Sorted_Ascendingly = True
Cols_Sorted_Descendingly = False
to_sort_cols = False


def closest_colour(requested_colour):
    min_colours = {}
    for key, name in webcolors.css3_hex_to_names.items():
        r_c, g_c, b_c = webcolors.hex_to_rgb(key)
        rd = (r_c - requested_colour[0]) ** 2
        gd = (g_c - requested_colour[1]) ** 2
        bd = (b_c - requested_colour[2]) ** 2
        min_colours[(rd + gd + bd)] = name
    return min_colours[min(min_colours.keys())]


def get_colour_name(requested_colour):
    try:
        closest_name = actual_name = webcolors.rgb_to_name(requested_colour)
    except ValueError:
        closest_name = closest_colour(requested_colour)
        actual_name = None
    return actual_name, closest_name


def something(a):
    if (65 + a) > 90:
        b = a + 1
        while b > 26:
            b -= 26
        c = (b) % 66

        if chr(64 + c) == "Z":
            return (((a + 1) // 26))*(chr(64 + c))
        else:
            return (((a + 1) // 26) + 1)*(chr(64 + c))
    else:
        return chr(65 + a)


class InputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.first = QLineEdit(self)
        self.second = QLineEdit(self)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)

        layout = QFormLayout(self)
        layout.addRow("Image Height", self.first)
        layout.addRow("Image Width", self.second)
        layout.addWidget(buttonBox)

        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)


class QImageViewer(QDialog):
    def __init__(self, parent, filename, height, width, idx, iidx, st, name=None):
        super().__init__(parent)
        self.name = name
        # print(self.name)
        self.layout = QVBoxLayout()
        self.imageLabel = QLabel(self)
        self.imageLabel.setBackgroundRole(QPalette.Base)
        self.imageLabel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.imageLabel.setScaledContents(True)
        self.width = width
        self.index = idx
        self.inner_index = iidx
        self.height = height
        self.filename = filename

        image = QImage(self.filename)
        if image.isNull():
            QMessageBox.information(self, "Image Viewer",
                                    "Cannot load %s." % self.filename)
            return

        # self.imageLabel.setPixmap(QPixmap(self.filename).scaled(self.width, self.height))
        self.pxmap = QPixmap(self.filename)
        self.imageLabel.setPixmap(self.pxmap)
        self.cb = QCheckBox("Processed")
        self.cb.setChecked(eval(st))
        self.layout.addWidget(self.cb)
        self.cb.stateChanged.connect(self.btnstate)
        self.layout.addWidget(self.imageLabel)

        if self.name:
            self.setWindowTitle(name)
            self.setWindowTitle("Image and PDF Viewer")

        self.buttonBox = QDialogButtonBox(self)
        self.create_Btn = self.buttonBox.addButton('Print', QDialogButtonBox.ApplyRole)
        self.create_Btn.clicked.connect(self._print_)
        self.layout.addWidget(self.create_Btn)
        self.setLayout(self.layout)
        # self.resize(int(600), int(900))
        # print(int(pxmap.size().width()))
        # print(int(pxmap.size().height()))
        self.resize(int(self.pxmap.size().width()) + 30,
                    int(self.pxmap.size().height()) + 100)

    def _print_(self):
        # printer = QPrinter(QPrinter.ScreenResolution)

        # printer = QPrinter(QPrinter.ScreenResolution)
        # # dlg = QPrintPreviewDialog(printer)
        # painter = QtGui.QPainter()
        # painter.begin(printer)
        # screen = self.imageLabel.grab()
        # # dlg.paintRequested.connect(self.handlePaintRequest)
        # painter.drawPixmap(10, 10, screen)
        # painter.end()
        # view = PrintView()
        # view.setModel(self.table.model())
        # dlg.paintRequested.connect(view.print_)
        # dlg.exec_()
        # Here!!!!!!!!!!!!
        # printer = QPrinter(QPrinter.HighResolution)
        # dialog = QPrintDialog(printer, self)
        # if dialog.exec_() == QPrintDialog.Accepted:
        #     painter = QtGui.QPainter()
        #     painter.begin(printer)
        #     screen = self.imageLabel.grab()
        #     painter.drawPixmap(10, 10, screen)
        #     painter.end()
            # self.layout.print_(printer)
        # printer = QPrinter(QPrinter.ScreenResolution)
        # # printer.setPageSize(QPrinter.Custom)
        # printer.setPaperSize(QSizeF(int(self.pxmap.size().width()) + 30,
        #                             int(self.pxmap.size().height()) + 100), QPrinter.DevicePixel)
        # # printer.setFullPage(True)
        # # self.resize(printer.width(), printer.height())
        # previewDialog = QPrintPreviewDialog(printer, self)
        # previewDialog.resize(int(self.pxmap.size().width()) + 30,
        #                      int(self.pxmap.size().height()) + 100)
        # previewDialog.paintRequested.connect(self.printPreview)
        # previewDialog.exec_()
        self.printer = QPrinter(QPrinter.ScreenResolution)
        if int(self.pxmap.size().width()) > int(self.pxmap.size().height()):
            self.printer.setOrientation(QPrinter.Landscape)
        # dialog = QPrintDialog(self.printer, self)
        dialog = QPrintPreviewDialog(self.printer, self)
        dialog.resize(int(self.pxmap.size().width()) + 30,
                      int(self.pxmap.size().height()) + 100)
        dialog.paintRequested.connect(self.printPreview)
        dialog.exec_()
        # if dialog.exec_():
        #     painter = QPainter(self.printer)
        #     rect = painter.viewport()
        #     size = self.imageLabel.pixmap().size()
        #     size.scale(rect.size(), Qt.KeepAspectRatio)
        #     painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
        #     painter.setWindow(self.imageLabel.pixmap().rect())
        #     painter.drawPixmap(0, 0, self.imageLabel.pixmap())

    def printPreview(self, printer):
        painter = QPainter(self.printer)
        rect = painter.viewport()
        size = self.imageLabel.pixmap().size()
        size.scale(rect.size(), Qt.KeepAspectRatio)
        painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
        painter.setWindow(self.imageLabel.pixmap().rect())
        painter.drawPixmap(0, 0, self.imageLabel.pixmap())
        painter.end()

    def btnstate(self):
        images_[self.index][2][self.inner_index][3] = str(self.cb.isChecked())


class AddImageWidget(QWidget):

    def __init__(self, list_of_imgs, parent=None):
        super(AddImageWidget, self).__init__(parent)

        layout = QHBoxLayout()

        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        for i in list_of_imgs:
            layout.addWidget(ImageWidget(i[2], i[0], i[1], self))

        self.setLayout(layout)


class ImageWidget(QWidget):

    def __init__(self, imagePath, height, width, parent):
        super(ImageWidget, self).__init__(parent)
        self.picture = str(imagePath)

        self.height = height
        self.width = width

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.drawPixmap(0, 0, QPixmap(self.picture).scaled(10, 10))


class SpreadSheet(QMainWindow):

    dateFormats = ["dd/MM/yyyy", "yyyy/MM/dd", "dd.MM.yyyy"]

    currentDateFormat = dateFormats[0]

    # keyPressed = pyqtSignal(QEvent)

    def __init__(self, rows, cols, datecol=None, name=None, titlerow=None, parent=None, phcol=None, emailcol=None):
        super(SpreadSheet, self).__init__(parent)
        self.datecol = datecol
        self.rs = rows
        self.cls = cols
        self.namecol = name
        # self.keyPressed = pyqtSignal(QEvent)
        # print(self.namecol)
        if phcol:
            self.phcol = phcol[0]
        else:
            self.phcol = None

        if emailcol != None:
            self.emailcol = emailcol[0]
        else:
            self.emailcol = None
        # print(self.emailcol)
        # print(self.phcol)
        self.toolBar = QToolBar()
        self.addToolBar(self.toolBar)
        self.formulaInput = QLineEdit()
        self.cellLabel = QLabel(self.toolBar)
        self.cellLabel.setMinimumSize(80, 0)
        self.toolBar.addWidget(self.cellLabel)
        self.toolBar.addWidget(self.formulaInput)
        self.table = QTableWidget(rows, cols, self)
        for c in range(cols):
            character = something(c)
            self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))

        self.table.horizontalHeader().sectionDoubleClicked.connect(self.changeHorizontalHeader)
        self.table.setItemPrototype(self.table.item(rows - 1, cols - 1))
        self.table.setItemDelegate(SpreadSheetDelegate(self, col=datecol, title=titlerow))
        self.createActions()
        self.updateColor(0)
        self.setupMenuBar()
        self.setupContents()
        self.setupContextMenu()
        self.setCentralWidget(self.table)
        self.statusBar()
        self.table.currentItemChanged.connect(self.updateStatus)
        self.table.currentItemChanged.connect(self.updateColor)
        self.table.currentItemChanged.connect(self.updateLineEdit)
        self.table.itemChanged.connect(self.updateStatus)
        self.formulaInput.returnPressed.connect(self.returnPressed)
        self.table.itemChanged.connect(self.updateLineEdit)
        self.setWindowTitle("Spreadsheet")

    def changeHorizontalHeader(self, index):
        it = self.table.horizontalHeaderItem(index)
        if it is None:
            val = self.table.model().headerData(index, Qt.Horizontal)
            it = QTableWidgetItem(str(val))
            self.table.setHorizontalHeaderItem(index, it)
        oldHeader = it.text()
        newHeader, okPressed = QInputDialog.getText(self,
                                                    ' Change header label for column %d', "New heading:",
                                                    QLineEdit.Normal, oldHeader)
        if okPressed:
            it.setText(newHeader)

    def createActions(self):
        self.cell_sumAction = QAction("Sum", self)
        self.cell_sumAction.triggered.connect(self.actionSum)

        self.cell_addAction = QAction("&Add", self)
        self.cell_addAction.setShortcut(Qt.CTRL | Qt.Key_Plus)
        self.cell_addAction.triggered.connect(self.actionAdd)

        self.cell_subAction = QAction("&Subtract", self)
        self.cell_subAction.setShortcut(Qt.CTRL | Qt.Key_Minus)
        self.cell_subAction.triggered.connect(self.actionSubtract)

        self.cell_mulAction = QAction("&Multiply", self)
        self.cell_mulAction.setShortcut(Qt.CTRL | Qt.Key_Asterisk)
        self.cell_mulAction.triggered.connect(self.actionMultiply)

        self.cell_divAction = QAction("&Divide", self)
        self.cell_divAction.setShortcut(Qt.CTRL | Qt.Key_Slash)
        self.cell_divAction.triggered.connect(self.actionDivide)

        self.fontAction = QAction("Font...", self)
        self.fontAction.setShortcut(Qt.CTRL | Qt.Key_F)
        self.fontAction.triggered.connect(self.selectFont)

        self.emailAction = QAction("Email Selected Rows", self)
        self.emailAction.triggered.connect(self.emailrow)

        self.emailoneAction = QAction("Email Selected", self)
        self.emailoneAction.triggered.connect(self.emailselected)

        self.rowimport = QAction("Import Row via Row number", self)
        self.rowimport.triggered.connect(partial(self.importfrmdb, 0))

        self.textimport = QAction("Import Rows with similar Text", self)
        self.textimport.triggered.connect(partial(self.importfrmdb, 1))

        self.dtimport = QAction("Import Row via Date", self)
        self.dtimport.triggered.connect(partial(self.importfrmdb, 2))

        self.allimport = QAction("Import Full Database", self)
        self.allimport.triggered.connect(partial(self.importfrmdb, 3))

        self.imgAction = QAction("View Images and PDFs", self)
        self.imgAction.setShortcut(Qt.CTRL | Qt.Key_Space)
        self.imgAction.triggered.connect(self.view_selected)

        self.viewAction = QAction("Add Images and PDFs", self)
        self.viewAction.setShortcut(Qt.CTRL | Qt.Key_I)
        self.viewAction.triggered.connect(self.file_open)

        self.delRows = QAction("Delete Selected Rows", self)
        self.delRows.triggered.connect(self.delextraRow)

        self.delColumns = QAction("Delete Selected Columns", self)
        self.delColumns.triggered.connect(self.delextraCol)

        self.colorAction = QAction(QIcon(QPixmap(16, 16)), "Background &Color...", self)
        self.colorAction.triggered.connect(self.selectColor)

        self.clearAction = QAction("Clear", self)
        self.clearAction.setShortcut(Qt.Key_Delete)
        self.clearAction.triggered.connect(self.clear)

        self.fiterbynameAction = QAction("Filter by Text.", self)
        self.fiterbynameAction.triggered.connect(self.filterbyname)

        self.fiterbydateAction = QAction("Filter by Date.", self)
        self.fiterbydateAction.triggered.connect(self.filterbydate)

        self.fiterbynumberAction = QAction("Filter by Number.", self)
        self.fiterbynumberAction.triggered.connect(self.filterbynumber)

        self.fiterbyrowAction = QAction("Filter by Row.", self)
        self.fiterbyrowAction.triggered.connect(self.filterbyrow)

        self.clearfilterAction = QAction("Clear Filter.", self)
        self.clearfilterAction.triggered.connect(self.clearFilter)

        self.sortrowsbyascendingAction = QAction("Sort Rows in Ascending.", self)
        self.sortrowsbyascendingAction.triggered.connect(partial(self.sortbyrow, 0))

        self.sortrowsbydescendingAction = QAction("Sort Rows in Descending.", self)
        self.sortrowsbydescendingAction.triggered.connect(partial(self.sortbyrow, 1))

        self.sortcolsbyascendingAction = QAction("Sort Columns in Ascending.", self)
        self.sortcolsbyascendingAction.triggered.connect(partial(self.sortbycol, 0))

        self.sortcolsbydescendingAction = QAction("Sort Columns in Descending.", self)
        self.sortcolsbydescendingAction.triggered.connect(partial(self.sortbycol, 1))

        self.viewUnprocessAction = QAction("View all unprocessed items.", self)
        self.viewUnprocessAction.triggered.connect(self.viewUnprocess)

        self.viewprocessAction = QAction("View all processed items.", self)
        self.viewprocessAction.triggered.connect(self.viewprocess)

        self.aboutSpreadSheet = QAction("About Spreadsheet", self)
        self.aboutSpreadSheet.triggered.connect(self.showAbout)

        self.addRows = QAction("Add Rows", self)
        self.addRows.triggered.connect(self.addextraRow)

        self.addColumns = QAction("Add Columns", self)
        self.addColumns.triggered.connect(self.addextraCol)

        self.exitAction = QAction("E&xit", self)
        self.exitAction.setShortcut(QKeySequence.Quit)
        self.exitAction.triggered.connect(QApplication.instance().quit)

        self.printAction = QAction("&Print", self)
        self.printAction.setShortcut(QKeySequence.Print)
        self.printAction.triggered.connect(self.print_)

        self.savecsvAction = QAction("&Save", self)
        self.savecsvAction.setShortcut(Qt.CTRL | Qt.Key_S)
        self.savecsvAction.triggered.connect(self.handleSave)

        self.loadcsvAction = QAction("&Open", self)
        self.loadcsvAction.setShortcut(Qt.CTRL | Qt.Key_O)
        self.loadcsvAction.triggered.connect(self.handleOpen)

        self.opencsvAction = QAction("&Open from .CSV file...", self)
        self.opencsvAction.triggered.connect(self.handleOpenCSV)

        self.openExcelAction = QAction("&Open from .XLSX or .XLS file...", self)
        self.openExcelAction.triggered.connect(self.handleOpenXL)

        self.firstSeparator = QAction(self)
        self.firstSeparator.setSeparator(True)

        self.secondSeparator = QAction(self)
        self.secondSeparator.setSeparator(True)

        self.thirdSeparator = QAction(self)
        self.thirdSeparator.setSeparator(True)

        self.fourthSeparator = QAction(self)
        self.fourthSeparator.setSeparator(True)

        self.fifthSeparator = QAction(self)
        self.fifthSeparator.setSeparator(True)

        self.sixthSeparator = QAction(self)
        self.sixthSeparator.setSeparator(True)

    def setupMenuBar(self):
        self.fileMenu = self.menuBar().addMenu("&File")
        self.dateFormatMenu = self.fileMenu.addMenu("&Date format")
        self.dateFormatGroup = QActionGroup(self)
        for f in self.dateFormats:
            action = QAction(f, self, checkable=True,
                             triggered=self.changeDateFormat)
            self.dateFormatGroup.addAction(action)
            self.dateFormatMenu.addAction(action)
            if f == self.currentDateFormat:
                action.setChecked(True)

        self.fileMenu.addAction(self.printAction)
        self.fileMenu.addAction(self.savecsvAction)
        self.fileMenu.addAction(self.loadcsvAction)
        self.fileMenu.addSeparator()
        self.fileMenu.addAction(self.opencsvAction)
        self.fileMenu.addAction(self.openExcelAction)
        self.fileMenu.addSeparator()
        self.fileMenu.addAction(self.exitAction)
        self.cellMenu = self.menuBar().addMenu("&Cell")
        self.cellMenu.addAction(self.cell_addAction)
        self.cellMenu.addAction(self.cell_subAction)
        self.cellMenu.addAction(self.cell_mulAction)
        self.cellMenu.addAction(self.cell_divAction)
        self.cellMenu.addAction(self.cell_sumAction)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.colorAction)
        self.cellMenu.addAction(self.fontAction)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.rowimport)
        self.cellMenu.addAction(self.textimport)
        self.cellMenu.addAction(self.dtimport)
        self.cellMenu.addAction(self.allimport)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.delRows)
        self.cellMenu.addAction(self.delColumns)
        self.cellMenu.addSeparator()
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.imgAction)
        self.cellMenu.addAction(self.viewAction)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.emailAction)
        self.cellMenu.addAction(self.emailoneAction)
        self.cellMenu.addSeparator()
        self.cellMenu.addAction(self.clearAction)
        self.menuBar().addSeparator()
        self.filterMenu = self.menuBar().addMenu("&Filter\\Sort")
        self.filterMenu.addAction(self.fiterbynameAction)
        self.filterMenu.addAction(self.fiterbydateAction)
        self.filterMenu.addAction(self.fiterbynumberAction)
        self.filterMenu.addAction(self.fiterbyrowAction)
        self.filterMenu.addAction(self.clearfilterAction)
        self.filterMenu.addSeparator()
        self.filterMenu.addAction(self.sortrowsbyascendingAction)
        self.filterMenu.addAction(self.sortrowsbydescendingAction)
        self.filterMenu.addAction(self.sortcolsbyascendingAction)
        self.filterMenu.addAction(self.sortcolsbydescendingAction)
        self.filterMenu.addSeparator()
        self.filterMenu.addAction(self.viewUnprocessAction)
        self.filterMenu.addAction(self.viewprocessAction)
        self.menuBar().addSeparator()
        self.addExtra = self.menuBar().addMenu("&Add Fields")
        self.addExtra.addAction(self.addRows)
        self.addExtra.addAction(self.addColumns)
        self.menuBar().addSeparator()
        self.aboutMenu = self.menuBar().addMenu("&Help")
        self.aboutMenu.addAction(self.aboutSpreadSheet)

    def addextraRow(self):
        intgr, ok = QInputDialog.getInt(self, "Add Rows",
                                        "Enter number of rows you want to add:", 25, 0, 100, 1)

        if ok:
            inewHeader, okPressed = QInputDialog.getText(self,
                                                         ' Set Default Value %d', "Default Value to Set:",
                                                         QLineEdit.Normal, "Nil")
            if okPressed:
                for i in range(intgr):
                    currentRowCount = self.table.rowCount()
                    self.table.insertRow(currentRowCount)
                    for c in range(self.table.columnCount()):
                        self.table.setItem(self.table.rowCount()-1, c,
                                           SpreadSheetItem(str(inewHeader)))
                        self.table.item(self.table.rowCount()-1,
                                        c).setBackground(Qt.white)
            else:
                for i in range(intgr):
                    currentRowCount = self.table.rowCount()
                    self.table.insertRow(currentRowCount)
                    for c in range(self.table.columnCount()):
                        self.table.setItem(self.table.rowCount()-1, c,
                                           SpreadSheetItem(""))
                        self.table.item(self.table.rowCount()-1,
                                        c).setBackground(Qt.white)

    def addextraCol(self):
        intgr, ok = QInputDialog.getInt(self, "Add Columns",
                                        "Enter number of columns you want to add:", 25, 0, 100, 1)

        if ok:
            inewHeader, okPressed = QInputDialog.getText(self,
                                                         ' Set Default Value %d', "Default Value to Set:",
                                                         QLineEdit.Normal, "Nil")
            if okPressed:
                prev_cols = self.table.columnCount()
                for i in range(intgr):
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)
                    for r in range(self.table.rowCount()):
                        self.table.setItem(r, self.table.columnCount()-1,
                                           SpreadSheetItem(str(inewHeader)))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)

                cols = self.table.columnCount()
                for c in range(prev_cols, cols):
                    character = something(c)
                    self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))
            else:
                prev_cols = self.table.columnCount()
                for i in range(intgr):
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)
                    for r in range(self.table.rowCount()):
                        self.table.setItem(r, self.table.columnCount()-1,
                                           SpreadSheetItem(""))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)

                cols = self.table.columnCount()
                for c in range(prev_cols, cols):
                    character = something(c)
                    self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))

    def delextraRow(self):
        for i in self.table.selectedItems():
            try:
                self.table.removeRow(self.table.row(i))
            except:
                pass

    def delextraCol(self):
        for i in self.table.selectedItems():
            try:
                self.table.removeColumn(self.table.column(i))
            except:
                pass

    def changeDateFormat(self):
        action = self.sender()
        dtFormats = ["dd/MM/yyyy", "yyyy/MM/dd", "dd.MM.yyyy"]
        newFormat = self.currentDateFormat = action.text()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, self.datecol)
            try:
                if '/' in item.text():
                    if len(item.text().split('/')[0]) == 2:
                        oldFormat = dtFormats[0]
                    else:
                        oldFormat = dtFormats[1]
                else:
                    oldFormat = dtFormats[2]
                date = QDate.fromString(item.text(), oldFormat)
            except:
                pass
            try:
                item.setText(date.toString(newFormat))
            except:
                pass

    def updateStatus(self, item):
        if item and item == self.table.currentItem():
            self.statusBar().showMessage(item.data(Qt.StatusTipRole), 1000)
            self.cellLabel.setText("Cell: (%s)" % encode_pos(self.table.row(item),
                                                             self.table.column(item)))

    def updateColor(self, item):
        pixmap = QPixmap(16, 16)
        color = QColor()
        if item:
            pass
        if not color.isValid():
            color = self.palette().base().color()
        painter = QPainter(pixmap)
        painter.fillRect(0, 0, 16, 16, color)
        lighter = color.lighter()
        painter.setPen(lighter)
        painter.drawPolyline(QPoint(0, 15), QPoint(0, 0), QPoint(15, 0))
        painter.setPen(color.darker())
        painter.drawPolyline(QPoint(1, 15), QPoint(15, 15), QPoint(15, 1))
        painter.end()
        self.colorAction.setIcon(QIcon(pixmap))

    def updateLineEdit(self, item):
        if item != self.table.currentItem():
            return
        if item:
            self.formulaInput.setText(item.data(Qt.EditRole))
        else:
            self.formulaInput.clear()

    def returnPressed(self):
        text = self.formulaInput.text()
        row = self.table.currentRow()
        col = self.table.currentColumn()
        item = self.table.item(row, col)
        if not item:
            self.table.setItem(row, col, SpreadSheetItem(text))
        else:
            item.setData(Qt.EditRole, text)
        self.table.viewport().update()

    def selectColor(self):
        item = self.table.currentItem()
        color = item and QColor(item.background()) or self.table.palette().base().color()
        color = QColorDialog.getColor(color, self)
        if not color.isValid():
            return
        selected = self.table.selectedItems()
        if not selected:
            return
        for i in selected:
            i and i.setBackground(color)
        self.updateColor(self.table.currentItem())

    def selectFont(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        font, ok = QFontDialog.getFont(self.font(), self)
        if not ok:
            return
        for i in selected:
            i and i.setFont(font)

    def emailthething(self, images, body, the_to_send):
        try:
            EMAIL_ID = "pina.tax.bhavnagar@gmail.com"
            EMAIL_PASS = "wqvzlprqpdmxjinz"
            msg = EmailMessage()
            msg['Subject'] = 'This is your Information'
            msg['From'] = EMAIL_ID
            msg['To'] = ", ".join(the_to_send)
            msg.set_content(body)
            if images != None:
                for img in images:
                    try:
                        with open(img, 'rb') as f:
                            file_data = f.read()
                            file_type = imghdr.what(f.name)
                            file_name = f.name
                        msg.add_attachment(file_data, maintype='image',
                                           subtype=file_type, filename=file_name)
                    except:
                        pass
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ID, EMAIL_PASS)

                smtp.send_message(msg)
        except:
            pass

    def emailrow(self):
        try:
            selected = self.table.selectedItems()
            if not selected:
                return
            else:
                # text, ok = QInputDialog.getText(self, "Email Selected Rows",
                #                                 "Enter email address:", QLineEdit.Normal, "")
                # if ok:
                body = '|||Column: Information|||\n\n'
                images_to_mail = []
                text = []
                # print("hi")
                for rn, r in enumerate(selected):
                    # print("Wassup")
                    z = 0
                    for c in range(self.table.columnCount()):
                        if (self.table.item(self.table.row(r), c).text() != "" and z == 0) or (self.table.cellWidget(self.table.row(r), c) != None and z == 0):
                            z += 1
                            # print(z)
                            if self.emailcol != None:
                                text1, ok = QInputDialog.getText(self, "Email Selected Rows",
                                                                 "Enter email address:", QLineEdit.Normal, str(self.table.item(self.table.row(r), self.emailcol).text()))
                            else:
                                text1, ok = QInputDialog.getText(self, "Email Selected Rows",
                                                                 "Enter email address:", QLineEdit.Normal, "")
                            if ok:
                                text.append(text1)
                            else:
                                return
                            body += f"|{str(self.table.horizontalHeaderItem(c).text())}: {str(self.table.item(self.table.row(r), c).text())} |\n"
                        elif z != 0:
                            break
                selected = [self.table.row(r) for r in selected]
                for imgs in images_:
                    if imgs[0] in selected:
                        for img in imgs[2]:
                            images_to_mail.append(img[2])
                if len(images_to_mail) != 0:
                    body += "\n|||P.S. Images attached with the mail|||"

                self.emailthething(images_to_mail, body, text)

        except Exception as e:
            print(str(e))
            raise

    def emailselected(self):
        try:
            selected = self.table.selectedItems()
            if not selected:
                return
            else:
                # text, ok = QInputDialog.getText(self, "Email Selected Rows",
                #                                 "Enter email address:", QLineEdit.Normal, "")
                # if ok:
                # print("hi")
                body = '|||Column: Information|||\n\n'
                images_to_mail = []
                text = []
                for rn, r in enumerate(selected):
                    z = 0
                    for cn, c in enumerate(selected):
                        # print("asds")
                        if (self.table.item(self.table.row(r), self.table.column(c)).text() != "" and z == 0) or (self.table.cellWidget(self.table.row(r), self.table.column(c)) != None and z == 0):
                            # z += 1
                            if self.emailcol != None:
                                text1, ok = QInputDialog.getText(self, "Email Selected",
                                                                 "Enter email address:", QLineEdit.Normal, str(self.table.item(self.table.row(r), self.emailcol).text()))
                            else:
                                text1, ok = QInputDialog.getText(self, "Email Selected",
                                                                 "Enter email address:", QLineEdit.Normal, "")
                            if ok:
                                text.append(text1)
                            else:
                                return
                            body += f"|{str(self.table.horizontalHeaderItem(self.table.column(c)).text())}: {str(self.table.item(self.table.row(r), self.table.column(c)).text())} |\n"
                        elif z != 0:
                            break
                selectedr = [self.table.row(r) for r in selected]
                selectedc = [self.table.column(c) for c in selected]
                for imgs in images_:
                    if imgs[0] in selectedr and imgs[1] in selectedc:
                        for img in imgs[2]:
                            images_to_mail.append(img[2])

                if len(images_to_mail) != 0:
                    body += "\n|||P.S. Images attached with the mail|||"

                self.emailthething(images_to_mail, body, text)

        except:
            pass
    # def keyPressEvent(self, event):
    #     super(SpreadSheet, self).keyPressEvent(event)
    #     self.keyPressed.emit(event)

    # def view_allUnprocess(self, event, img, i, idx_img, idx_i):
    #     if event.key() == Qt.Key_Tab:
    #         if self.namecol != None:
    #             self.win = QImageViewer(self, i[2], i[1], i[0], images_.index(img), img[2].index(
    #                 i), i[3], f"{str(self.table.item(img[0], self.namecol).text())} - {str(self.table.horizontalHeaderItem(img[1]).text())}")
    #         else:
    #             self.win = QImageViewer(
    #                 self, i[2], i[1], i[0], images_.index(img), img[2].index(i), i[3], str(self.table.horizontalHeaderItem(img[1]).text()))
    #         self.win.show()
    #         if idx_img+1 < len(images_):
    #             if idx_i+1 == len(img[2]):
    #                 idx_i = 0
    #                 idx_img += 1
    #                 self.viewUnprocess(idx_img, idx_i)
    #             else:
    #                 idx_i += 1
    #                 self.viewUnprocess(idx_img, idx_i)
    #         else:
    #             pass

    def viewprocess(self):
        for img in images_:
            for i in img[2]:
                if eval(i[3]):
                    if self.namecol != None:
                        self.win = QImageViewer(self, i[2], i[1], i[0], images_.index(img), img[2].index(
                            i), i[3], f"{str(self.table.item(img[0], self.namecol).text())} - {str(self.table.horizontalHeaderItem(img[1]).text())}")
                    else:
                        self.win = QImageViewer(
                            self, i[2], i[1], i[0], images_.index(img), img[2].index(i), i[3], str(self.table.horizontalHeaderItem(img[1]).text()))
                    self.win.show()

    def viewUnprocess(self, idx_img=0, idx_i=0):
        for img in images_:
            for i in img[2]:
                if not eval(i[3]):
                    if self.namecol != None:
                        self.win = QImageViewer(self, i[2], i[1], i[0], images_.index(img), img[2].index(
                            i), i[3], f"{str(self.table.item(img[0], self.namecol).text())} - {str(self.table.horizontalHeaderItem(img[1]).text())}")
                    else:
                        self.win = QImageViewer(
                            self, i[2], i[1], i[0], images_.index(img), img[2].index(i), i[3], str(self.table.horizontalHeaderItem(img[1]).text()))
                    self.win.show()

    def view_selected(self):
        # print(self.namecol)
        selected = self.table.selectedItems()
        if not selected:
            pass
        else:
            for select in selected:
                r = self.table.row(select)
                c = self.table.column(select)
                if self.table.cellWidget(r, c) != None:
                    for img in images_:
                        if img[0] == r and img[1] == c:
                            for i in img[2]:
                                if self.namecol != None:
                                    self.win = QImageViewer(self, i[2], i[1], i[0], images_.index(img), img[2].index(
                                        i), i[3], f"{str(self.table.item(img[0], self.namecol).text())} - {str(self.table.horizontalHeaderItem(img[1]).text())}")
                                else:
                                    self.win = QImageViewer(
                                        self, i[2], i[1], i[0], images_.index(img), img[2].index(i), i[3], str(self.table.horizontalHeaderItem(img[1]).text()))
                                self.win.show()
                                # self.win.open()
                                # self.win.exec_()
                            # keyboard.wait('tab')

                else:
                    pass

    def file_open(self):
        global mv_frwrd
        names, oks = QFileDialog.getOpenFileNames(
            self, 'Open File', '')

        selected = self.table.selectedItems()
        if not selected:
            return
        if not oks:
            return
        for name in names:
            name_of_img = name.split('/')[-1]
            # dialog = InputDialog()
            # dialog.setWindowTitle(f'{name_of_img}')
            # dialog.exec()
            if name.endswith(".pdf"):
                # print(name)
                try:
                    try:
                        pages = convert_from_path(name, 300)
                    except:
                        text, ok = QInputDialog.getText(self, "Password",
                                                        "Enter password:", QLineEdit.Normal, "password")
                        output = "no_pass_pdf.pdf"
                        with open(name, 'rb') as input_file, \
                                open(output, 'wb') as output_file:
                            reader = PdfFileReader(input_file)
                            reader.decrypt(text)

                            writer = PdfFileWriter()

                            for i in range(reader.getNumPages()):
                                writer.addPage(reader.getPage(i))
                            writer.write(output_file)
                        pages = convert_from_path(output, 300)

                    for i, page in enumerate(pages):
                        page.save(f"{name_of_img}-{i}.jpg", 'JPEG')
                        name = f"{name_of_img}-{i}.jpg"
                        try:
                            for i in selected:
                                mv_frwrd = True

                                if len(images_):
                                    for img in images_:
                                        if img[0] == self.table.row(i) and img[1] == self.table.column(i):
                                            # print(name)
                                            img[2].append([int(842),
                                                           int(595), name, "False"])
                                            mv_frwrd = False
                                            break
                                        else:
                                            pass

                                    if mv_frwrd:
                                        # print(name)
                                        images_.append([self.table.row(i), self.table.column(i), [[int(
                                            842), int(595), name, "False"]]])
                                else:
                                    images_.append([self.table.row(i), self.table.column(i), [[int(
                                        842), int(595), name, "False"]]])

                                for img in images_:
                                    if img[0] == self.table.row(i) and img[1] == self.table.column(i):
                                        self.table.setCellWidget(self.table.row(i), self.table.column(i),
                                                                 AddImageWidget(img[2], self))

                            # os.remove(f"{name_of_img}.jpg")
                        except:
                            pass
                except Exception as e:
                    # print(str(e))
                    pass

            else:
                try:
                    for i in selected:
                        mv_frwrd = True

                        if len(images_):
                            for img in images_:
                                if img[0] == self.table.row(i) and img[1] == self.table.column(i):
                                    img[2].append([int(595),
                                                   int(842), name, "False"])
                                    mv_frwrd = False
                                    break
                                else:
                                    pass

                            if mv_frwrd:
                                images_.append([self.table.row(i), self.table.column(i), [[int(
                                    595), int(842), name, "False"]]])
                        else:
                            images_.append([self.table.row(i), self.table.column(i), [[int(
                                595), int(842), name, "False"]]])

                        for img in images_:
                            if img[0] == self.table.row(i) and img[1] == self.table.column(i):
                                self.table.setCellWidget(self.table.row(i), self.table.column(i),
                                                         AddImageWidget(img[2], self))
                except:
                    pass

    def loadCsv(self, fileName):
        items = []
        with open(fileName, "r") as fileInput:
            try:
                for row in csv.reader(fileInput):
                    item_row = [
                        QtGui.QStandardItem(field)
                        for field in row
                    ]
                    items.append(item_row)
            except:
                pass
        return np.array(items)

    def loadExcel(self, fileName):
        try:
            items = []
            wb = xlrd.open_workbook(fileName)
            sheet = wb.sheet_by_index(0)

            for r in range(sheet.nrows):
                items.append(sheet.row_values(r))

            return np.array(items)

        except:
            pass

    def handleOpenCSV(self):
        name, oks = QFileDialog.getOpenFileName(
            self, 'Open File', '', 'CSV(*.csv)')

        if not oks:
            return
        if name == "":
            return
        else:
            self.setupContents()
            items = self.loadCsv(name)

            if len(items):
                z_ = 0
                for i in items:
                    if len(i) > z_:
                        z_ = len(i)
                while self.table.columnCount() < z_:
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)

                colz = self.table.columnCount()
                for c in range(colz):
                    character = something(c)
                    self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))

                while self.table.rowCount() < len(items):
                    currentRowCount = self.table.rowCount()
                    self.table.insertRow(currentRowCount)

                for r in range(len(items)):
                    for c in range(z_):
                        try:
                            to_insert = items[r][c].text()
                            if type(to_insert) == float:
                                to_insert = int(to_insert)
                            self.table.setItem(r, c, SpreadSheetItem(to_insert))
                        except:
                            pass

            else:
                return

    def handleOpenXL(self):
        name, oks = QFileDialog.getOpenFileName(
            self, 'Open File', '')

        if not oks:
            return
        if name == "":
            return
        else:
            self.setupContents()
            items = self.loadExcel(name)
            try:
                if len(items):
                    z_ = 0
                    for i in items:
                        if len(i) > z_:
                            z_ = len(i)
                    while self.table.columnCount() < z_:
                        currentColumnCount = self.table.columnCount()
                        self.table.insertColumn(currentColumnCount)

                    colz = self.table.columnCount()
                    for c in range(colz):
                        character = something(c)
                        self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))

                    while self.table.rowCount() < len(items):
                        currentRowCount = self.table.rowCount()
                        self.table.insertRow(currentRowCount)

                    for r in range(len(items)):
                        for c in range(z_):
                            try:
                                to_insert = items[r][c]
                                try:
                                    to_insert = to_insert.astype(np.float)
                                    to_insert = to_insert.astype(np.int64)
                                    to_insert = str(to_insert)
                                except Exception:
                                    pass
                                self.table.setItem(r, c, SpreadSheetItem(to_insert))
                            except:
                                pass

                else:
                    return
            except:
                pass

    def handleSave(self, cloud=False):
        global booleanVal, cursor, conn, nameofdb, db_folder
        # z = 0
        to_mail_list = []
        try:
            if not booleanVal:
                nameofdb, oks = QFileDialog.getSaveFileName(
                    self, 'Save Database')
                if not oks:
                    return
                if nameofdb != "":
                    db_folder = nameofdb
                    os.mkdir(nameofdb)
                    Images_folder = f"{nameofdb}/Images"
                    os.mkdir(Images_folder)
                    conn = sqlite3.connect(f"{nameofdb}/{nameofdb.split('/')[-1]}.db")
                    cursor = conn.cursor()
                    cursor.execute(
                        "CREATE TABLE IF NOT EXISTS texTable(row BLOB, col BLOB, tdata BLOB)")
                    cursor.execute(
                        "CREATE TABLE IF NOT EXISTS ImgTable(row BLOB, col BLOB, idata BLOB)")
                    cursor.execute(
                        "CREATE TABLE IF NOT EXISTS BgTable(row BLOB, col BLOB, k0data BLOB, k1data BLOB, k2data BLOB, k3data BLOB)")
                    cursor.execute(
                        "CREATE TABLE IF NOT EXISTS HeadingTable(col BLOB, heading BLOB)")
                    cursor.execute(
                        "CREATE TABLE IF NOT EXISTS SpecialTable(dtcol BLOB, phncol BLOB, emailcol BLOB)")
                    try:
                        cursor.execute("INSERT INTO SpecialTable(dtcol, phncol, emailcol) VALUES (?, ?, ?)",
                                       (str(self.datecol), str(self.phcol), str(self.emailcol)))

                    except Exception as e:
                        # print(str(e))
                        pass

                    for r in range(self.table.rowCount()):
                        Empty_Cells = []
                        for c in range(self.table.columnCount()):
                            if self.table.item(r, c).text() != "" and self.table.cellWidget(r, c) != None:
                                t = self.table.item(r, c).text()
                                cursor.execute("INSERT INTO texTable(row, col, tdata) VALUES (?, ?, ?)",
                                               (r, c, t))

                            # elif z == 0 and get_colour_name(self.table.item(r, c).background().color().getRgb()[:-1])[1] == "white" and self.table.cellWidget(r, c) == None:
                            #     try:
                            #         z += 1
                            #         num = self.table.item(r, self.phcol).text()
                            #         sendPostRequest(URL, 'provided-api-key', 'provided-secret',
                            #                         'prod/stage', 'valid-to-mobile', 'active-sender-id', 'message-text', num, f"Empty cell at {str(self.table.horizontalHeaderItem(c).text())}")

                            #     except Exception as e:
                            #         print(str(e))
                            #         pass
                            else:
                                Empty_Cells.append(self.table.horizontalHeaderItem(
                                    c).text())

                            l = self.table.item(r, c).background().color().getRgb()
                            if get_colour_name(l[:-1])[1] != "white":
                                cursor.execute("INSERT INTO BgTable(row, col, k0data, k1data, k2data, k3data) VALUES (?, ?, ?, ?, ?, ?)",
                                               (r, c, l[0], l[1], l[2], l[3]))
                        if len(Empty_Cells) != 0:
                            # print(str(self.table.item(r, int(self.emailcol)).text()))
                            if self.phcol != None and self.namecol != None and str(self.table.item(r, int(self.phcol)).text()) != "":
                                to_mail_list.append(
                                    [Empty_Cells, str(self.table.item(r, int(self.phcol)).text()), str(self.table.item(r, int(self.namecol)).text())])
                            else:
                                if self.phcol != None and str(self.table.item(r, int(self.phcol)).text()) != "":
                                    to_mail_list.append(
                                        [Empty_Cells, str(self.table.item(r, int(self.phcol)).text()), ""])
                                else:
                                    pass

                    for img_ in images_:
                        rimg = img_[0]
                        cimg = img_[1]
                        imgi = list(img_[2])
                        for i in imgi:
                            himg = i[0]
                            wimg = i[1]
                            imgcw = f"./Images/{i[2].split('/')[-1]}"
                            pimg = i[3]
                            try:
                                shutil.copy(i[2], f"{nameofdb}/Images/")
                            except Exception as e:
                                # print(str(e))
                                pass
                            # idx = img_[2].index(i)
                            # img_[2].pop(idx)

                            # img_[2].insert(idx, [himg, wimg, imgcw, pimg])
                            idx = imgi.index(i)
                            imgi.pop(idx)

                            imgi.insert(idx, [himg, wimg, imgcw, pimg])
                        cursor.execute("INSERT INTO ImgTable(row, col, idata) VALUES (?, ?, ?)",
                                       (rimg, cimg, str(imgi)))

                    for i in range(self.table.columnCount()):
                        it = self.table.horizontalHeaderItem(i).text()
                        cursor.execute("INSERT INTO HeadingTable(col, heading) VALUES (?, ?)",
                                       (i, it))

                    conn.commit()
                    for i in to_mail_list:
                        body = f"Dear {i[2]} your information in the Columns, "
                        for empty in i[0]:
                            body += str(empty) + "  "
                        body += "is required please send it to Shripalbhai(PINA TAX)."
                        sendPostRequest(i[1], body)
                    booleanVal = True

                else:
                    return

            else:
                z = 0
                cursor.execute("DELETE FROM texTable")
                cursor.execute("DELETE FROM BgTable")
                cursor.execute("DELETE FROM ImgTable")
                cursor.execute("DELETE FROM HeadingTable")
                cursor.execute("DELETE FROM SpecialTable")
                try:
                    cursor.execute("INSERT INTO SpecialTable(dtcol, phncol, emailcol) VALUES (?, ?, ?)",
                                   (str(self.datecol), str(self.phcol), str(self.emailcol)))

                except Exception as e:
                    # print(str(e))
                    pass
                for r in range(self.table.rowCount()):
                    Empty_Cells = []
                    for c in range(self.table.columnCount()):
                        if self.table.item(r, c).text() != "" and self.table.cellWidget(r, c) != None:
                            t = self.table.item(r, c).text()
                            cursor.execute("INSERT INTO texTable(row, col, tdata) VALUES (?, ?, ?)",
                                           (r, c, t))

                        # elif z == 0 and get_colour_name(self.table.item(r, c).background().color().getRgb()[:-1])[1] == "white" and self.table.cellWidget(r, c) == None:
                        #     try:
                        #         z += 1
                        #         num = self.table.item(r, self.phcol).text()
                        #         sendPostRequest(URL, 'provided-api-key', 'provided-secret',
                        #                             'prod/stage', 'valid-to-mobile', 'active-sender-id', 'message-text', num, f"Empty cell at {str(self.table.horizontalHeaderItem(c).text())}")

                        #     except:
                        #         pass
                        else:
                            Empty_Cells.append(self.table.horizontalHeaderItem(
                                c).text())
                        l = self.table.item(r, c).background().color().getRgb()
                        if get_colour_name(l[:-1])[1] != "white":
                            cursor.execute("INSERT INTO BgTable(row, col, k0data, k1data, k2data, k3data) VALUES (?, ?, ?, ?, ?, ?)",
                                           (r, c, l[0], l[1], l[2], l[3]))

                    if len(Empty_Cells) != 0:
                        # print(str(self.table.item(r, int(self.emailcol)).text()))
                        if self.phcol != None and self.namecol != None and str(self.table.item(r, int(self.phcol)).text()) != "":
                            to_mail_list.append(
                                [Empty_Cells, str(self.table.item(r, int(self.phcol)).text()), str(self.table.item(r, int(self.namecol)).text())])
                        else:
                            if self.phcol != None and str(self.table.item(r, int(self.phcol)).text()) != "":
                                to_mail_list.append(
                                    [Empty_Cells, str(self.table.item(r, int(self.phcol)).text()), ""])
                            else:
                                pass

                for img_ in images_:
                    rimg = img_[0]
                    cimg = img_[1]
                    imgi = list(img_[2])
                    for i in imgi:
                        himg = i[0]
                        wimg = i[1]
                        imgcw = f"./Images/{i[2].split('/')[-1]}"
                        pimg = i[3]
                        try:
                            shutil.copy(i[2], f"{db_folder}/Images/")
                        except:
                            pass
                        # idx = img_[2].index(i)
                        # img_[2].pop(idx)

                        # img_[2].insert(idx, [himg, wimg, imgcw, pimg])
                        idx = imgi.index(i)
                        imgi.pop(idx)

                        imgi.insert(idx, [himg, wimg, imgcw, pimg])

                    cursor.execute("INSERT INTO ImgTable(row, col, idata) VALUES (?, ?, ?)",
                                   (rimg, cimg, str(imgi)))

                for i in range(self.table.columnCount()):
                    it = self.table.horizontalHeaderItem(i).text()
                    cursor.execute("INSERT INTO HeadingTable(col, heading) VALUES (?, ?)",
                                   (i, it))
                conn.commit()
                # print(to_mail_list)
                for i in to_mail_list:
                    body = f"Dear {i[2]} your "
                    for empty in i[0]:
                        body += str(empty) + "  "
                    body += "is required please send it to Shripalbhai(PINA TAX)."
                    sendPostRequest(i[1], body)

        except Exception as e:
            # print(str(e))
            pass

    def additivecolsandrows(self, cursor_):
        cursor_.execute(
            "SELECT * FROM texTable")
        txt_data = [row for row in cursor_.fetchall()]
        if len(txt_data):
            txt_rows_ = [row[0] for row in txt_data]
            txt_cols_ = [row[1] for row in txt_data]
            max_txt_row = np.max(txt_rows_)
            max_txt_col = np.max(txt_cols_)
        else:
            txt_rows_ = []
            txt_cols_ = []
            max_txt_row = 0
            max_txt_col = 0
        cursor_.execute(
            "SELECT * FROM BgTable")
        bg_data = [row for row in cursor_.fetchall()]
        if len(bg_data):
            bg_rows_ = [row[0] for row in bg_data]
            bg_cols_ = [row[1] for row in bg_data]
            max_bg_row = np.max(bg_rows_)
            max_bg_col = np.max(bg_cols_)
        else:
            bg_rows_ = []
            bg_cols_ = []
            max_bg_row = 0
            max_bg_col = 0

        cursor_.execute(
            "SELECT * FROM ImgTable")
        img_data = [row for row in cursor_.fetchall()]
        if len(img_data):
            img_rows_ = [row[0] for row in img_data]
            img_cols_ = [row[1] for row in img_data]
            max_img_row = np.max(img_rows_)
            max_img_col = np.max(img_cols_)
        else:
            img_rows_ = []
            img_cols_ = []
            max_img_row = 0
            max_img_col = 0

        rowsmx = np.argmax([max_txt_row, max_bg_row, max_img_row])
        colsmx = [max_txt_col, max_bg_col, max_img_col]

        self.selected = self.table.selectedItems()
        if len(self.selected) > 1:
            self.selected = [self.selected[0]]

        try:
            if rowsmx == 0:
                rows_ = txt_rows_
            elif rowsmx == 1:
                rows_ = bg_rows_
            elif rowsmx == 2:
                rows_ = img_rows_
        except:
            rows_ = []
        prev_cols = self.table.columnCount()

        for i in range(np.max(colsmx)):
            currentColumnCount = self.table.columnCount()
            self.table.insertColumn(currentColumnCount)
            for r in range(self.table.rowCount()):
                self.table.setItem(
                    r, self.table.columnCount()-1, SpreadSheetItem(""))
                self.table.item(r, self.table.columnCount() -
                                1).setBackground(Qt.white)

        cols = self.table.columnCount()
        for c in range(prev_cols, cols):
            character = something(c)
            self.table.setHorizontalHeaderItem(
                c, QTableWidgetItem(character))

        rc = self.table.rowCount()
        if self.selected:
            for i in range(abs((len(set(rows_))+2)-(rc-self.table.row(self.selected[0])))):
                currentRowCount = self.table.rowCount()
                self.table.insertRow(currentRowCount)
                for c in range(self.table.columnCount()):
                    self.table.setItem(
                        self.table.rowCount()-1, c, SpreadSheetItem(""))
                    self.table.item(self.table.rowCount()-1,
                                    c).setBackground(Qt.white)
        else:
            for i in range(len(set(rows_))+2):
                currentRowCount = self.table.rowCount()
                self.table.insertRow(currentRowCount)
                for c in range(self.table.columnCount()):
                    self.table.setItem(
                        self.table.rowCount()-1, c, SpreadSheetItem(""))
                    self.table.item(self.table.rowCount()-1,
                                    c).setBackground(Qt.white)
        return rc

    def importingtxt(self, rows_, cols_, data):
        self.selected = self.table.selectedItems()
        if len(self.selected) > 1:
            self.selected = [self.selected[0]]
        prev_cols = self.table.columnCount()
        for i in range(max(cols_)):
            currentColumnCount = self.table.columnCount()
            self.table.insertColumn(currentColumnCount)
            try:
                for r in range(self.table.rowCount()):
                    try:
                        self.table.setItem(
                            r, self.table.columnCount()-1, SpreadSheetItem(""))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)

                    except:
                        pass
            except:
                pass
        cols = self.table.columnCount()
        for c in range(prev_cols, cols):
            try:
                character = something(c)
                self.table.setHorizontalHeaderItem(
                    c, QTableWidgetItem(character))
            except:
                pass
        rc = self.table.rowCount()
        for i in range(len(set(rows_))):
            try:
                currentRowCount = self.table.rowCount()
                self.table.insertRow(currentRowCount)
                for c in range(self.table.columnCount()):
                    try:
                        self.table.setItem(self.table.rowCount()-1,
                                           c, SpreadSheetItem(""))
                        self.table.item(self.table.rowCount()-1,
                                        c).setBackground(Qt.white)
                    except:
                        pass
            except:
                pass
        if self.selected:
            if len(set(rows_)) == 1:
                for txtdatum in data:
                    try:
                        self.table.setItem(self.table.row(self.selected[0]), txtdatum[1], SpreadSheetItem(f"{txtdatum[2]}"))
                    except:
                        pass
            else:
                for ri in range(len(set(rows_))):
                    for txtdatum in data:
                        try:
                            self.table.setItem(self.table.row(self.selected[0]) + ri, txtdatum[1], SpreadSheetItem(f"{txtdatum[2]}"))
                        except:
                            pass
        else:
            if len(set(rows_)) == 1:
                for txtdatum in data:
                    try:
                        self.table.setItem(rc, txtdatum[1], SpreadSheetItem(f"{txtdatum[2]}"))
                    except:
                        pass
            else:
                for ri in range(len(set(rows_))):
                    for txtdatum in data:
                        try:
                            self.table.setItem(rc + ri, txtdatum[1], SpreadSheetItem(f"{txtdatum[2]}"))
                        except:
                            pass

        return (list(set(rows_)), rc)

    def importingbg(self, rs_, cursor_, row):
        try:
            self.selected = self.table.selectedItems()
            if len(self.selected) > 1:
                self.selected = [self.selected[0]]
            if len(rs_) == 1:
                cursor_.execute(
                    f"SELECT * FROM BgTable WHERE row = {rs_[0]}")
            else:
                cursor_.execute(
                    f"SELECT * FROM BgTable WHERE row IN {tuple(rs_)}")
            _data_ = [row for row in cursor_.fetchall()]
            _cols_ = [row[1] for row in _data_]
            if np.max(_cols_) > self.table.columnCount():
                prev_cols = self.table.columnCount()
                for i in range(np.max(_cols_)):
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)
                    for r in range(self.table.rowCount()):
                        self.table.setItem(
                            r, self.table.columnCount()-1, SpreadSheetItem(""))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)
                cols = self.table.columnCount()
                for c in range(prev_cols, cols):
                    character = something(c)
                    self.table.setHorizontalHeaderItem(
                        c, QTableWidgetItem(character))

            if self.selected:
                if len(rs_) == 1:
                    for bgdatum in _data_:
                        try:
                            self.table.item(self.table.row(self.selected[0]), bgdatum[1]).setBackground(
                                QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
                        except:
                            pass
                else:
                    for ri, datum in enumerate(_data_):
                        for bgdatum in datum:
                            try:
                                self.table.item(self.table.row(self.selected[0])+ri, bgdatum[1]).setBackground(
                                    QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
                            except:
                                pass
            else:
                if len(rs_) == 1:
                    for bgdatum in _data_:
                        try:
                            self.table.item(row, bgdatum[1]).setBackground(
                                QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
                        except:
                            pass
                else:
                    for ri, datum in enumerate(_data_):
                        for bgdatum in datum:
                            try:
                                self.table.item(row+ri, bgdatum[1]).setBackground(
                                    QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
                            except:
                                pass
        except:
            pass

    def importingimg(self, rs_, cursor_, row, name):
        try:
            self.selected = self.table.selectedItems()
            if len(self.selected) > 1:
                self.selected = [self.selected[0]]
            if len(rs_) == 1:
                cursor_.execute(
                    f"SELECT * FROM ImgTable WHERE row = {rs_[0]}")
            else:
                cursor_.execute(
                    f"SELECT * FROM ImgTable WHERE row IN {tuple(rs_)}")
            data = [row for row in cursor_.fetchall()]
            rwdata = [row[0] for row in data]
            _cols_ = [row[1] for row in data]
            prev_cols = self.table.columnCount()
            if np.max(_cols_) > self.table.columnCount():
                for i in range(np.max(_cols_)):
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)
                    for r in range(self.table.rowCount()):
                        self.table.setItem(
                            r, self.table.columnCount()-1, SpreadSheetItem(""))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)
                for c in range(prev_cols, self.table.columnCount()):
                    character = something(c)
                    self.table.setHorizontalHeaderItem(
                        c, QTableWidgetItem(character))
            if self.selected:
                if len(list(set(rwdata))) == 1:
                    # print(data)
                    for imgdatum in data:
                        try:
                            imgd_ = eval(imgdatum[2])
                            # print(imgd_)
                            img_list = []
                            for imgd in imgd_:
                                db_folder = name.split('/')[:-1]
                                db_folder = '/'.join(db_folder)
                                pic = str(db_folder) + imgd[2][1:]
                                img_list.append([imgd[0], imgd[1], pic, imgd[3]])

                            self.table.setCellWidget(self.table.row(self.selected[0]), imgdatum[1],
                                                     AddImageWidget(img_list, self))
                            images_.append(list((self.table.row(self.selected[0]), imgdatum[1],
                                                 imgdatum[2])))
                        except:
                            pass
                else:
                    for ri, datum in enumerate(list(set(rwdata))):
                        for imgdatum in datum:
                            try:
                                imgd_ = eval(imgdatum[2])
                                img_list = []
                                for imgd in imgd_:
                                    db_folder = name.split('/')[:-1]
                                    db_folder = '/'.join(db_folder)
                                    pic = str(db_folder) + imgd[2][1:]
                                    img_list.append([imgd[0], imgd[1], pic, "False"])

                                self.table.setCellWidget(self.table.row(self.selected[0])+ri, imgdatum[1],
                                                         AddImageWidget(img_list, self))
                                images_.append(list((self.table.row(self.selected[0])+ri, imgdatum[1],
                                                     imgdatum[2])))
                            except:
                                pass
            else:
                if len(list(set(rwdata))) == 1:
                    for imgdatum in data:
                        try:
                            imgd_ = eval(imgdatum[2])
                            img_list = []
                            for imgd in imgd_:
                                db_folder = name.split('/')[:-1]
                                db_folder = '/'.join(db_folder)
                                pic = str(db_folder) + imgd[2][1:]
                                img_list.append([imgd[0], imgd[1], pic, imgd[3]])

                            self.table.setCellWidget(row, imgdatum[1],
                                                     AddImageWidget(img_list, self))
                            images_.append(list((row, imgdatum[1],
                                                 imgdatum[2])))
                        except:
                            pass
                else:
                    for ri, datum in enumerate(list(set(rwdata))):
                        for imgdatum in datum:
                            try:
                                imgd_ = eval(imgdatum[2])
                                img_list = []
                                for imgd in imgd_:
                                    db_folder = name.split('/')[:-1]
                                    db_folder = '/'.join(db_folder)
                                    pic = str(db_folder) + imgd[2][1:]
                                    img_list.append([imgd[0], imgd[1], pic, imgd[3]])

                                self.table.setCellWidget(row+ri, imgdatum[1],
                                                         AddImageWidget(img_list, self))
                                images_.append(list((row+ri, imgdatum[1],
                                                     imgdatum[2])))
                            except:
                                pass
        except:
            pass

    def importfrmdb(self, type):
        try:
            name, oks = QFileDialog.getOpenFileName(
                self, 'Open Database', '', 'DB(*.db)')
            if not oks:
                return
            if name != "":

                conn_ = sqlite3.connect(name)
                cursor_ = conn_.cursor()

                if type == 1:
                    try:
                        text, ok = QInputDialog.getText(self, "Import Rows with similar Text",
                                                        "Enter text:", QLineEdit.Normal, "Enter Text")

                        self.ok = ok
                        if self.ok:
                            self.text = text
                            cursor.execute(
                                f"SELECT * FROM texTable")
                            txt_data__ = np.array([row for row in cursor.fetchall()])
                            # print(txt_data__)
                            rows = []
                            for row in txt_data__:
                                # print(text.strip().lower())
                                # print(str(row[2]).lower())
                                if text.strip().lower() in str(row[2]).lower():
                                    rows.append(int(row[0]))

                            data = np.array(rows)

                            if len(data) == 1:
                                cursor_.execute(
                                    f"SELECT * FROM texTable WHERE row = {data[0]}")
                                txt_data = [row for row in cursor_.fetchall()]
                                if len(txt_data):
                                    txt_rows_ = [row[0] for row in txt_data]
                                    txt_cols_ = [row[1] for row in txt_data]
                                    max_txt_row = np.max(txt_rows_)
                                    max_txt_col = np.max(txt_cols_)
                                else:
                                    txt_rows_ = []
                                    txt_cols_ = []
                                    max_txt_row = 0
                                    max_txt_col = 0
                                    cursor_.execute(
                                        f"SELECT * FROM BgTable WHERE row = {data[0]}")
                                bg_data = [row for row in cursor_.fetchall()]
                                if len(bg_data):
                                    bg_rows_ = [row[0] for row in bg_data]
                                    bg_cols_ = [row[1] for row in bg_data]
                                    max_bg_row = np.max(bg_rows_)
                                    max_bg_col = np.max(bg_cols_)
                                else:
                                    bg_rows_ = []
                                    bg_cols_ = []
                                    max_bg_row = 0
                                    max_bg_col = 0

                                cursor_.execute(
                                    f"SELECT * FROM ImgTable WHERE row = {data[0]}")
                                img_data = [row for row in cursor_.fetchall()]
                                if len(img_data):
                                    img_rows_ = [row[0] for row in img_data]
                                    img_cols_ = [row[1] for row in img_data]
                                    max_img_row = np.max(img_rows_)
                                    max_img_col = np.max(img_cols_)
                                else:
                                    img_rows_ = []
                                    img_cols_ = []
                                    max_img_row = 0
                                    max_img_col = 0

                                rowsmx = np.argmax([max_txt_row, max_bg_row, max_img_row])
                                colsmx = [max_txt_col, max_bg_col, max_img_col]

                                try:
                                    if rowsmx == 0:
                                        rows_ = txt_rows_
                                    elif rowsmx == 1:
                                        rows_ = bg_rows_
                                    elif rowsmx == 2:
                                        rows_ = img_rows_
                                except:
                                    rows_ = []
                                rs_, rc = self.importingtxt(rows_, colsmx, txt_data)
                                self.importingbg(rs_, cursor_, rc)

                                self.importingimg(rs_, cursor_, rc, name)
                            else:
                                for datum in data:
                                    cursor_.execute(
                                        f"SELECT * FROM texTable WHERE row = {datum}")
                                    txt_data = [row for row in cursor_.fetchall()]
                                    if len(txt_data):
                                        txt_rows_ = [row[0] for row in txt_data]
                                        txt_cols_ = [row[1] for row in txt_data]
                                        max_txt_row = np.max(txt_rows_)
                                        max_txt_col = np.max(txt_cols_)
                                    else:
                                        txt_rows_ = []
                                        txt_cols_ = []
                                        max_txt_row = 0
                                        max_txt_col = 0
                                    cursor_.execute(
                                        f"SELECT * FROM BgTable WHERE row = {datum}")
                                    bg_data = [row for row in cursor_.fetchall()]
                                    if len(bg_data):
                                        bg_rows_ = [row[0] for row in bg_data]
                                        bg_cols_ = [row[1] for row in bg_data]
                                        max_bg_row = np.max(bg_rows_)
                                        max_bg_col = np.max(bg_cols_)
                                    else:
                                        bg_rows_ = []
                                        bg_cols_ = []
                                        max_bg_row = 0
                                        max_bg_col = 0

                                    cursor_.execute(
                                        f"SELECT * FROM ImgTable WHERE row = {datum}")
                                    img_data = [row for row in cursor_.fetchall()]
                                    if len(img_data):
                                        img_rows_ = [row[0] for row in img_data]
                                        img_cols_ = [row[1] for row in img_data]
                                        max_img_row = np.max(img_rows_)
                                        max_img_col = np.max(img_cols_)
                                    else:
                                        img_rows_ = []
                                        img_cols_ = []
                                        max_img_row = 0
                                        max_img_col = 0

                                    rowsmx = np.argmax(
                                        [max_txt_row, max_bg_row, max_img_row])
                                    colsmx = [max_txt_col, max_bg_col, max_img_col]

                                    try:
                                        if rowsmx == 0:
                                            rows_ = txt_rows_
                                        elif rowsmx == 1:
                                            rows_ = bg_rows_
                                        elif rowsmx == 2:
                                            rows_ = img_rows_
                                    except:
                                        rows_ = []
                                    rs_, rc = self.importingtxt(rows_, colsmx, txt_data)
                                    self.importingbg(rs_, cursor_, rc)

                                    self.importingimg(rs_, cursor_, rc, name)
                        else:
                            return
                    except:
                        pass

                elif type == 0:
                    intgr, ok = QInputDialog.getInt(self, "Import Rows",
                                                    "Enter which numbered row you want to import:", 25, 0, 100, 1)
                    if ok:
                        try:
                            cursor_.execute(
                                f"SELECT * FROM texTable WHERE row = {intgr-1}")
                            txt_data = [row for row in cursor_.fetchall()]
                            if len(txt_data):
                                txt_rows_ = [row[0] for row in txt_data]
                                txt_cols_ = [row[1] for row in txt_data]
                                max_txt_row = np.max(txt_rows_)
                                max_txt_col = np.max(txt_cols_)
                            else:
                                txt_rows_ = []
                                txt_cols_ = []
                                max_txt_row = 0
                                max_txt_col = 0
                            cursor_.execute(
                                f"SELECT * FROM BgTable WHERE row = {intgr-1}")
                            bg_data = [row for row in cursor_.fetchall()]
                            if len(bg_data):
                                bg_rows_ = [row[0] for row in bg_data]
                                bg_cols_ = [row[1] for row in bg_data]
                                max_bg_row = np.max(bg_rows_)
                                max_bg_col = np.max(bg_cols_)
                            else:
                                bg_rows_ = []
                                bg_cols_ = []
                                max_bg_row = 0
                                max_bg_col = 0
                            cursor_.execute(
                                f"SELECT * FROM ImgTable WHERE row = {intgr-1}")
                            img_data = [row for row in cursor_.fetchall()]
                            if len(img_data):
                                img_rows_ = [row[0] for row in img_data]
                                img_cols_ = [row[1] for row in img_data]
                                max_img_row = np.max(img_rows_)
                                max_img_col = np.max(img_cols_)
                            else:
                                img_rows_ = []
                                img_cols_ = []
                                max_img_row = 0
                                max_img_col = 0

                            rowsmx = np.argmax([max_txt_row, max_bg_row, max_img_row])
                            colsmx = [max_txt_col, max_bg_col, max_img_col]

                            try:
                                if rowsmx == 0:
                                    rows_ = txt_rows_
                                elif rowsmx == 1:
                                    rows_ = bg_rows_
                                elif rowsmx == 2:
                                    rows_ = img_rows_
                            except:
                                rows_ = []
                            rs_, rc = self.importingtxt(rows_, colsmx, txt_data)
                            self.importingbg(rs_, cursor_, rc)

                            self.importingimg(rs_, cursor_, rc, name)
                        except:

                            pass
                    else:
                        return

                elif type == 2:
                    text, ok = QInputDialog.getText(self, "Import Rows using Date",
                                                    "Enter date:", QLineEdit.Normal, "Enter Date")
                    if ok:
                        try:
                            dtFormats = ["dd/MM/yyyy", "yyyy/MM/dd", "dd.MM.yyyy"]
                            if '/' in text:
                                if len(text.split('/')[0]) == 2:
                                    crntdt = dtFormats[0]
                                else:
                                    crntdt = dtFormats[1]
                            else:
                                crntdt = dtFormats[2]
                            cursor_.execute(
                                f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                            rows = [row[0] for row in cursor_.fetchall()]
                            while not len(rows) or not len(dtFormats):
                                dtFormats.remove(crntdt)
                                date = QDate.fromString(text, crntdt)
                                text = date.toString(dtFormats[0])
                                cursor_.execute(
                                    f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                                rows = [row[0] for row in cursor_.fetchall()]
                                crntdt = dtFormats[0]
                            data = rows
                            if len(data) == 1:
                                cursor_.execute(
                                    f"SELECT * FROM texTable WHERE row = {data[0]}")
                                txt_data = [row for row in cursor_.fetchall()]
                                if len(txt_data):
                                    txt_rows_ = [row[0] for row in txt_data]
                                    txt_cols_ = [row[1] for row in txt_data]
                                    max_txt_row = np.max(txt_rows_)
                                    max_txt_col = np.max(txt_cols_)
                                else:
                                    txt_rows_ = []
                                    txt_cols_ = []
                                    max_txt_row = 0
                                    max_txt_col = 0
                                cursor_.execute(
                                    f"SELECT * FROM BgTable WHERE row = {data[0]}")
                                bg_data = [row for row in cursor_.fetchall()]
                                if len(bg_data):
                                    bg_rows_ = [row[0] for row in bg_data]
                                    bg_cols_ = [row[1] for row in bg_data]
                                    max_bg_row = np.max(bg_rows_)
                                    max_bg_col = np.max(bg_cols_)
                                else:
                                    bg_rows_ = []
                                    bg_cols_ = []
                                    max_bg_row = 0
                                    max_bg_col = 0

                                cursor_.execute(
                                    f"SELECT * FROM ImgTable WHERE row = {data[0]}")
                                img_data = [row for row in cursor_.fetchall()]
                                if len(img_data):
                                    img_rows_ = [row[0] for row in img_data]
                                    img_cols_ = [row[1] for row in img_data]
                                    max_img_row = np.max(img_rows_)
                                    max_img_col = np.max(img_cols_)
                                else:
                                    img_rows_ = []
                                    img_cols_ = []
                                    max_img_row = 0
                                    max_img_col = 0

                                rowsmx = np.argmax([max_txt_row, max_bg_row, max_img_row])
                                colsmx = [max_txt_col, max_bg_col, max_img_col]

                                try:
                                    if rowsmx == 0:
                                        rows_ = txt_rows_
                                    elif rowsmx == 1:
                                        rows_ = bg_rows_
                                    elif rowsmx == 2:
                                        rows_ = img_rows_
                                except:
                                    rows_ = []
                                rs_, rc = self.importingtxt(rows_, colsmx, txt_data)
                                self.importingbg(rs_, cursor_, rc)

                                self.importingimg(rs_, cursor_, rc, name)
                            else:
                                for datum in data:
                                    cursor_.execute(
                                        f"SELECT * FROM texTable WHERE row = {datum}")
                                    txt_data = [row for row in cursor_.fetchall()]
                                    if len(txt_data):
                                        txt_rows_ = [row[0] for row in txt_data]
                                        txt_cols_ = [row[1] for row in txt_data]
                                        max_txt_row = np.max(txt_rows_)
                                        max_txt_col = np.max(txt_cols_)
                                    else:
                                        txt_rows_ = []
                                        txt_cols_ = []
                                        max_txt_row = 0
                                        max_txt_col = 0
                                    cursor_.execute(
                                        f"SELECT * FROM BgTable WHERE row = {datum}")
                                    bg_data = [row for row in cursor_.fetchall()]
                                    if len(bg_data):
                                        bg_rows_ = [row[0] for row in bg_data]
                                        bg_cols_ = [row[1] for row in bg_data]
                                        max_bg_row = np.max(bg_rows_)
                                        max_bg_col = np.max(bg_cols_)
                                    else:
                                        bg_rows_ = []
                                        bg_cols_ = []
                                        max_bg_row = 0
                                        max_bg_col = 0

                                    cursor_.execute(
                                        f"SELECT * FROM ImgTable WHERE row = {datum}")
                                    img_data = [row for row in cursor_.fetchall()]
                                    if len(img_data):
                                        img_rows_ = [row[0] for row in img_data]
                                        img_cols_ = [row[1] for row in img_data]
                                        max_img_row = np.max(img_rows_)
                                        max_img_col = np.max(img_cols_)
                                    else:
                                        img_rows_ = []
                                        img_cols_ = []
                                        max_img_row = 0
                                        max_img_col = 0

                                    rowsmx = np.argmax(
                                        [max_txt_row, max_bg_row, max_img_row])
                                    colsmx = [max_txt_col, max_bg_col, max_img_col]

                                    try:
                                        if rowsmx == 0:
                                            rows_ = txt_rows_
                                        elif rowsmx == 1:
                                            rows_ = bg_rows_
                                        elif rowsmx == 2:
                                            rows_ = img_rows_
                                    except:
                                        rows_ = []
                                    rs_, rc = self.importingtxt(rows_, colsmx, txt_data)
                                    self.importingbg(rs_, cursor_, rc)

                                    self.importingimg(rs_, cursor_, rc, name)
                        except:
                            pass
                    else:
                        return

                elif type == 3:
                    try:
                        rc = self.additivecolsandrows(cursor_)
                        cursor_.execute(
                            "SELECT * FROM texTable")
                        txt_data = [row for row in cursor_.fetchall()]

                        cursor_.execute(
                            "SELECT * FROM BgTable")
                        bg_data = [row for row in cursor_.fetchall()]

                        cursor_.execute(
                            "SELECT * FROM ImgTable")
                        img_data = [row for row in cursor_.fetchall()]

                        rc -= self.table.rowCount()
                        if self.selected:
                            rc = self.table.row(self.selected[0])
                        else:
                            rc = abs(rc) + 2
                        for txtdatum in txt_data:
                            try:
                                self.table.setItem(txtdatum[0] + rc, txtdatum[1],
                                                   SpreadSheetItem(txtdatum[2]))
                                self.table.item(
                                    txtdatum[0] + rc, txtdatum[1]).setBackground(Qt.white)
                            except:
                                pass

                        for bgdatum in bg_data:
                            try:
                                self.table.item(bgdatum[0] + rc, bgdatum[1]).setBackground(
                                    QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
                            except:
                                pass

                        for imgdatum in img_data:
                            try:
                                imgd_ = eval(imgdatum[2])
                                img_list = []
                                for imgd in imgd_:
                                    db_folder = name.split('/')[:-1]
                                    db_folder = '/'.join(db_folder)
                                    pic = str(db_folder) + imgd[2][1:]
                                    img_list.append([imgd[0], imgd[1], pic, imgd[3]])

                                self.table.setCellWidget(imgdatum[0] + rc, imgdatum[1],
                                                         AddImageWidget(img_list, self))
                                images_.append(list((imgdatum[0] + rc, imgdatum[1],
                                                     imgdatum[2])))
                            except:
                                pass

                    except:
                        pass

                else:
                    return

        except:
            pass

    def addrowsandcols(self, rcount, rows_, ccount, cols_):
        try:
            if rcount < np.max(rows_)+1:
                for i in range(np.max(rows_) - rcount + 1):
                    currentRowCount = self.table.rowCount()
                    self.table.insertRow(currentRowCount)
                    for c in range(self.table.columnCount()):
                        self.table.setItem(self.table.rowCount()-1,
                                           c, SpreadSheetItem(""))
                        self.table.item(self.table.rowCount()-1,
                                        c).setBackground(Qt.white)
        except:
            pass
        prev_cols = self.table.columnCount()
        try:
            if ccount < np.max(cols_)+1:
                for i in range(np.max(cols_) - ccount + 1):
                    currentColumnCount = self.table.columnCount()
                    self.table.insertColumn(currentColumnCount)
                    for r in range(self.table.rowCount()):
                        self.table.setItem(
                            r, self.table.columnCount()-1, SpreadSheetItem(""))
                        self.table.item(r, self.table.columnCount() -
                                        1).setBackground(Qt.white)
        except:
            pass
        cols = self.table.columnCount()
        for c in range(prev_cols, cols):
            character = something(c)
            self.table.setHorizontalHeaderItem(c, QTableWidgetItem(character))

        self.setupContents()

    def add_to(self, cursor, db_folder):
        cursor.execute(
            "SELECT * FROM SpecialTable")
        sp_data = [datum for datum in cursor.fetchall()]
        # print(sp_data)
        self.datecol = eval(sp_data[0][0])
        self.table.setItemDelegate(SpreadSheetDelegate(self, col=self.datecol))
        if sp_data[0][1] == "None":
            self.phcol = eval(sp_data[0][1])
        else:
            self.phcol = eval(sp_data[0][1])
        if sp_data[0][2] == "None":
            self.emailcol = eval(sp_data[0][2])
        else:
            self.emailcol = eval(sp_data[0][2])
        cursor.execute(
            "SELECT * FROM texTable")
        txt_data = [row for row in cursor.fetchall()]
        txt_rows_ = [row[0] for row in txt_data]
        txt_cols_ = [row[1] for row in txt_data]
        rcount = self.table.rowCount()
        ccount = self.table.columnCount()
        if len(txt_data) != 0:
            self.addrowsandcols(rcount, txt_rows_, ccount, txt_cols_)
        cursor.execute(
            "SELECT * FROM BgTable")
        bg_data = [row for row in cursor.fetchall()]
        bg_rows_ = [row[0] for row in bg_data]
        bg_cols_ = [row[1] for row in bg_data]
        rcount = self.table.rowCount()
        ccount = self.table.columnCount()
        if len(bg_data) != 0:
            self.addrowsandcols(rcount, bg_rows_, ccount, bg_cols_)
        cursor.execute(
            "SELECT * FROM ImgTable")
        img_data = [row for row in cursor.fetchall()]
        img_rows_ = [row[0] for row in img_data]
        img_cols_ = [row[1] for row in img_data]
        rcount = self.table.rowCount()
        ccount = self.table.columnCount()
        if len(img_data) != 0:
            self.addrowsandcols(rcount, img_rows_, ccount, img_cols_)
        cursor.execute(
            "SELECT * FROM HeadingTable")
        header_data = [row for row in cursor.fetchall()]
        header_cols_ = [row[0] for row in header_data]
        rcount = self.table.rowCount()
        ccount = self.table.columnCount()
        if len(header_data) != 0:
            self.addrowsandcols(rcount, img_rows_, ccount, header_cols_)
        for txtdatum in txt_data:
            try:
                self.table.setItem(txtdatum[0], txtdatum[1],
                                   SpreadSheetItem(txtdatum[2]))
                self.table.item(txtdatum[0], txtdatum[1]).setBackground(Qt.white)
            except:

                pass
        for bgdatum in bg_data:
            try:
                self.table.item(bgdatum[0], bgdatum[1]).setBackground(
                    QtGui.QColor(bgdatum[2], bgdatum[3], bgdatum[4], bgdatum[5]))
            except:

                pass

        images_.clear()
        for imgdatum in img_data:
            try:
                imgd_ = eval(imgdatum[2])
                img_list = []
                for imgd in imgd_:
                    pic = str(db_folder) + imgd[2][1:]
                    img_list.append([imgd[0], imgd[1], pic, imgd[3]])
                self.table.setCellWidget(imgdatum[0], imgdatum[1],
                                         AddImageWidget(img_list, self))
                images_.append(list((imgdatum[0], imgdatum[1],
                                     img_list)))
            except:
                pass

        for hdatum in header_data:
            try:
                self.table.horizontalHeaderItem(hdatum[0]).setText(str(hdatum[1]))
                # print(hdatum[0])
                # print(hdatum[1])
            except:
                pass

    def handleOpen(self):
        global booleanVal, cursor, conn, nameofdb, db_name, db_folder
        nameofdb, oks = QFileDialog.getOpenFileName(
            self, 'Open Database', '', 'DB(*.db)')
        if not oks:
            return
        if nameofdb != "":
            db_folder = nameofdb.split('/')[:-1]
            db_folder = "/".join(db_folder)
            db_name = nameofdb.split('/')[-1]
            self.setupContents()
            conn = sqlite3.connect(nameofdb)
            cursor = conn.cursor()
            self.add_to(cursor, db_folder)
            booleanVal = True
        else:
            return

    def runInputDialog(self, title, c1Text, c2Text, opText,
                       outText, cell1, cell2, outCell):

        rows = []
        cols = []
        for r in range(self.table.rowCount()):
            rows.append(str(r + 1))
        for c in range(self.table.columnCount()):
            cols.append(something(c))

        addDialog = QDialog(self)
        addDialog.setWindowTitle(title)
        addDialog.setWhatsThis("Cell Selector")
        group = QGroupBox(title, addDialog)
        group.setMinimumSize(250, 100)
        cell1Label = QLabel(c1Text, group)
        cell1RowInput = QComboBox(group)
        c1Row, c1Col = decode_pos(cell1)
        cell1RowInput.addItems(rows)
        cell1RowInput.setCurrentIndex(c1Row)
        cell1ColInput = QComboBox(group)
        cell1ColInput.addItems(cols)
        cell1ColInput.setCurrentIndex(c1Col)
        operatorLabel = QLabel(opText, group)
        operatorLabel.setAlignment(Qt.AlignHCenter)
        cell2Label = QLabel(c2Text, group)
        cell2RowInput = QComboBox(group)
        c2Row, c2Col = decode_pos(cell2)
        cell2RowInput.addItems(rows)
        cell2RowInput.setCurrentIndex(c2Row)
        cell2ColInput = QComboBox(group)
        cell2ColInput.addItems(cols)
        cell2ColInput.setCurrentIndex(c2Col)
        equalsLabel = QLabel("=", group)
        equalsLabel.setAlignment(Qt.AlignHCenter)
        outLabel = QLabel(outText, group)
        outRowInput = QComboBox(group)
        outRow, outCol = decode_pos(outCell)
        outRowInput.addItems(rows)
        outRowInput.setCurrentIndex(outRow)
        outColInput = QComboBox(group)
        outColInput.addItems(cols)
        outColInput.setCurrentIndex(outCol)

        cancelButton = QPushButton("Cancel", addDialog)
        cancelButton.clicked.connect(addDialog.reject)
        okButton = QPushButton("OK", addDialog)
        okButton.setDefault(True)
        okButton.clicked.connect(addDialog.accept)
        buttonsLayout = QHBoxLayout()
        buttonsLayout.addStretch(1)
        buttonsLayout.addWidget(okButton)
        buttonsLayout.addSpacing(10)
        buttonsLayout.addWidget(cancelButton)

        dialogLayout = QVBoxLayout(addDialog)
        dialogLayout.addWidget(group)
        dialogLayout.addStretch(1)
        dialogLayout.addItem(buttonsLayout)

        cell1Layout = QHBoxLayout()
        cell1Layout.addWidget(cell1Label)
        cell1Layout.addSpacing(10)
        cell1Layout.addWidget(cell1ColInput)
        cell1Layout.addSpacing(10)
        cell1Layout.addWidget(cell1RowInput)

        cell2Layout = QHBoxLayout()
        cell2Layout.addWidget(cell2Label)
        cell2Layout.addSpacing(10)
        cell2Layout.addWidget(cell2ColInput)
        cell2Layout.addSpacing(10)
        cell2Layout.addWidget(cell2RowInput)
        outLayout = QHBoxLayout()
        outLayout.addWidget(outLabel)
        outLayout.addSpacing(10)
        outLayout.addWidget(outColInput)
        outLayout.addSpacing(10)
        outLayout.addWidget(outRowInput)
        vLayout = QVBoxLayout(group)
        vLayout.addItem(cell1Layout)
        vLayout.addWidget(operatorLabel)
        vLayout.addItem(cell2Layout)
        vLayout.addWidget(equalsLabel)
        vLayout.addStretch(1)
        vLayout.addItem(outLayout)
        if addDialog.exec_():
            cell1 = cell1ColInput.currentText() + cell1RowInput.currentText()
            cell2 = cell2ColInput.currentText() + cell2RowInput.currentText()
            outCell = outColInput.currentText() + outRowInput.currentText()
            return True, cell1, cell2, outCell

        return False, cell1, cell2, outCell

    def actionSum(self):
        row_first = 0
        row_last = 0
        row_cur = 0
        col_first = 0
        col_last = 0
        col_cur = 0
        selected = self.table.selectedItems()
        if selected:
            first = selected[0]
            last = selected[-1]
            row_first = self.table.row(first)
            row_last = self.table.row(last)
            col_first = self.table.column(first)
            col_last = self.table.column(last)

        current = self.table.currentItem()
        if current:
            row_cur = self.table.row(current)
            col_cur = self.table.column(current)

        cell1 = encode_pos(row_first, col_first)
        cell2 = encode_pos(row_last, col_last)
        out = encode_pos(row_cur, col_cur)
        ok, cell1, cell2, out = self.runInputDialog("Sum cells", "First cell:",
                                                    "Last cell:",
                                                    u"\N{GREEK CAPITAL LETTER SIGMA}",
                                                    "Output to:",
                                                    cell1, cell2, out)
        if ok:
            row, col = decode_pos(out)
            try:
                self.table.item(row, col).setText("sum %s %s" % (cell1, cell2))
            except:

                self.table.setItem(row, col, SpreadSheetItem(
                    "sum %s %s" % (cell1, cell2)))

    def actionMath_helper(self, title, op):
        cell1 = "C1"
        cell2 = "C2"
        out = "C3"
        current = self.table.currentItem()
        if current:
            out = encode_pos(self.table.currentRow(), self.table.currentColumn())
        ok, cell1, cell2, out = self.runInputDialog(title, "Cell 1", "Cell 2",
                                                    op, "Output to:", cell1, cell2, out)
        if ok:
            row, col = decode_pos(out)
            try:
                self.table.item(row, col).setText("%s %s %s" % (op, cell1, cell2))
            except:

                self.table.setItem(row, col, SpreadSheetItem(
                    "%s %s %s" % (op, cell1, cell2)))

    def actionAdd(self):
        self.actionMath_helper("Addition", "+")

    def actionSubtract(self):
        self.actionMath_helper("Subtraction", "-")

    def actionMultiply(self):
        self.actionMath_helper("Multiplication", "*")

    def actionDivide(self):
        self.actionMath_helper("Division", "/")

    def clear(self):
        for i in self.table.selectedItems():
            i.setText("")
            self.table.setCellWidget(self.table.row(i), self.table.column(i), None)
            self.table.item(self.table.row(i), self.table.column(i)
                            ).setBackground(Qt.white)

    def setupContextMenu(self):
        self.addAction(self.cell_addAction)
        self.addAction(self.cell_subAction)
        self.addAction(self.cell_mulAction)
        self.addAction(self.cell_divAction)
        self.addAction(self.cell_sumAction)
        self.addAction(self.firstSeparator)
        self.addAction(self.colorAction)
        self.addAction(self.fontAction)
        self.addAction(self.secondSeparator)
        self.addAction(self.rowimport)
        self.addAction(self.textimport)
        self.addAction(self.dtimport)
        self.addAction(self.allimport)
        self.addAction(self.thirdSeparator)
        self.addAction(self.delRows)
        self.addAction(self.delColumns)
        self.addAction(self.fourthSeparator)
        self.addAction(self.imgAction)
        self.addAction(self.viewAction)
        self.addAction(self.fifthSeparator)
        self.addAction(self.emailAction)
        self.addAction(self.emailoneAction)
        self.addAction(self.sixthSeparator)
        self.addAction(self.clearAction)
        self.setContextMenuPolicy(Qt.ActionsContextMenu)

    def setupContents(self, DfltVals=""):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                self.table.setItem(r, c, SpreadSheetItem(DfltVals))
                self.table.item(r, c).setBackground(Qt.white)
                self.table.setCellWidget(r, c, None)

    def showAbout(self):
        QMessageBox.about(self, "About Spreadsheet", """
            <HTML>
            <p><b>This is still a demo shows use of <c>The speadsheet</c> ,final project will have cloud integration, sign in and a lot more.</b></p>
            <p>Current Features include:
            <ul>
            <li>Adding two cells.</li>
            <li>Subtracting one cell from another.</li>
            <li>Multiplying two cells.</li>
            <li>Dividing one cell with another.</li>
            <li>Summing the contents of an arbitrary number of cells.</li>
            <li>Adding Images</li>
            <li>Date and Calender function</li>
            <li>Filtering</li>
            </HTML>
        """)

    def filterbyname(self):
        text, ok = QInputDialog.getText(self, "Text Filter",
                                        "Enter the text by which you want to filter the data:", QLineEdit.Normal, "Enter Text")
        if text and ok:
            if not booleanVal:
                self.handleSave()
            try:
                # cursor.execute(
                #     f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                cursor.execute(
                    f"SELECT * FROM texTable")
                txt_data__ = np.array([row for row in cursor.fetchall()])
                # print(txt_data__)
                rows = []
                for row in txt_data__:
                    # print(text.strip().lower())
                    # print(str(row[2]).lower())
                    if text.strip().lower() in str(row[2]).lower():
                        rows.append(int(row[0]))

                rows = np.array(rows)
                # print(rows)
                rcount_ = self.table.rowCount()
                r = 0
                while rcount_ > 0:
                    if r in rows:
                        r += 1
                    else:
                        rows -= 1
                        self.table.removeRow(r)
                    rcount_ -= 1
            except:
                pass

    def filterbydate(self):
        global crntdt
        text, ok = QInputDialog.getText(self, "Date Filter",
                                        "Enter the date by which you want to filter the data:", QLineEdit.Normal, "Enter Date")
        if text and ok:
            if not booleanVal:
                self.handleSave()
            try:
                dtFormats = ["dd/MM/yyyy", "yyyy/MM/dd", "dd.MM.yyyy"]
                if '/' in text:
                    if len(text.split('/')[0]) == 2:
                        crntdt = dtFormats[0]
                    else:
                        crntdt = dtFormats[1]
                else:
                    crntdt = dtFormats[2]
                cursor.execute(
                    f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                rows = [row[0] for row in cursor.fetchall()]
                while not len(rows) or not len(dtFormats):
                    dtFormats.remove(crntdt)
                    date = QDate.fromString(text, crntdt)
                    text = date.toString(dtFormats[0])
                    cursor.execute(
                        f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                    rows = [row[0] for row in cursor.fetchall()]
                    crntdt = dtFormats[0]
                rows = np.array(rows)
                rcount_ = self.table.rowCount()
                r = 0
                while rcount_ > 0:
                    if r in rows:
                        r += 1
                    else:
                        rows -= 1
                        self.table.removeRow(r)
                    rcount_ -= 1
            except:
                pass

    def filterbynumber(self):
        text, ok = QInputDialog.getText(self, "Number Filter",
                                        "Enter the number by which you want to filter the data:", QLineEdit.Normal, "Enter Number")
        if text and ok:
            if not booleanVal:
                self.handleSave()
            try:
                cursor.execute(
                    f"SELECT row FROM texTable WHERE tdata = '{text.strip()}'")
                rows = np.array([row[0] for row in cursor.fetchall()])
                rcount_ = self.table.rowCount()
                r = 0
                while rcount_ > 0:
                    if r in rows:
                        r += 1
                    else:
                        rows -= 1
                        self.table.removeRow(r)
                    rcount_ -= 1
            except:
                pass

    def sortbyrow(self, type):
        global Rows_Sorted_Descendingly, Rows_Sorted_Ascendingly, to_sort_rows
        if type == 0:
            if not Rows_Sorted_Ascendingly:
                to_sort_rows = True
                Rows_Sorted_Ascendingly = True
                Rows_Sorted_Descendingly = False
        elif type == 1:
            if not Rows_Sorted_Descendingly:
                to_sort_rows = True
                Rows_Sorted_Descendingly = True
                Rows_Sorted_Ascendingly = False
        if to_sort_rows:
            rows = [row for row in range(self.table.rowCount())][::-1]
            cols = [col for col in range(self.table.columnCount())]
            txtofrows = [[self.table.item(r, c).text() for c in cols] for r in rows]
            bgofrows = [[self.table.item(r, c).background() for c in cols] for r in rows]
            imgofrows = [img for img in images_]
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    try:
                        self.table.setItem(r, c, SpreadSheetItem(txtofrows[r][c]))
                        self.table.item(r, c).setBackground(bgofrows[r][c])
                        self.table.setCellWidget(r, c, None)
                    except:

                        pass
            for img in imgofrows:
                self.table.setCellWidget(
                    abs(img[0]-self.table.rowCount()+1), img[1], AddImageWidget(img[2], self))
                img[0] = abs(img[0]-self.table.rowCount()+1)
            to_sort_rows = False

    def sortbycol(self, type):
        global Cols_Sorted_Ascendingly, Cols_Sorted_Descendingly, to_sort_cols
        if type == 0:
            if not Cols_Sorted_Ascendingly:
                to_sort_cols = True
                Cols_Sorted_Ascendingly = True
                Cols_Sorted_Descendingly = False
        elif type == 1:
            if not Cols_Sorted_Descendingly:
                to_sort_cols = True
                Cols_Sorted_Descendingly = True
                Cols_Sorted_Ascendingly = False
        if to_sort_cols:
            rows = [row for row in range(self.table.rowCount())]
            cols = [col for col in range(self.table.columnCount())][::-1]
            txtofcols = [[self.table.item(r, c).text() for c in cols] for r in rows]
            bgofcols = [[self.table.item(r, c).background() for c in cols] for r in rows]
            hdrofcols = [self.table.horizontalHeaderItem(
                c).text() for c in cols]
            if self.datecol != None:
                for dt in self.datecol:
                    self.table.setItemDelegate(SpreadSheetDelegate(
                        self, col=abs(dt-self.table.columnCount()) - 1, title=None))
                    dt = abs(dt-self.table.columnCount()) - 1
            if self.namecol != None:
                self.namecol = abs(self.namecol-self.table.columnCount())
            if self.phcol != None:
                self.phcol = abs(self.phcol-self.table.columnCount())
            if self.emailcol != None:
                self.emailcol = abs(self.emailcol-self.table.columnCount())
            for c in range(self.table.columnCount()):
                self.table.horizontalHeaderItem(c).setText(hdrofcols[c])
            imgofcols = [img for img in images_]
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    try:
                        self.table.setItem(r, c, SpreadSheetItem(txtofcols[r][c]))
                        self.table.item(r, c).setBackground(bgofcols[r][c])
                        self.table.setCellWidget(r, c, None)
                    except:
                        pass
            for img in imgofcols:
                self.table.setCellWidget(
                    img[0], abs(img[1]-self.table.rowCount()+1), AddImageWidget(img[2], self))
                img[1] = abs(img[1]-self.table.rowCount()+1)
            to_sort_cols = False

    def clearFilter(self):
        try:
            self.setupContents()
            self.add_to(cursor, db_folder)
        except:
            pass

    def filterbyrow(self):
        text, ok = QInputDialog.getInt(self, "Row Filter",
                                       "Enter the Row by which you want to filter the data:", 25, 0, 100, 1)
        if text and ok:
            if not booleanVal:
                self.handleSave()
            try:
                cursor.execute(
                    f"SELECT row FROM texTable WHERE row = '{text}'")
                rows = np.array([text - 1])
                rcount_ = self.table.rowCount()
                r = 0
                while rcount_ > 0:
                    if r in rows:
                        r += 1
                    else:
                        rows -= 1
                        self.table.removeRow(r)
                    rcount_ -= 1
            except:
                pass

    def print_(self):
        # printer = QPrinter(QPrinter.ScreenResolution)
        # painter = QtGui.QPainter()
        # painter.begin(printer)
        # screen = self.table.grab()
        # painter.drawPixmap(10, 10, screen)
        # painter.end()
        printer = QPrinter(QPrinter.ScreenResolution)
        dlg = QPrintPreviewDialog(printer)
        view = PrintView()
        view.setModel(self.table.model())
        dlg.paintRequested.connect(view.print_)
        dlg.exec_()
