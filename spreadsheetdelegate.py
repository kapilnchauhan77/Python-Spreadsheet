from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import QCompleter, QDateTimeEdit, QItemDelegate, QLineEdit


class SpreadSheetDelegate(QItemDelegate):

    def __init__(self, parent=None, col=[1], title=None):
        super(SpreadSheetDelegate, self).__init__(parent)
        self.col = col
        self.title = title

    def createEditor(self, parent, styleOption, index):
        try:
            if index.column() in self.col:
                if index.row() != self.title:
                    editor = QDateTimeEdit(parent)
                    editor.setDisplayFormat(self.parent().currentDateFormat)
                    editor.setCalendarPopup(True)
                    return editor
        except:
            pass

        try:
            editor = QLineEdit(parent)
            allStrings = []
            for i in range(1, index.model().rowCount()):
                strItem = index.model().data(index.sibling(i, index.column()), Qt.EditRole)
                if strItem not in allStrings:
                    allStrings.append(strItem)
        except:
            pass
        try:
            autoComplete = QCompleter(allStrings)
            editor.setCompleter(autoComplete)
            editor.editingFinished.connect(self.commitAndCloseEditor)
            return editor
        except:
            pass

    def commitAndCloseEditor(self):
        editor = self.sender()
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QItemDelegate.NoHint)

    def setEditorData(self, editor, index):
        if isinstance(editor, QLineEdit):
            editor.setText(index.model().data(index, Qt.EditRole))
        elif isinstance(editor, QDateTimeEdit):
            editor.setDate(QDate.fromString(
                index.model().data(index, Qt.EditRole), self.parent().currentDateFormat))

    def setModelData(self, editor, model, index):
        if isinstance(editor, QLineEdit):
            model.setData(index, editor.text())
        elif isinstance(editor, QDateTimeEdit):
            model.setData(index, editor.date().toString(self.parent().currentDateFormat))
