from spreadsheet import SpreadSheet
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QApplication
import sys
import eel
import numpy as np

eel.init('web')


def change_to_array(str_to_cols):
    if str_to_cols:
        str_to_cols = np.array(list(map(int, str_to_cols.split(','))))
    try:
        str_to_cols = str_to_cols - 1
        np.where(str_to_cols < 0, None, str_to_cols)
    except:
        str_to_cols = None
    return str_to_cols


@eel.expose
def dummy(rows, cols, dtcols=None, DfltVals=None, phcol=None, emailcol=None):
    if DfltVals != None:
        name = change_to_array(DfltVals)
    if dtcols != None:
        dtcols = change_to_array(dtcols)
    if phcol != None:
        phcol = change_to_array(phcol)
    if emailcol != None:
        emailcol = change_to_array(emailcol)

    try:
        rows = int(rows)
    except:
        rows = 0
    try:
        cols = int(cols)
    except:
        cols = 0

    app = QApplication(sys.argv)
    # print(phcol)
    # print(emailcol)
    sheet = SpreadSheet(rows, cols, datecol=dtcols, name=name,
                        titlerow=None, phcol=phcol, emailcol=emailcol)
    sheet.setWindowIcon(
        QIcon(QPixmap(".\\images\\image.png")))
    sheet.resize(1280, 720)
    sheet.show()
    sys.exit(app.exec_())


# @eel.expose
# def generate_qr(data):
#     img = pyqrcode.create(data)
#     buffers = io.BytesIO()
#     img.png(buffers, scale=8)
#     encoded = b64encode(buffers.getvalue()).decode("ascii")
#     print("QR code generation successful.")
#     return "data:image/png;base64, " + encoded


eel.start('index.html', size=(1080, 720))
