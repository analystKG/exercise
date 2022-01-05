from api.Kiwoom import *
import sys

app = QApplication(sys.argv)
kiwoom = Kiwoom()

position = kiwoom.get_balance()
print(position)

app.exec_()


