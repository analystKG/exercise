from api.Kiwoom import *
import sys

app = QApplication(sys.argv)
kiwoom = Kiwoom()

order_result = kiwoom.send_order('send_buy_order', '1001', 1, '007700', 1, 35000, '00')
print(order_result)

app.exec_()


