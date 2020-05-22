import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from Kiwoom import *
# import pandas as pd

form_class = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.need_to_execute_autotrade = False

        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect()

        self.timer = QTimer(self)
        self.timer.start(1000) # Timer for time update: refresh at every 1000ms
        self.timer.timeout.connect(self.timeout)
        self.timer2 = QTimer(self)
        self.timer2.start(1000*10) # Timer for realtime data tracking (if check box chceked)
        self.timer2.timeout.connect(self.timeout2)

        accouns_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")
        accounts_list = accounts.split(';')[0:accouns_num]
        self.comboBox.addItems(accounts_list)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order_ui)
        self.pushButton_2.clicked.connect(self.check_balance)
        self.pushButton_3.clicked.connect(self.execute_autotrade)

        self.load_buy_sell_list()
        self.check_balance()

    def timeout(self):
        market_start_time = QTime(9, 0, 0)
        market_finish_time = QTime(15,30, 0)
        current_time = QTime.currentTime()

        if current_time > market_start_time and current_time < market_finish_time and self.need_to_execute_autotrade is True:
            self.trade_stocks()
            self.need_to_execute_autotrade = False
            self.label_8.setText("Autotrade executed")
            time.sleep(1)

        text_time = current_time.toString("hh:mm:ss")
        time_msg = "Time: " + text_time

        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "Server Connected"
        else:
            state_msg = "Server Not Connected"

        self.statusbar.showMessage(state_msg + " | " + time_msg)
        if self.kiwoom.chejan_received: 
            self.label_9.setText("Transaction data received")
            self.kiwoon.chejan_received = False
        else: 
            self.label_9.setText("")

    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()
        self.label_7.setText("")

        if self.need_to_execute_autotrade == False: 
            self.label_8.setText("Nothing to autotrade")

    def code_changed(self):
        code = self.lineEdit.text()
        name = self.kiwoom.get_master_code_name(code)
        if name != "":
            self.kiwoom.set_input_value("종목코드", code)
            self.kiwoom.comm_rq_data("opt10001_req", "opt10001", 0, "1001")
            self.lineEdit_2.setText(name)
            self.spinBox_2.setValue(self.kiwoom.cur_price)

    def send_order_ui(self):
        order_type_lookup = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hoga_lookup = {'지정가': "00", '시장가': "03"} # fixed price or market price

        account = self.comboBox.currentText()
        order_type = self.comboBox_2.currentText()
        code = self.lineEdit.text()
        hoga = hoga_lookup[self.comboBox_3.currentText()]
        num = self.spinBox.value()
        if hoga == "00": 
            price = self.spinBox_2.value()
        elif hoga == "03": 
            price = 0
        res = self.kiwoom.send_order("send_order_req", "0101", account, order_type_lookup[order_type], code, num, price, hoga, "")
        if res == 0:
            self.label_7.setText("Order sent")
            time.sleep(1)
        else: 
            self.label_7.setText("Error")
            time.sleep(1)
        
    def check_balance(self):
        account_number = self.kiwoom.get_login_info("ACCNO")
        account_number = account_number.split(';')[0]

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        while self.kiwoom.remained_data:
            time.sleep(TR_REQ_TIME_INTERVAL)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        # opw00001
        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit)
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018_output['single'][i - 1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        # Item list
        item_count = len(self.kiwoom.opw00018_output['multi'])
        self.tableWidget_2.setRowCount(item_count)

        for j in range(item_count):
            row = self.kiwoom.opw00018_output['multi'][j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(j, i, item)

        self.tableWidget_2.resizeRowsToContents()
    
    def execute_autotrade(self):
        self.need_to_execute_autotrade = True
        self.label_8.setText("Autotrade command queued")


    def load_buy_sell_list(self):
        try: 
            buy_list = pd.read_excel(EXCEL_BUY_LIST, index_col=None, converters={'Code': str})
        except Exception as e:
            print(e)
            buy_list = pd.DataFrame()
        
        try:
            sell_list = pd.read_excel(EXCEL_SELL_LIST, index_col=None, converters={'Code': str})
        except Exception as e:
            print(e)
            sell_list = pd.DataFrame()
        
        return [buy_list, sell_list]

    def trade_stocks(self):
        [buy_list, sell_list] = self.load_buy_sell_list()
        
        # account
        account = self.comboBox.currentText()
        buy_order = 1  
        sell_order = 2
        hoga_lookup = {"fixed": "00", "mkt": "03"}
        
        # buy_list
        for i in buy_list.index: 
            if buy_list["Tr"][i] == 'yet':
                hoga = hoga_lookup[buy_list["Order_type"][i]]
                if hoga == "00": 
                    price = buy_list["Price"][i]
                elif hoga == "03":
                    price = 0 
                self.kiwoom.send_order("send_order_req", "0101", account, buy_order, buy_list["Code"][i], int(buy_list["Amount"][i]), price, hoga,"")
                buy_list.at[i, "Tr"] = 'done'

        buy_list.to_excel(EXCEL_BUY_LIST, index=False)

        # sell_list
        for i in sell_list.index: 
            if sell_list["Tr"][i] == 'yet':
                hoga = hoga_lookup[sell_list["Order_type"][i]]
                if hoga == "00": 
                    price = sell_list["Price"][i]
                elif hoga == "03":
                    price = 0 
                self.kiwoom.send_order("send_order_req", "0101", account, sell_order, sell_list["Code"][i], int(sell_list["Amount"][i]), price, hoga,"")
                sell_list.at[i, "Tr"] = 'done'

        sell_list.to_excel(EXCEL_SELL_LIST, index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()