from Kiwoom import *

form_class = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect()

        self.kiwoom.excelfile_initiator()
        self.kospi_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSPI)
        self.kosdaq_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSDAQ)
        self.candidate_codes = ['005930', '252670', '122630', '000660', '207940', '035420', '006400', '035720', '005380', '034730', '036570', '017670', '105560', '096770', '090430', '097950', '018260', '003550', '006800', '078930']

        # There can be limits in the number of timers
        self.timer = QTimer(self)
        self.timer.start(1000*10) # Timer for time update: refresh at every 10*1000ms
        self.timer.timeout.connect(self.timeout) 

        # self.timer_balance_check = QTimer(self)
        # self.timer_balance_check.start(1000*10) 
        # self.timer_balance_check.timeout.connect(self.check_balance)

        self.timer_autotrade_run = QTimer(self)
        self.timer_autotrade_run.start(1000*AUTOTRADE_INTERVAL)
        self.timer_autotrade_run.timeout.connect(self.timeout_autotrade_run)

        accouns_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")
        accounts_list = accounts.split(';')[0:accouns_num]
        self.comboBox.addItems(accounts_list)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order_ui)
        self.pushButton_2.clicked.connect(self.check_balance)
        self.pushButton_3.clicked.connect(self.timeout_autotrade_run)

        self.check_balance()

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm")
        time_msg = "Time: " + text_time

        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "Server Connected"
        else:
            state_msg = "Server Not Connected"
        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def timeout_autotrade_run(self):
        current_time = QTime.currentTime()
        if RUN_AUTOTRADE and not RUN_ANYWAY_OUT_OF_MARKET_OPEN_TIME: 
            if datetime.datetime.today().weekday() in range(0,5) and current_time > MARKET_START_TIME and current_time < MARKET_FINISH_TIME:
                self.autotrade_list_gen()
                self.trade_stocks()
                self.label_8.setText("Autotrade executed")
            else:
                self.label_8.setText("Market not open")
        elif RUN_AUTOTRADE: 
            self.label_8.setText("Autotrade / market may not open")
            self.autotrade_list_gen()
            self.trade_stocks()
        else: 
            self.label_8.setText("Autotrade disabled")
            
    def code_changed(self):
        code = self.lineEdit.text()
        if len(code) >= CODE_MIN_LENGTH:
            name = self.kiwoom.get_master_code_name(code)
            if name != "":
                self.kiwoom.set_input_value("종목코드", code)
                self.kiwoom.comm_rq_data("opt10001_req", "opt10001", 0, "1001")
                self.lineEdit_2.setText(name)
                self.spinBox_2.setValue(self.kiwoom.cur_price)
            else:
                self.lineEdit_2.setText("")
                self.spinBox_2.setValue(0)
        else:
            self.lineEdit_2.setText("")
            self.spinBox_2.setValue(0)

    def send_order_ui(self):
        order_type_lookup = {'Buy': 1, 'Sell': 2}

        account = self.comboBox.currentText()
        order_type = self.comboBox_2.currentText()
        code = self.lineEdit.text()
        hoga = HOGA_LOOKUP[self.comboBox_3.currentText()]
        num = self.spinBox.value()
        if hoga == "00": 
            price = self.spinBox_2.value()
        elif hoga == "03": 
            price = 0
        res = self.kiwoom.send_order("send_order_req", "0101", account, order_type_lookup[order_type], code, num, price, hoga, "")
        if res[0] == 0 and res[1] != "":
            self.label_8.setText("Order sent")
            ###################################
            if self.kiwoom.order_chejan_finished == True:
                self.label_8.setText("Order completed")
        else:
            self.label_8.setText("Errer in order processing")
        
    def check_balance(self):
        account_number = self.kiwoom.get_login_info("ACCNO")
        account_number = account_number.split(';')[0]

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        while self.kiwoom.remained_data:
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

###########################################################################
##### EXECUTION  
###########################################################################

    def trade_stocks(self):
        [buy_list, sell_list] = self.load_buy_sell_list()
        
        # account
        account = self.comboBox.currentText()
        buy_order = 1  
        sell_order = 2
        
        # buy_list
        for i in buy_list.index: 
            if buy_list["Tr"][i] == 'yet' or buy_list["Tr"][i] == 'failed':
                hoga = HOGA_LOOKUP[buy_list["Order_type"][i]]
                if hoga == "00": 
                    price = buy_list["Price"][i]
                elif hoga == "03":
                    price = 0 
                res = self.kiwoom.send_order("send_order_req", "0101", account, buy_order, buy_list["Code"][i], int(buy_list["Amount"][i]), price, hoga,"")

                if res[0] == 0 and res[1] != "":
                    self.label_8.setText("Order sent: " + str(res[1]))
                    buy_list.at[i, "Tr"] = 'ordered'
                    ###################################
                    if self.kiwoom.order_chejan_finished == True:
                        buy_list.at[i, "Tr"] = 'done'
                else:
                    self.label_8.setText("Errer in order processing")
                    buy_list.at[i, "Tr"] = 'failed'

        buy_list.to_excel(EXCEL_BUY_LIST, index=False)

        # sell_list
        for i in sell_list.index: 
            if sell_list["Tr"][i] == 'yet' or sell_list["Tr"][i] == 'failed':
                hoga = HOGA_LOOKUP[sell_list["Order_type"][i]]
                if hoga == "00": 
                    price = sell_list["Price"][i]
                elif hoga == "03":
                    price = 0 
                res = self.kiwoom.send_order("send_order_req", "0101", account, sell_order, sell_list["Code"][i], int(sell_list["Amount"][i]), price, hoga,"")
                if res[0] == 0 and res[1] != "":
                    self.label_8.setText("Order sent: "+str(res[1]))
                    sell_list.at[i, "Tr"] = 'ordered'
                    ###################################
                    if self.kiwoom.order_chejan_finished == True:
                        buy_list.at[i, "Tr"] = 'done'
                else:
                    self.label_8.setText("Errer in order processing")
                    sell_list.at[i, "Tr"] = 'failed'

        sell_list.to_excel(EXCEL_SELL_LIST, index=False)

###########################################################################
##### ALGORITHMS
###########################################################################

    def autotrade_list_gen(self): 

        [code, price] = self.algo_random_choose_buy(3)
        # sell_list = self.algo_random_choose_sell(3)
        sell_list = self.algo_sell_by_return_range(10, -1)

        self.kiwoom.update_buy_list(code, price)
        self.kiwoom.update_sell_list(sell_list)

    def check_speedy_rising_volume(self, code):
        today = datetime.datetime.today().strftime("%Y%m%d")
        df = self.kiwoom.get_ohlcv(code, today)
        volumes = df['volume']

        if len(volumes) < 21:
            return False

        sum_vol20 = 0
        today_vol = 0

        for i, vol in enumerate(volumes):
            if i == 0:
                today_vol = vol
            elif 1 <= i <= 20:
                sum_vol20 += vol
            else:
                break
        avg_vol20 = sum_vol20 / 20
        
        if today_vol > avg_vol20 * 3:
            return True

        return False

    def algo_speedy_rising_volume(self): 
        buy_list_code = []
        for i, code in enumerate(self.kosdaq_codes):
            if self.check_speedy_rising_volume(code): 
                buy_list_code.append(code)
        return buy_list_code

    def algo_random_choose_buy(self, BUY_SELL_SIZE):
        buy_list_code = []
        buy_list_price = []
        for i in range(BUY_SELL_SIZE):
            code = random.choice(self.candidate_codes)
            self.kiwoom.set_input_value("종목코드", code)
            self.kiwoom.comm_rq_data("opt10001_req", "opt10001", 0, "1001")
            name = self.kiwoom.get_master_code_name(code) 
            price = self.kiwoom.cur_price
            if price > MIN_STOCK_PRICE: 
                buy_list_code.append(code)
                buy_list_price.append(price)
        return [buy_list_code, buy_list_price]

    def algo_random_choose_sell(self, BUY_SELL_SIZE):
        my_stock_list = self.kiwoom.get_my_stock_list()
        if len(my_stock_list) > BUY_SELL_SIZE:
            n = list(range(len(my_stock_list))) 
            set_to_sell = random.sample(n, BUY_SELL_SIZE)
            return my_stock_list.iloc[set_to_sell]
        else: 
            return my_stock_list
        
    def algo_sell_by_return_range(self, upperlimit, lowerlimit): # upperlimit, lowerlimit in percentage
        my_stocks = self.kiwoom.get_my_stock_list()
        profit_sell_list = my_stocks[my_stocks['earning_rate'] > upperlimit] 
        loss_sell_list = my_stocks[my_stocks['earning_rate'] < lowerlimit] 
        print('Profit Sell List (up to 50 items): \n', tabulate(profit_sell_list[:50], headers='keys', tablefmt='psql'))
        print('Loss Sell List (up to 50 items): \n', tabulate(loss_sell_list[:50], headers='keys', tablefmt='psql'))
        return profit_sell_list.append(loss_sell_list)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()