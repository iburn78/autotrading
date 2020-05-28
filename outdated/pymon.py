import sys
from PyQt5.QtWidgets import *
from Kiwoom import * 
import time
import datetime
import pandas as pd
from pandas import ExcelWriter, ExcelFile
import xlsxwriter
import os.path
from tabulate import tabulate 
import random

ALLOCATION_SIZE = 10000000 # Target amount to be purchased in KRW

class Pymon:
    def __init__(self):
        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect()  
        self.get_code_list()
        self.excelfile_initiator()

    def get_code_list(self):
        self.kospi_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSPI)
        self.kosdaq_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSDAQ)

    def get_ohlcv(self, code, start):
        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.set_input_value("기준일자", start)
        self.kiwoom.set_input_value("수정주가구분", 1)
        self.kiwoom.comm_rq_data("opt10081_req", "opt10081", 0, "0101")

        df = pd.DataFrame(self.kiwoom.ohlcv, columns=['open', 'high', 'low', 'close', 'volume'],
                       index=self.kiwoom.ohlcv['date'])
        return df

    def update_buy_list(self, buy_list_code):
        buy_list = pd.read_excel(EXCEL_BUY_LIST, index_col=None, converters={'Code':str})

        for i, code in enumerate(buy_list_code):
            name = self.kiwoom.get_master_code_name(code) 
            today = datetime.datetime.today().strftime("%Y%m%d")
            time_ = datetime.datetime.now().strftime("%H:%M:%S")
            self.kiwoom.set_input_value("종목코드", code)
            self.kiwoom.comm_rq_data("opt10001_req", "opt10001", 0, "1001")

            amount = round(ALLOCATION_SIZE/self.kiwoom.cur_price)
            buy_list = buy_list.append({'Date': today, 'Time': time_, 'Name': name, 'Code': code, 'Order_type': 'mkt', 'Tr': 'yet', 'Price': self.kiwoom.cur_price, 'Amount': amount }, ignore_index=True)

        print('Buy List: \n', tabulate(buy_list, headers='keys', tablefmt='psql'))

        buy_list.to_excel(EXCEL_BUY_LIST, index=False)

    def get_stock_list(self):
        account_number = self.kiwoom.get_login_info("ACCNO")
        account_number = account_number.split(';')[0]

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")
        
        while self.kiwoom.remained_data:
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        stock_list = pd.DataFrame(self.kiwoom.opw00018_rawoutput, columns=['name', 'code', 'quantity', 'purchase_price', 'current_price', 'invested_amount', 'current_total', 'eval_profit_loss_price', 'earning_rate'])
        for i in stock_list.index:
            stock_list.at[i, 'code'] = stock_list['code'][i][1:]   # taking off "A" in front of returned code
        return stock_list.set_index('code')

    def update_sell_list(self, sell_list_code):
        sell_list = pd.read_excel(EXCEL_SELL_LIST, index_col=None, converters={'Code':str})
        stock_list = self.get_stock_list()
        
        for i, code in enumerate(sell_list_code):
            if code in list(stock_list.index):
                name = self.kiwoom.get_master_code_name(code) 
                today = datetime.datetime.today().strftime("%Y%m%d")
                time_ = datetime.datetime.now().strftime("%H:%M:%S")
                self.kiwoom.set_input_value("종목코드", code)
                self.kiwoom.comm_rq_data("opt10001_req", "opt10001", 0, "1001") # getting current price

                amount = stock_list['quantity'][code]  
                sell_list = sell_list.append({'Date': today, 'Time': time_, 'Name': name, 'Code': code, 'Order_type': 'mkt', 'Tr': 'yet', 'Price': self.kiwoom.cur_price, 'Amount': amount }, ignore_index=True)

        print('Sell List: \n', tabulate(sell_list, headers='keys', tablefmt='psql'))
        sell_list.to_excel(EXCEL_SELL_LIST, index=False)

    def excelfile_initiator(self):
        if not os.path.exists(EXCEL_BUY_LIST): 
            # create buy list
            bl = xlsxwriter.Workbook(EXCEL_BUY_LIST)
            blws = bl.add_worksheet() 
    
            blws.write('A1', 'Date') # Date / Time when the item is added
            blws.write('B1', 'Time') 
            blws.write('C1', 'Name') 
            blws.write('D1', 'Code') 
            blws.write('E1', 'Order_type') # 시장가 ('mkt') vs 지정가 ('fixed')
            blws.write('F1', 'Tr') # yet: not, done: done
            blws.write('G1', 'Price') # latest price when the list is populated
            blws.write('H1', 'Amount') 
            # blws.write('I1', 'Invested_total') # Before any fee and tax 
            # blws.write('J1', 'Date_Trans') # Date / Time when the item is purchased 
            # blws.write('K1', 'Time_Trans') 
            bl.close()
    
        if not os.path.exists(EXCEL_SELL_LIST): 
            # create sell list
            sl = xlsxwriter.Workbook(EXCEL_SELL_LIST)
            slws = sl.add_worksheet() 
    
            slws.write('A1', 'Date') # Date / Time when the item is added
            slws.write('B1', 'Time') 
            slws.write('C1', 'Name') 
            slws.write('D1', 'Code') 
            slws.write('E1', 'Order_type') # 시장가 ('mkt') vs 지정가 ('fixed')
            slws.write('F1', 'Tr') # yet: not, done: done
            slws.write('G1', 'Price') # latest price when the list is populated
            slws.write('H1', 'Amount') # Amount to sell
            # slws.write('I1', 'Fee_Tax')
            # slws.write('J1', 'Harvested_total') # After fee and tax  
            # slws.write('K1', 'Date_Trans') # Date / Time when the item is purchased 
            # slws.write('L1', 'Time_Trans') 
            sl.close()



###########################################################################
##### ALGORITHMS
###########################################################################


    def check_speedy_rising_volume(self, code):
        today = datetime.datetime.today().strftime("%Y%m%d")
        df = self.get_ohlcv(code, today)
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


###########################################################################
##### EXECUTION  
###########################################################################
    def run(self): # has to return True if the list is updated
        current_time = QTime.currentTime()

        # Timer 
        while datetime.datetime.today().weekday() in range(0,5) and current_time > MARKET_START_TIME and current_time < MARKET_FINISH_TIME:
            # Algo
            buy_list_code = []
            buy_list_code.append(random.choice(self.kospi_codes))
            buy_list_code.append(random.choice(self.kospi_codes))
            buy_list_code.append(random.choice(self.kospi_codes))
            buy_list_code.append(random.choice(self.kospi_codes))
            buy_list_code.append(random.choice(self.kospi_codes))

            stock_list_codelist = self.get_stock_list().index.tolist()
            sell_list_code = []
            sell_list_code.append(random.choice(stock_list_codelist))
            sell_list_code.append(random.choice(stock_list_codelist))
            sell_list_code.append(random.choice(stock_list_codelist))
            sell_list_code.append(random.choice(stock_list_codelist))
            sell_list_code.append(random.choice(stock_list_codelist))

            pymon.update_buy_list(buy_list_code)
            pymon.update_sell_list(sell_list_code)

            # Checking 
            time.sleep(AUTOTRADE_INTERVAL)
            current_time = QTime.currentTime()
            print(current_time.toString("hh:mm:ss"))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    pymon = Pymon()
    pymon.run()