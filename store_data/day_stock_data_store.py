from sqlalchemy import create_engine
import pandas as pd
import win32com.client
import time
import mariadb
import pymysql
pymysql.install_as_MySQLdb()

class Day_stock_data_store:
    def __init__(self):
        self.objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        self.instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") 
        self.objStockChart = win32com.client.Dispatch('CpSysDib.StockChart')
        self.conn = mariadb.connect(
            user='root',
            password='1234',
            database='mystock',
            host='localhost',
            port=3306
        )
        self.cs = self.conn.cursor()
        self.engine = create_engine("mysql://{user}:{pw}@localhost/{db}".format(user='root', pw='1234', db='mystock'))

    def check_connection(self):
        bConnect = self.objCpStatus.IsConnect
        if bConnect == False:
            print('연결 실패')
            exit()
        else:
            print('연결 완료')
        return True
    
    def checkRemainTime(self):
        # 연속 요청 가능 여부 체크
        remainTime = self.objCpStatus.LimitRequestRemainTime / 1000.
        remainCount = self.objCpStatus.GetLimitRemainCount(1)  # 시세 제한
        print("남은시간: ",remainTime, " 남은개수: ", remainCount)
        if remainCount <= 0:
            print("15초당 60건으로 제한합니다.")
            time.sleep(remainTime)

    def write2mariadb(self, stock_code, stock_data_tuples):
        # mariadb에 dataframe 입력
        sql_create_table = 'CREATE TABLE IF NOT EXISTS ' + ''.join(stock_code)+'(date varchar(20), start varchar(20), high varchar(20), low varchar(20), end varchar(20), volume varchar(20), transaction_price varchar(20))'
        self.cs.execute(sql_create_table)

        sql_insert_value = 'INSERT INTO ' + ''.join(stock_code)+'( date, start, high, low, end, volume, transaction_price) VALUES ( ?, ?, ?, ?, ?, ?, ?)'
        self.cs.executemany(sql_insert_value, stock_data_tuples)
        self.conn.commit()

    def read_stockNameList(self):
        # 대신증권 api를 사용해 종목코드, 종목명, kospi or kosdaq 정보를 읽음
        self.check_connection()

        CPE_MARKET_KIND = {'KOSPI':1, 'KOSDAQ':2} 
        rows = []

        for key, value in CPE_MARKET_KIND.items(): 
            codeList = self.instCpCodeMgr.GetStockListByMarket(value) 
            for code in codeList: 
                name = self.instCpCodeMgr.CodeToName(code) 
                sectionKind = self.instCpCodeMgr.GetStockSectionKind(code)
                row = [code, name, key, sectionKind]
                rows.append(row)

        print('모든 종목을 불러왔습니다. ')

        stockitems = pd.DataFrame(rows, columns=['code', 'name', 'section','sectionKind'])
        stockitems.loc[stockitems['sectionKind'] == 10, 'section'] = 'ETF'
        stockitems.to_csv('stockitems.csv', index = False)

        no_etf_data = stockitems.loc[stockitems['sectionKind'] != 10]
        no_etf_data = no_etf_data.loc[no_etf_data['sectionKind'] != 17]
        no_etf_data.to_csv('no_etf_data.csv', index=False)

        print('파일 저장 완료')
        return no_etf_data

    def read_stockData(self):
        stock = []

        while True:
            self.checkRemainTime()

            self.objStockChart.BlockRequest()                
            len = self.objStockChart.GetHeaderValue(3)
            
            for j in range(len):
                row = []
                for i in range(7):
                    row.append(self.objStockChart.GetDataValue(i,j))
                stock.append(row)

            rqStatus = self.objStockChart.Continue
            if rqStatus == False:
                break            
        return stock

    def set_objStockChart(self, stock_code):
        self.objStockChart.SetInputValue(0, stock_code)
        self.objStockChart.SetInputValue(1, ord('1'))
        self.objStockChart.SetInputValue(2, 0)
        self.objStockChart.SetInputValue(3, '19800101')
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9])
        self.objStockChart.SetInputValue(6, ord('D'))
        self.objStockChart.SetInputValue(9, ord('1'))
   
    def store_data(self):
        # kospi, kosdaq 종목코드, 종목명
        stockNameList = self.read_stockNameList()

        # 
        for stock_code in stockNameList['code']:
            stockData = self.read_stockData()
            self.set_objStockChart(stock_code)             

            stock_data = pd.DataFrame(stockData, columns = ['date', 'start', 'high', 'low', 'end', 'volume', 'transaction_price'])
            stock_data_tuples = list(stock_data.itertuples(index=False, name=None))

            print(stock_code)
            self.write2mariadb(stock_code, stock_data_tuples)     

if __name__ == "__main__":
    store_obj = Day_stock_data_store()
    store_obj.store_data()