if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)
    
import pandas as pd
# import pyodbc
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
from config import *

class db_hr(): #讀取excel 單一零件
    def __init__(self):
        # self.cn = pyodbc.connect(config_conn_HR) # connect str 連接字串
        # self.rpt = pyodbc.connect(config_conn_RPT) # connect str 連接字串
        self.cn = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_HR})).connect()
        self.rpt = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_RPT})).connect()
        
    def ps18Getps01(self, ps18): #使用者電腦名稱 查詢 使用者代號
        s = "SELECT TOP 1 ps01 FROM rec_ps WHERE ps18 = '{0}'"
        s = s.format(ps18)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['ps01'] if len(df.index) > 0 else ''

    def pc02Getpc01(self, pc02): #設備 電腦名稱 查詢 設備代號
        s = "SELECT TOP 1 pc01 FROM rec_pc WHERE pc02 = '{0}'"
        s = s.format(pc02)
        df = pd.read_sql(s, self.rpt) #轉pd
        return df.iloc[0]['pc01'] if len(df.index) > 0 else ''

    def cpGer_qs_lis(self, qs01): # qs01使用者代號或設備代號 查詢權限 list 
        s = "SELECT qs02 FROM rec_qs WHERE qs01 ='{0}'"
        s = s.format(qs01)
        df = pd.read_sql(s, self.rpt) #轉pd
        return df['qs02'].tolist() if len(df.index) > 0 else []

def test1():
    # new id
    hr = db_hr()
    lis = hr.cpGer_qs_lis(32)
    print(lis)

if __name__ == '__main__':
    test1()        
    print('ok')