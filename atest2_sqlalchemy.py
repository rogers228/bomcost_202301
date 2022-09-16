if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

from sqlalchemy.engine import URL
from sqlalchemy import create_engine
import pandas as pd
from config import *

class db_hr(): #讀取excel 單一零件
    def __init__(self):

        self.cn = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_HR})).connect()

    def ger_ps_df(self): #使用者電腦名稱 查詢 使用者代號
        s = "SELECT Top 10 ps01,ps02,ps03 FROM rec_ps"
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

def test1():
    db = db_hr()
    df = db.ger_ps_df()
    print(df)

if __name__ == '__main__':
    test1()