if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import re
import tool_db_yst

class BMK(): # 產品製程  to_df() 方法
    def __init__(self, mkpdno, mkno):
        self.mkpdno = mkpdno #標準途程品號
        self.mkno =   mkno   #標準途程代號
        self.db = tool_db_yst.db_yst() # db
        self.bmk_init()  # 初始BOM  self.df_bom
        self.wash_data_1() # 清洗 MF023 為數字

    def bmk_init(self):
        self.df_bmk = self.db.get_bmk(self.mkpdno, self.mkno) 

    def wash_data_1(self): # 清洗 MF023 為數字
        df = self.df_bmk.copy()
        for i, r in df.iterrows():
            self.df_bmk.at[i,'MF023'] = self.isRegMath_float(r['MF023'])

    def to_df(self):
        return self.df_bmk

    def get_columns(self):
        return{
            'MF001': '途程品號',
            'MF002': '途程代號',
            'MF003': '順序',
            'MF004': '製程代號',
            'MF005': '自製外包性質',
            'MF006': '供應商代號',
            'MF007': '供應商簡稱',
            'MF008': '製程敘述', 
            'MF022': '檢驗方式代號',
            'MF019': '工時批量',
            'MF009': '固定人時',
            'MF024': '固定機時',
            'MF017': '加工單位',
            'MF018': '加工單價',
            'MF023': '備註(自訂單價)',
            'MD002': '換算單位',
        }

    def isRegMath_float(self,findStr):
        if findStr == None:
            return ''
        mRegex = re.compile(r'\d+\.?\d*')
        match = mRegex.search(findStr)
        return match.group() if match else '' #返回第一個被找到的值

def test1():
    mk = BMK('4A306001', '01')
    df = mk.to_df()
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None)       # 顯示最多欄位
    print(df)


if __name__ == '__main__':
    test1()        