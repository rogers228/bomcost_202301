if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import tool_db_yst
import tool_func

class BMK(): #　產品製程資料 (僅鼎新資料)
    def __init__(self, mkpdno, mkno):
        # self.bomrow_dic = bomrow_dic # 一筆產品結構tool_bom的 dic資料(包含所有欄位)
        self.mkpdno = mkpdno # 標準途程品號
        self.mkno = mkno     # 標準途程代號
        self.db = tool_db_yst.db_yst() # db
        self.bmk_init()  # 初始BOM  self.df_bom
        self.wash_data_2() # 添加 MF009 MF024 時間格式化的欄位

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
            'MF009': '固定人時(秒)',
            'F_MF009': '固定人時(時分秒)',
            'MF024': '固定機時(秒)',
            'F_MF024': '固定機時(時分秒)',
            'MF017': '加工單位',
            'MF018': '加工單價',
            'MF023': '備註(自訂單價)',
            'MW002': '製程'
        }
            # 以下由 tool_cost 後製
            # 'SS001': '製程單價',  # 該製程加工單位的單價  例如1KG多少錢 (非鼎新資料)
            # 'SS002': '單價',      # 該製程換算為PCS的單價 (非鼎新資料)   
    def bmk_init(self):
        df = pd.DataFrame(None, columns = list(self.get_columns().keys())) # None dataframe
        df_m = self.db.get_bmk(self.mkpdno, self.mkno)
        if df_m is not None:
            # 有產品製程資料
            df = df.append(df_m, ignore_index = True)
        self.df_bmk = df

    def wash_data_2(self): # 添加 MF009 MF024 時間格式化的欄位
        df = self.df_bmk.copy()
        for i, r in df.iterrows():
            self.df_bmk.at[i,'F_MF009'] = tool_func.fmt_seconds(r['MF009'])
            self.df_bmk.at[i,'F_MF024'] = tool_func.fmt_seconds(r['MF024'])

    def to_df(self):
        return self.df_bmk

def test1():
    mk = BMK('4B101050','01')
    df = mk.to_df()
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None)       # 顯示最多欄位
    print(df)


if __name__ == '__main__':
    test1()        