if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

# import re
import pandas as pd
# import pyodbc
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
from config import *

class db_yst(): #讀取excel 單一零件
    def __init__(self):
        # self.cn = pyodbc.connect(config_conn_YST) # connect str 連接字串\
        self.cn = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_YST})).connect()
    
    def get_pd_test(self, pdno):
        s = "SELECT MB001,MB002,MB003 FROM INVMB WHERE MB001 = '{0}'"
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def is_exist_pd(self, pdno): # 品號是否存在
        s = "SELECT MB001 FROM INVMB WHERE MB001 = '{0}'"
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return (len(df.index)>0)

    def get_pur_ma002(self, ma001): # 供應商簡稱
        s = "SELECT MA002 FROM PURMA WHERE MA001 = '{0}'"
        s = s.format(ma001)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['MA002'] if len(df.index) > 0 else ''

    def get_purluck_to_list(self): # 不准交易供應商
        s = "SELECT MA001,MA002 FROM PURMA WHERE MA016 = 3"
        df = pd.read_sql(s, self.cn) #轉pd
        df[['MA001']] = df[['MA001']].apply(lambda e: e.str.strip())
        return df['MA001'].tolist() if len(df.index) > 0 else []
        # lis = df['MA001'].tolist()
        # lis.append('1020001')
        # lis.append('1020010')
        # return lis

    def get_pd_one_to_dic(self, pdno):
        s = """
        SELECT TOP 1
        RTRIM(MB001) AS MB001,MB002,MB003,MB004,MB025,MB032,MB050,MB010,MB011,MB053,MB155,MB156
        FROM INVMB
        WHERE MB001 = '{0}'
        """
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0].to_dict() if len(df.index) > 0 else None

    def get_bom(self, pdno):
        s = """
            SELECT RTRIM(BOMMD.MD001) AS MD001,
                BOMMD.MD002,
                BOMMD.MD003,
                BOMMD.MD006,
                BOMMD.MD007,
                INVMB.MB002,
                INVMB.MB003,
                INVMB.MB004,
                INVMB.MB025,
                INVMB.MB032,
                INVMB.MB050,
                INVMB.MB010,
                INVMB.MB011,
                INVMB.MB053,
                INVMB.MB155,
                INVMB.MB156
            FROM BOMMD
                LEFT JOIN INVMB ON BOMMD.MD003 = INVMB.MB001
            WHERE MD001 = '{0}'
            ORDER BY MD002
            """
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def get_bmk(self, mf001, mf002):
        s = """
            SELECT RTRIM(MF001) AS MF001,MF002,MF003,MF004,MF005,MF006,MF007,MF008,MF022,MF019,MF009,MF024,MF017,MF018,MF023,
            MW002
            FROM BOMMF
                LEFT JOIN CMSMW ON BOMMF.MF004 = CMSMW.MW001
            WHERE MF001 = '{0}' AND MF002 = '{1}'
            ORDER BY MF003
            """
        s = s.format(mf001, mf002)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def stmk_to_df(self):
        # 標準廠商途程單價
        # 例如 士中 熱處理 固定單價 30
        # 人員可在系統維護
        # MF004 製程代號
        # MF006 廠商代號
        # MF018 未稅加工單價
        df = self.get_bmk('92N001','01')
        df = df[['MF004', 'MF006', 'MF018']]
        return df if len(df.index) > 0 else None

    def strk_to_df(self):
        # 標準途程加工單位
        # 例如 熱處理 固定加工單位為KG
        # 人員可在系統維護
        # MF004 製程代號
        # MF017 加工單位
        df = self.get_bmk('92N001','02')
        df = df[['MF004', 'MF017']]
        return df if len(df.index) > 0 else None

    def stpk_to_df(self):
        # 標準廠商採購加工單位
        # 例如 達輝 固定加工單位為KG
        # 人員可在系統維護
        # MF006 廠商代號
        # MF017 加工單位
        df = self.get_bmk('92N001','03')
        df = df[['MF006', 'MF017']]
        return df if len(df.index) > 0 else None

    def wget_imd(self, pdno_arr=''):
        # 品號單位換算
        # pdno_arr 品號 文字陣列
        if pdno_arr == "":
            return None
        pdno_arr = str(pdno_arr).replace(' ','') # 去除空格
        pdno_inSTR = "('" + "','".join(pdno_arr.split(',')) + "')"
        s = """
            SELECT RTRIM(MD001) AS MD001,RTRIM(MD002) AS MD002,MD003,MD004 FROM INVMD
            WHERE MD001 IN {0}
            ORDER BY MD001, MD002 ASC
            """
        s = s.format(pdno_inSTR)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def wget_cti(self, pdno_arr):
        # 品號製令進貨 最新進貨
        # 以 製程代號(TI015)、 廠商代號(TH005) 品號(TI004) 為群組尋找最後一筆
        # pdno_arr 品號 文字陣列
        if pdno_arr == "":
            return None
        pdno_arr = str(pdno_arr).replace(' ','') # 去除空格
        pdno_inSTR = "('" + "','".join(pdno_arr.split(',')) + "')"
        s = """
            SELECT TI015,TH005,TI004,TI024,TI023,TI001,TI002
            FROM MOCTI
                LEFT JOIN MOCTH ON TI001=TH001 AND TI002=TH002
            WHERE
                (TI002+'-'+TI001) IN (
                    SELECT MAX(TI002+'-'+TI001)
                    FROM MOCTI LEFT JOIN MOCTH ON TI001=TH001 AND TI002=TH002
                    WHERE TI004 IN {0}
                    GROUP BY TI004,TI015,TH005
                    )
            """
        s = s.format(pdno_inSTR)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    # def get_cti_last(self, pdno, ti015):
    #     # 品號製令進貨 最新進貨 3筆
    #     # 以 品號(TI004) 製程代號(TI015) 尋找最後x筆
    #     s = """
    #         SELECT TOP 5 TI015,TH005,TI004,TI024,TI023,TI001,TI002
    #         FROM MOCTI
    #             LEFT JOIN MOCTH ON TI001=TH001 AND TI002=TH002
    #         WHERE
    #             TI004 = {0} AND 
    #             TI015 = {1}
    #         """
    #     s = s.format(pdno)
    #     df = pd.read_sql(s, self.cn)
    #     return df if len(df.index) > 0 else None

    def wget_pui(self, pdno_arr):
        # 品號採購進貨 最新進貨
        # 以 品號(TH004) 廠商代號(TG005)為群組尋找最後一筆
        # pdno_arr 品號 文字陣列
        if pdno_arr == "":
            return None
        pdno_arr = str(pdno_arr).replace(' ','') # 去除空格
        pdno_inSTR = "('" + "','".join(pdno_arr.split(',')) + "')"
        s = """
            SELECT TH004,TG005,TG007,TG008,TH018,TH019,TH016,TH056,TH007,TH008,TH001,TH002
            FROM PURTH
                LEFT JOIN PURTG ON TH001=TG001 AND TH002=TG002
            WHERE
                (TH002+'-'+TH001) IN (
                    SELECT MAX(TH002+'-'+TH001)
                    FROM PURTH LEFT JOIN PURTG ON TH001=TG001 AND TH002=TG002
                    WHERE TH004 IN {0}
                    GROUP BY TH004,TG005
                    )
            """
        s = s.format(pdno_inSTR)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def get_puilast_to_df(self, pdno, last_time):
        # 品號採購進貨 最後新進貨N筆 的供應商清單  (進貨 非 採購)
        # pdno: 品號(TH004) 
        # last_time :進貨N筆
        # TG005       MA002
        # 供應商代號  簡稱
        s = """
            SELECT TOP {1} TG005,MA002
            FROM PURTH
                LEFT JOIN PURTG ON TH001=TG001 AND TH002=TG002
                LEFT JOIN PURMA ON TG005=MA001
            WHERE
                TH004 = '{0}'
            ORDER BY TH002 DESC
            """
        # SELECT TOP {1} TH004,TG005,TG007,TH018,TH008,TH001,TH002
        s = s.format(pdno, last_time)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def get_cma_to_dic(self, pdno, ma002, ma003, ma012):
        # 取最新一筆加工計價資料 (定位為事先核定)
        # 相對進貨單價(事後確定)
        # MA001,  MA002, MA003,  MA004,   MA005,  MA012
        # 品號 製程代號 廠商代號 計價單位 加工單價   生效日
        s = """
        SELECT TOP 1 
            -- RTRIM(MA001) AS MA001,MA002,MA003,RTRIM(MA004) AS MA004,MA005,MA012
            RTRIM(MA004) AS MA004,MA005
        FROM MOCMA
        WHERE
            MA001 = '{0}' AND --品號
            MA002 = '{1}' AND --製程代號
            MA003 = '{2}' AND --廠商代號
            MA012 <= '{3}'  --已生效
        ORDER BY MA012 DESC 
        """
        s = s.format(pdno, ma002, ma003, ma012)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0].to_dict() if len(df.index) > 0 else None

    def get_pmb_to_dic(self, pdno, mb002, mb014):
        # 取最新一筆採購計價資料 (定位為事先核定)
        # 相對進貨單價(事後確定)
        # MB001,  MB002, MB003, MB004, MB011,  MB014 
        # 品號  廠商代號 幣別  計價單位 採購單價 生效日
        s = """
        SELECT TOP 1 
            RTRIM(MB004) AS MB004,MB011
        FROM PURMB
        WHERE
            MB001 = '{0}' AND --品號
            MB002 = '{1}' AND --廠商代號
            MB014 <= '{2}'  --已生效
        ORDER BY MB014 DESC 
        """
        s = s.format(pdno, mb002, mb014)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0].to_dict() if len(df.index) > 0 else None

    # def test_df(self):
    #     s = """
    #     SELECT
    #         TA003,TA004,RTRIM(MW002) AS MW002 ,RTRIM(TA006) AS TA006,TA007,TA024,
    #         SUBSTRING(TA008,1,4) +'/'+ SUBSTRING(TA008,5,2) +'/'+ SUBSTRING(TA008,7,2) AS TA008,
    #         SUBSTRING(TA009,1,4) +'/'+ SUBSTRING(TA009,5,2) +'/'+ SUBSTRING(TA009,7,2) AS TA009,
    #         TA010,TA011,TA012,TA034
    #     FROM SFCTA
    #         LEFT JOIN CMSMW ON SFCTA.TA004 = CMSMW.MW001

    #     WHERE 
    #         TA001 = '5101' AND
    #         TA002 = '20220418001'
    #     ORDER BY TA003
    #     """
    #     s = s.format('4DD0020085')
    #     df = pd.read_sql(s, self.cn) #轉pd
    #     return df if len(df.index) > 0 else None

    # def test_df(self):
    #     s = """
    #     (SELECT MD001, MD002 FROM CMSMD)
    #     UNION
    #     (SELECT MA001, MA002 FROM PURMA)
    #     """
    #     df = pd.read_sql(s, self.cn) #轉pd
    #     return df if len(df.index) > 0 else None        

    # def test_df(self):
    #     s = """
    #     SELECT 
    #         TC001,TC002,TC003,TC047,TC048,TC049,
    #         TC023, md1.MD002 AS 移出部門,TC006 AS 移出工序,TC007 AS 移出製程,RTRIM(c1.MW002) AS cw1,
    #         TC041, md2.MD002 AS 移入部門,TC008,TC009,RTRIM(c2.MW002) AS cw2,
    #         ta.TA024,
    #         CASE
    #             WHEN TC013 = 1 THEN '1.正常完成'
    #             WHEN TC013 = 2 THEN '2.重工完成'
    #             WHEN TC013 = 3 THEN '3.退回重工'
    #             WHEN TC013 = 4 THEN '4.撥轉'
    #             WHEN TC013 = 5 THEN '5.盤盈損'
    #             WHEN TC013 = 6 THEN '6.投入'
    #             ELSE ''
    #         END AS TC013,
    #         TC010,TC016,TC036,TC014,TC037
    #     FROM SFCTC as tc
    #         LEFT JOIN CMSMW as c1 ON tc.TC007 = c1.MW001
    #         LEFT JOIN CMSMW as c2 ON tc.TC009 = c2.MW001
    #         LEFT JOIN SFCTA as ta ON tc.TC004 = ta.TA001 AND tc.TC005 = ta.TA002 AND tc.TC008=ta.TA003
    #         LEFT JOIN (
    #         (SELECT MD001, MD002 FROM CMSMD) UNION (SELECT MA001, MA002 FROM PURMA)
    #         ) as md1 ON tc.TC023 = md1.MD001
    #         LEFT JOIN (
    #         (SELECT MD001, MD002 FROM CMSMD) UNION (SELECT MA001, MA002 FROM PURMA)
    #         ) as md2 ON tc.TC041 = md2.MD001

    #     WHERE 
    #         TC004 = '5101' AND
    #         TC005 = '20220418001'
    #     ORDER BY TC006,TC002,TC003
    #     """
    #     s = s.format('4DD0020085')
    #     df = pd.read_sql(s, self.cn) #轉pd
    #     return df if len(df.index) > 0 else None

    def test_df(self):
        s = """
            SELECT TOP 100
                TA001 AS 製令單別,
                TA002 AS 製令單號,
                TA003 AS 開單日期,
                RTRIM(TA006) AS 品號,
                TA034 AS 品名,
                TA013 AS 確認碼,
                CASE WHEN TA011 = '1' THEN '1.未生產' WHEN TA011 = '2' THEN '2.已發料' WHEN TA011 = '3' THEN '3.生產中' WHEN TA011 LIKE 'Y' THEN 'Y.已完工' WHEN TA011 LIKE 'y' THEN 'y.指定完工' ELSE '' END AS 狀態碼
            FROM MOCTA
            WHERE
                -- TA001 LIKE '5101' AND
                -- TA002 LIKE '20230303001'
                RTRIM(TA006) LIKE '4A401009' AND
                TA013 LIKE 'Y'
            ORDER BY TA003 DESC
        """
        df = pd.read_sql(s, self.cn) #轉pd
        print(df.iloc[0]['品號'])
        print(len(df.iloc[0]['品號']))
        return df if len(df.index) > 0 else None
                # TA001 LIKE '5101' AND
                # TA002 LIKE '20230303001'
                # TA006 LIKE '4A401009' AND
                # TA013 LIKE 'Y'
def test1():
    db = db_yst()
    df = db.test_df()
    # df = db.get_cma_to_dic('4N0000308','S073','1020025','20221010')
    # df = db.get_pmb_to_dic('3AAB1A3205','1030198','20221014')
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None) #顯示最多欄位
    print(df)



if __name__ == '__main__':
    test1()        