if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import tool_db_yst

class BOM(): # 產生bom to_df() 方法
    def __init__(self, pdno, pump_lock = False):
        self.pdno = pdno #品號
        self.pump_lock = pump_lock #泵浦鎖定 True:不展階  False:展階
        self.db = tool_db_yst.db_yst() # db
        self.bom_init()  # 初始BOM  self.df_bom
        self.ci = 0 # 避免迴圈過多
        self.generator() # 迴圈取得BOM架構
        self.comp_1()    # 後處理計算
    def to_df(self):
        return self.df_bom

    def get_columns(self):
        return{
            'gid':           '件號id',
            'pdno':          '品號',
            'pdorder':       '序號',
            'parent_pdno':   '父件品號',
            'bom_mol':       '用量',
            'bom_den':       '底數',
            'bom_level':     '階數',
            'bom_extend':    '展階', # True需要展階 False不需要展階
            'pd_name':       '品名',
            'pd_spec':       '規格',
            'pd_type':       '品號屬性',
            'supply':        '主供應商',
            'last_price':    '最新進價(本國幣別NTD)',
            'pid':           '父件id',
            'child_quantity':'子件數量(筆數)',
            'root_quantity': '總用量',
            'self_quantity': '自身總用量',
            'stant_mkpdno':  '標準途程品號',
            'stant_mkno':    '標準途程代號',
            'cost_material': '材料費',
            'cost_make':     '加工費',
            'cost_price':    '單價',
            'cost_amount':   '金額',
            'bom_type':      'bom架構屬性', #1正常件 2材料件 3虛設件
            'sales_price_1': '售價定價一',
            'bom_level_lowest': '最底階', # True False
        }

    def get_pdbom_df(self, pdno, bom_level):
        # 單一品號取得bom為dataframe
        # 查詢df 加入df
        df = self.db.get_bom(pdno)
        if df is None:
            return
        # df['MD001'] = df['MD001'].str.strip() # 去除 MD001 頭尾空白
        new_columns = {
            'MD001': 'parent_pdno', # 父件品號
            'MD002': 'pdorder', # 序號
            'MD003': 'pdno', # 品號
            'MD006': 'bom_mol', #'用量'
            'MD007': 'bom_den', #'底數',
            'MB002': 'pd_name', #'品名'
            'MB003': 'pd_spec', #'規格',
            'MB025': 'pd_type', #'品號屬性',
            'MB032': 'supply', #'主供應商',
            'MB050': 'last_price', #'最新進價(本國幣別NTD)',
            'MB010': 'stant_mkpdno', #'標準途程品號',
            'MB011': 'stant_mkno', #'標準途程代號',
            'MB053': 'sales_price_1', #'售價定價一',
            }
        df = df.rename(columns=new_columns)

        # 以預設值加入df
        columns_value = {
            'gid':           0, # 件號id
            'bom_level':     bom_level + 1, #'階數  為本階+1',
            'bom_extend':    True, #'展階' 預設尚未展階 故為True需要展階
            'pid':           0, #'父件id',
            'child_quantity':0, #'子件數量(筆數)',
            'self_quantity': 0, #'自身總用量',
            'root_quantity': 0, #'總用量',
            'cost_material': 0, #'材料費',
            'cost_make':     0, #'加工費',
            'cost_price':    0, #'單價',
            'cost_amount':   0, #'金額',
            'bom_type':      1, #'bom架構屬性', #1正常件 2材料件 3虛設件
            'bom_level_lowest': False, #'最底階' True False
            }

        for k, v in columns_value.items():
            df.insert(len(df.columns), k, [v]*len(df.index), True) #插在最後

        # pump lock 不展bom功能
        if self.pump_lock == True:
            # 找出吻合的pump 設定為不展階
            df_w = df.loc[(df['bom_level']>0) & (df['pdno'].str.contains(r'^6AA.*'))]
            if len(df_w.index) > 0:
                for i, r in df_w.iterrows():
                    df.at[i,'bom_extend'] = False # 不展BOM
                    df.at[i,'pd_type'] = 'Q' # 改為Q 不展BOM銷售件
        return df

    def bom_init(self): # 初始BOM
        dic = self.db.get_pd_one_to_dic(self.pdno)
        self.df_bom = pd.DataFrame(None, columns = list(self.get_columns().keys())) # bom data
        data_row = {
            'gid':           0, # 件號id
            'pdno':          self.pdno, # 品號
            'pdorder':       '', # 序號
            'parent_pdno':   '', # 父件品號
            'bom_mol':       1, #'用量''
            'bom_den':       1, #'底數',
            'bom_level':     0, #'階數',
            'bom_extend':    True, #'展階' True需要展階 False不需要展階
            'pd_name':       dic['MB002'], #'品名',
            'pd_spec':       dic['MB003'], #'規格',
            'pd_type':       dic['MB025'], #'品號屬性',
            'supply':        dic['MB032'], #'主供應商',
            'last_price':    dic['MB050'], #'最新進價(本國幣別NTD)',
            'pid':           0, #'父件id',
            'child_quantity':0, #'子件數量',
            'root_quantity': 1, #'總用量',
            'stant_mkpdno':  dic['MB010'], #'標準途程品號',
            'stant_mkno':    dic['MB011'], #'標準途程代號',
            'cost_material': 0, #'材料費',
            'cost_make':     0, #'加工費',
            'cost_price':    0, #'單價',
            'cost_amount':   0, #'金額',
            'bom_type':      1, #'bom架構屬性', #1正常件 2材料件 3虛設件
            'sales_price_1': dic['MB053'], #'售價定價一',
            'bom_level_lowest': False, #'最底階' True False
            }
        self.df_bom = self.df_bom.append(data_row, ignore_index=True) #新增首筆

    def generator(self): #建立BOM架構
        while any(self.df_bom['bom_extend'].tolist()): # 需要展泵
            if self.ci == 2000:
                print('迴圈過多')
                break
            self.circul_bom()
            self.ci += 1

    def circul_bom(self): #迴圈取得BOM架構
        df = self.df_bom.copy() # 複製來遍歷
        for i, r in df.iterrows():
            if r['bom_extend'] == True: # 需要展階
                self.df_bom.at[i,'bom_extend'] = False # 不再展BOM
                df_extend = self.get_pdbom_df(r['pdno'], r['bom_level']) # 展階
                if df_extend is None:
                    return # 無bom
                else:
                    # 插入在a, b, 中間
                    df_a = self.df_bom.loc[:i] # 分解 df 為2部分之上部 a
                    df_b = self.df_bom.loc[i+1:] # 分解 df 為2部分之下部 b
                    df_new = df_a.append(df_extend, ignore_index = True)
                    df_new = df_new.append(df_b, ignore_index = True)
                    self.df_bom = df_new #覆蓋 更新
                    return

    def comp_1(self): # 後處理計算
        # gid
        self.df_bom['gid'] = [e for e in range(1, len(self.df_bom.index)+1)]

        # pid
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            if r['bom_level'] > 0: # 非0階
                ri = i           # 當前行數
                while True:
                    if df.iloc[ri-1]['pdno'] == r['parent_pdno'] and df.iloc[ri-1]['bom_level'] == r['bom_level']-1:
                        self.df_bom.at[i,'pid'] = df.iloc[ri-1]['gid']
                        break
                    ri -= 1 # 往前一筆找

        # child_quantity 子件數量(筆數)
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            df_w = df.loc[df['pid'] == r['gid']] # 篩選
            self.df_bom.at[i,'child_quantity'] = 0 if df_w is None else len(df_w.index)

        # self_quantity 自身總用量 (從單一零件往上階計算 用量)
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            total = r['bom_mol'] / r['bom_den']
            upper_id = r['pid'] # 上階的gid 
            while True:
                if upper_id == 0: # 無上階
                    break
                else:
                    df_w = df.loc[df['gid'] == upper_id] # 尋找父件
                    total *=  df_w.iloc[0]['bom_mol'] / df_w.iloc[0]['bom_den']
                    upper_id = df_w.iloc[0]['pid']

            self.df_bom.at[i,'self_quantity'] = total # 自身總用量

        # root_quantity 總用量 (從0階往下找 相同品號的總用量)
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            df_w = df.loc[df['pdno'] == r['pdno']] # 尋找所有 品號相同者 的 總用量 加總
            self.df_bom.at[i,'root_quantity'] = df_w['self_quantity'].sum()

        # bom_type bom架構屬性 1正常件(加工) 2材料件 3虛設件
        for i, r in df.iterrows():
            if r['pdno'][0:1] == '2':  #產品為2開頭  判定為材料
                self.df_bom.at[i,'bom_type'] = 2; continue

            if r['pd_type'] == 'Y': # 判定為虛設件
                self.df_bom.at[i,'bom_type'] = 3; continue

            # 市購件買來加工，代料加工件買來後加工  屬於此類
            if all([r['pd_type']=='P', r['pid']>0]): #P件,有父階
                df_w = df.loc[df['gid'] == r['pid']] # 父件
                if df_w.iloc[0]['child_quantity'] == 1: # 父件的子件數量為1 為代料件  故判定為材料
                    self.df_bom.at[i,'bom_type'] = 2; continue

            if r['pd_type'] == 'P': # P件 判定為材料(接近材料屬性 非純粹的材料)
                self.df_bom.at[i,'bom_type'] = 2; continue

            # else
            self.df_bom.at[i,'bom_type'] = 1  # 預設為正常加工

        # last_price 最新進價(本國幣別NTD)
        # supply 主供應商 (供應商代號)
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            if r['bom_type'] != 2:
                self.df_bom.at[i,'last_price'] = 0  # 非採購件不抓取最新進價 故為0
                self.df_bom.at[i,'supply'] = ''     # 非採購件不抓取主供應商 故為''

        # bom_level_lowest 最底階 True False
        df = self.df_bom.copy()
        for i, r in df.iterrows():
            if i == len(df.index)-1:
                self.df_bom.at[i,'bom_level_lowest'] = True  # 最尾必為最下階
                # arr_bottom.append(r['gid']) # 最尾必為最下階
            else:
                if r['bom_level'] >= df.iloc[i+1]['bom_level']:
                    self.df_bom.at[i,'bom_level_lowest'] = True  # 本階層 大於等於 下一筆的階層 必為最下階
                    # arr_bottom.append(r['gid']) # 本階層 大於等於 下一筆的階層 必為最下階

def test1():
    # bom = BOM('4A306001')
    bom = BOM('5A010100005')
    # bom = BOM('6AA03FA001EL1A01')
    # bom = BOM('7AA01001A01', pump_lock = True)
    df = bom.to_df()
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None) #顯示最多欄位
    df1 = df[['gid','pdno', 'pid','bom_level','bom_extend','pd_type','sales_price_1']]
    print(df1)




if __name__ == '__main__':
    test1()        