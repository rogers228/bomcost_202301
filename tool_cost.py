if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import time
import pandas as pd
import re
import array
from math import ceil
import tool_db_yst
import tool_bom, tool_bmk

class COST(): # 基於bom 與 製程bmk 合併產生出 cost data
    def __init__(self, pdno, pump_lock = False):
        #tool
        self.yst = tool_db_yst.db_yst()
        self.df_bom = tool_bom.BOM(pdno, pump_lock).to_df() # BOM
        pdno_arr_str = ','.join(list(set(self.df_bom['pdno'].tolist())))
        # print('pdno_arr_str:', pdno_arr_str)
        self.df_md = self.yst.wget_imd(pdno_arr_str) #所有品號單位換算
        # print(self.df_md)

        self.df_ct = self.yst.wget_cti(pdno_arr_str) #所有品號 托外最新進貨
        if self.df_ct is not None:
            self.df_ct = self.df_ct.sort_values(by='TI002', ascending=False) # 排序最新

        self.df_pu = self.yst.wget_pui(pdno_arr_str) #所有品號 採購最新進貨
        if self.df_pu is not None:
            self.df_pu = self.df_pu.sort_values(by='TH002', ascending=False) # 以單號排序最新

        # print(self.df_ct)
        # 標準廠商途程單價
        self.df_stmk = self.yst.stmk_to_df() # 標準廠商途程單價
        self.df_strk = self.yst.strk_to_df() # 標準途程加工單位
        self.df_stpk = self.yst.stpk_to_df() # 標準廠商採購加工單位
        self.lis_puk = self.yst.get_purluck_to_list() # 禁止交易供應商
        self.error_information = {} # 異常信息

        self.comp_1() # 計算成本群組 rerutn  self.cost_group_list
        self.comp_4() # 解析產品規格KG
        self.comp_2() # 計算所有製程成本 rerutn  self.dic_bmk, self.dic_gid_pdno
        self.comp_3() # 檢查資料異常

    def error_dic(self):
        return self.error_information
        '''
        {
            'err1': lis_err1, # (gid list) 最下階應為P, Q件，或應再建立P件為子件
            'err2': lis_err2, # (gid list) P件不應該有BOM架構，或有BOM應為S件or M件
            'err3':
            {
                'gid': lis_err3 (加工單位) 有製程單位卻無換算單位
            },
            'err4':  該製程 有禁止交易供應商  未更新
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err5':  該製程 已有最新托外進貨  未更新
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err6': 該製程  不符合 標準廠商加工單價
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err7': 該製程  不符合 標準加工單位
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err8': 該製程(採購)  最後N筆已更改供應商
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err9': 該製程(採購)  加工單位 不符合 第一順位標準廠商採購加工單未  第二順位最後進貨計價單價
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err10': 該製程(採購)  不符合 最新單位進價
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err11': 該製程(採購)  未維護PCS以外的單位換算
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            }
            'err12': 該製程(採購)  單位換算不符合最新進貨
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            },
            'err13': 該製程(採購)  採購件P採購單位與最新計價單位皆為PCS，卻有設定換算單位
            {
                'pdno':{
                    'mk_i':
                    'message'
                }
            },
        }
        '''

    def group_to_list(self):
        return self.cost_group_list

    def dlookup_bmk_pdno(self, pdno): # 產品製程
        return self.dic_bmk[pdno] # dict key: pdno,   value: ddaaframe

    def dlookup_pdmd(self, pdno, column):
        # 從 pdno 找對應的 column欄位名稱的值value (首筆)
        # MD001     MD002  MD003  MD004
        # 4A306025  KG     4.160    1.0
        result = ''
        df = self.df_md
        if df is None:
            return result

        df_w = df.loc[df['MD001']==pdno]
        if len(df_w.index)>0:
            result = df_w.iloc[0][column]
        return result

    def dlookup_stmk(self, mf006, mf004):
        # 從2個條件: 廠商代號mf006, 製程代號mf004
        # 找到 MF018 加工單價
        result = 0
        df = self.df_stmk
        if df is None:
            return result
        df_w = df.loc[(df['MF006']==mf006.ljust(10)) & (df['MF004']==mf004)]
        if len(df_w.index)>0:
            result = df_w.iloc[0]['MF018']
        return result

    def dlookup_strk(self, mf004):
        # 從1個條件: 製程代號mf004
        # 找到 MF017 加工單位
        result = ''
        df = self.df_strk
        if df is None:
            return result
        df_w = df.loc[df['MF004']==mf004]
        if len(df_w.index)>0:
            result = df_w.iloc[0]['MF017']
        return result

    def dlookup_stpk(self, mf006):
        # 從1個條件: 廠商代號mf006
        # 找到 MF017 加工單位
        result = ''
        df = self.df_stpk
        if df is None:
            return result
        df_w = df.loc[df['MF006']==mf006]
        if len(df_w.index)>0:
            result = df_w.iloc[0]['MF017']
        return result

    def dlookup_pdct(self, pdno, ti015, column):
        # 從2個條件: 品號pdno, 製程代號ti015
        # 找到第一筆(最新)對應的 column欄位名稱的值value
        # 廠商代號 單價 計價單位 進貨單別   單號
        # TH005   TI024 TI023 TI001        TI002
        #          25.0    KG  D302  20170607002
        result = ''
        df = self.df_ct
        if df is None:
            return result
        df_w = df.loc[(df['TI004']==pdno) & (df['TI015']==ti015)]
        if len(df_w.index)>0:
            result = df_w.iloc[0][column]
        return result

    # def dlookup_pdpu(self, pdno, column):
    #     # 最新採購進貨
    #     # 從1個條件: 品號pdno  (不分供應商)
    #     # 找到第一筆(最新)對應的 column欄位名稱的值value
    #     # 品號  廠商代號 幣別  原幣單位進價(單價) 計價數量 計價單位 進貨數量 單位  進貨單別 單號
    #     # TH004,TG005,  TG007,TH018,           TH016,  TH056,   TH007,  TH008,TH001,  TH002
    #     result = ''
    #     df = self.df_pu
    #     if df is None:
    #         return result
    #     df_w = df.loc[df['TH004']==pdno]
    #     if len(df_w.index)>0:
    #         result = df_w.iloc[0][column]
    #     return result

    def dlookdic_pdpu(self, pdno):
        # 最新採購進貨
        # 從1個條件: 品號pdno  (不分供應商)
        # 找到第一筆(最新) to dictionary
        # 品號  廠商代號 幣別  原幣單位進價(單價) 計價數量 計價單位 進貨數量 單位  進貨單別 單號
        # TH004,TG005,  TG007,TH018,           TH016,  TH056,   TH007,  TH008,TH001,  TH002
        result = {}
        df = self.df_pu
        if df is None:
            return result
        df_w = df.loc[df['TH004']==pdno]
        if len(df_w.index)>0:
            result = df_w.iloc[0].to_dict()
        return result

    def dlookup_pdmd_tolist(self, pdno):
        # 從 pdno 找所有換算單位
        result = []
        df = self.df_md
        if df is None:
            return result
        df_w = df.loc[df['MD001']==pdno]
        if len(df_w.index)> 0:
            result = df_w['MD002'].tolist()
        return result

    def dlookup_gid(self, gid, column):
        # 從 gid 找對應的 column欄位名稱的值value
        result = None
        df = self.df_bom
        if gid in df['gid'].tolist():
            result = df.loc[df['gid']==gid][column].item()
        return result

    def dlookup_pid(self, pid, column):
        # 從 gid 找對應的 column欄位名稱的值value
        result = None
        df = self.df_bom
        if pid in df['pid'].tolist():
            result = df.loc[df['pid']==pid][column].item()
        return result

    def dlookup_gid_pkg(self, gid):
        # 從品號基本資料中的 規格欄位裡 以re解析出KG
        # 從 gid 找kg的值value
        result = ''
        df = self.df_pkg
        if gid in df['gid'].tolist():
            result = df.loc[df['gid']==gid]['p_kg'].item()
        return result

    def dlookup_pkg(self, pdno):
        # 從品號基本資料中的 規格欄位裡 以re解析出KG
        # 從品號pdno 找kg的值value
        result = ''
        df = self.df_pkg
        if df is None:
            return result
        df_w = df.loc[df['pdno']==pdno]
        if len(df_w.index)>0:
            result = df_w.iloc[0]['p_kg']
        return result

    def isRegMath_float(self,findStr): #返回第一個被找到的數字值
        if any([findStr == None, findStr =='']):
            return ''
        mRegex = re.compile(r'\d+\.?\d*')
        match = mRegex.search(findStr)
        return match.group() if match else ''

    def is_pui_last_change(self, pdno, mf006):
        # 最後N筆交易紀錄 是否已改變供應商 (用來檢查是否修改主供應商)
        # 是否 已異於主供應商  且為同一間
        # mf006 目前廠商代號
        result, message = False, ''
        last_time = 3 # 最後3筆交易紀錄
        df_c = self.yst.get_puilast_to_df(pdno, last_time) # 最後n筆交易供應商 dataframe 1010033 同洋
        if df_c is not None:
            lis_c = df_c['TG005'].tolist()
            if len(lis_c) >= last_time: # 有超過N筆交易紀錄
                if lis_c.count(lis_c[0]) == len(lis_c): # N筆交易紀錄為同一間
                    if lis_c[0] != mf006:  # 異於主供應商
                        result = True
                        message =  f"最後3筆進貨紀錄為{df_c.iloc[0]['TG005'].strip()}{df_c.iloc[0]['MA002']}\n建議修改為主供應商"
        return result, message

    def comp_1(self):
        '''
            計算成本群組
            一個成本群組有可能是2個或以上 品號所組成，視為同一個零件成本 才符合成本計算
            例如:軸心與其材料件ㄤ; 零件成品與其子件(代料加工件)  實務直觀上為同一零件
            '''
        gid = self.dlookup_gid # dlookup for gid
        pid = self.dlookup_pid # dlookup for gid

        # step 1 找子件 bom_type 為 2 判定為材料(接近材料屬性 非純粹的材料)
        arr_c = array.array('i') # child 材料子件的gid
        arr_gid = array.array('i', self.df_bom['gid'].tolist()) # parent
        while len(arr_gid)>0:
            curr_id = arr_gid[0]
            arr_gid.remove(curr_id)
            if gid(curr_id,'bom_type') == 3: # 3虛設件
                continue
            if gid(curr_id,'child_quantity')==1: # 僅一個子件
                child_id = pid(curr_id,'gid')
                if pid(curr_id,'bom_type')==2:     # 往下找 該子件為 2材料 (bom_type 由 tool_bom class定義)
                    arr_c.append(child_id) # 添加到子層
                    continue
        # print('step 1 arr_c:', arr_c.tolist())

        # step 2 由材料找
        df = self.df_bom
        df_w = df[df['pdno'].str.contains(r'^2.*')] # 材料件 品號2開頭
        for i, r in df_w.iterrows():
            if not (r['gid'] in arr_c):
                arr_c.append(r['gid'])

        lis_c = arr_c.tolist(); lis_c.sort() # 排序
        arr_c = array.array('i', lis_c)
        # print('step 2 arr_c:', arr_c.tolist())

        # step 3 產生 group list
        # [[90], [89], [88, 87], [87], [86], [85, 84, 83]]
        lis_g = []
        arr_gid = array.array('i', self.df_bom['gid'].tolist()) # all gid

        arr_gid.reverse() # 反序遍歷
        # print('step 3 arr_gid:', arr_gid.tolist())
        while len(arr_gid)>0:
            lis_t = [] # 暫存
            curr_id = arr_gid[0]
            # print('curr_id:', curr_id)
            arr_gid.remove(curr_id) # 刪除本迴圈的id
            loop_id = curr_id

            while loop_id in arr_c: # 是材料
                arr_c.remove(loop_id) # 移除arr_c 直到沒有 才不會無線循環
                if curr_id not in lis_t:
                    lis_t.append(loop_id) # 添加自己

                parent_id = gid(loop_id,'pid')
                p_level =gid(parent_id, 'bom_level') # 父階層
                if all([parent_id not in lis_t]):
                    lis_t.append(parent_id) # 添加父  (有機率會重複添加gid為1者，後續應刪除 移除重複的首階)
                loop_id = parent_id
            else:
                if curr_id not in lis_t:
                    lis_t.append(curr_id)
            # print('lis_t:', lis_t)
            lis_g.append(lis_t)
        # print('lis_g:', lis_g)

        lis_g.reverse() # 恢復反序
        for lis in lis_g:
            if len(lis) > 1:
                lis.reverse()
        # print('lis_g:', lis_g)

        finish = False # 移除重複元素
        while finish == False:
            for i in range(len(lis_g)-1): # 遍歷  不歷最後一個
                if lis_g[i][0] in lis_g[i+1]:
                    del lis_g[i]
                    finish = False; break
                else:
                    finish = True
            if len(lis_g) <= 1:
                finish = True

        # 移除重複的首階
        # 首個gid 必為1，在第首個項目中必定有gid 為1
        # 固 第2個以後的項目中 肯定不能再有1 否則為重複
        for i in range(1, len(lis_g)): # 第2個開始遍歷
            if 1 in lis_g[i]:
                del lis_g[i][0]

        self.cost_group_list = lis_g

    def comp_2(self): # 計算宏觀產品製程成本 (包含 製程 與 採購)
        today = time.strftime("%Y%m%d", time.localtime()) # 今日日期8碼文字
        info = self.error_information # 異常信息
        df = self.df_bom
        # pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
        # pd.set_option('display.max_columns', None) #顯示最多欄位
        # print(df)
        re_f = self.isRegMath_float # re 文字找數字
        pmd = self.dlookup_pdmd     # 找換算單位
        pmd_tolist = self.dlookup_pdmd_tolist # 從 pdno 找所有換算單位
        stmk = self.dlookup_stmk     # 找標準廠商途程單價
        strk = self.dlookup_strk     # 找標準途程加工單位
        stpk = self.dlookup_stpk     # 找標準廠商採購加工單位
        pdct = self.dlookup_pdct     # 找托外最新進價
        pdpu_dic = self.dlookdic_pdpu  # 找採購最新進貨
        pkg = self.dlookup_pkg       # 找出規格裡的KG數
        pmb = self.yst.get_pmb_to_dic # 採購計價資料
        cma = self.yst.get_cma_to_dic # 加工計價資料
        # step 1
        dic_msy = {} #　所有M,S,Y品號的產品製程
        df_w = df[df['pd_type'].isin(['M','S','Y'])]
        for i, r in df_w.iterrows():
            df_bmk = tool_bmk.BMK(r['stant_mkpdno'], r['stant_mkno'])#
            dic_msy[r['pdno']] = df_bmk.to_df()

        # step 2
        dic_all = {} # 所有產品的宏觀產品製程
        dic_columns = {
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
            'MW002': '製程',
            'SS001': '製程單價',  # 該製程加工單位的單價  例如1KG多少錢 (非鼎新資料)
            'SS002': '單價',      # 該製程換算為PCS的單價 (非鼎新資料)
            'SS031': '最新進價',  #
            }
        lis_columns = list(dic_columns.keys())
        for i, r in df.iterrows():
            dic_err3 = {}; lis_err3 = [] #檢查 3 有製程單位卻無換算單位
            # 宏觀的產品製程( 包含產品製程資料 及 採購資料)  to_df() 方法
            df_m = pd.DataFrame(None, columns = list(dic_columns.keys())) # None dataframe
            if r['pd_type']=='P': # 品號屬性 P.採購件
                # print(r['pdno'])
                dic_p = { # 預設值
                'MW002': '採購', # 製程
                'MF006': '', # 供應商代號
                'MF007': '', # 供應商簡稱'
                'MF017': '', # 加工單位 第 1 順位 為 研發採購單位unit 第2順位採購單位 第3順位庫存單位
                'MF018': '', # 採購單價 第1順位 研發採購單價  第2順位 最新進價(本國幣別NTD)
                }
                rd_price, rd_unit, rd_mf006, rd_mf007 = self.yst.get_rdpu(r['pdno']) # 研發採購單價, 研發採購單位, 供應商代號, 供應商簡稱
                # print(rd_price, rd_unit, rd_mf006, rd_mf007)
                if rd_price != 0:
                    # 有設定研發採購單價
                    dic_p['MF006'] = rd_mf006 # 研發採購單價的  供應商代號,
                    dic_p['MF007'] = rd_mf007 # 研發採購單價的  供應商簡稱,
                    dic_p['MF017'] = rd_unit  # 採購加工單位第 1 順位 為 研發採購單位unit
                    dic_p['MF018'] = rd_price # 採購單價 第 1 順位 為研發採購單價
                else:
                    dic_p['MF006'] = r['supply'] # 供應商代號
                    dic_p['MF007'] = self.yst.get_pur_ma002(r['supply']) # 供應商簡稱'
                    dic_p['MF017'] = r['pd_unit'] if r['pur_unit'] == '' else r['pur_unit'] # 採購加工單位第 2 順位 為採購單位 第 3 順位 為庫存單位
                    dic_p['MF018'] = r['last_price'] # 採購單價 第 2 順位 最新進價(本國幣別NTD)

                # p_unit = r['pur_unit'] # 採購加工單位第 2 順位 為採購單位
                # # p_unit = pmd(r['pdno'], 'MD002') # 採購加工單位第一順位 為單位換算
                # if p_unit == '':
                #     p_unit = r['pd_unit'] # 加工單位第 3 順位 為庫存單位
                # dic_p = {
                # 'MW002': '採購', # 製程
                # 'MF006': r['supply'], # 供應商代號
                # 'MF007': self.yst.get_pur_ma002(r['supply']), # 供應商簡稱'
                # 'MF017': p_unit, # 加工單位 第一順位採購單位 第二順位庫存單位
                # 'MF018': r['last_price'] # 最新進價(本國幣別NTD)
                # }

                df_m = df_m.append(dic_p, ignore_index=True) # 20220915
                # pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
                # pd.set_option('display.max_columns', None) #顯示最多欄位
                # print(df_m)

            elif r['pd_type'] in 'MSY': # 品號屬性 M.自製件,S.託外加工件,Y.虛設品號
                df_make = dic_msy[r['pdno']] # 該品號的製程
                df_m = df_m.append(df_make, ignore_index = True) # 20220915

                # 檢查 3 有製程單位卻無換算單位
                lis_pmd = pmd_tolist(r['pdno']) # 該品號所有換算單位
                df_w = df_make.loc[df_make['MF017']!='PCS'] # 非PCS的加工單為
                lis_mk = []
                if len(df_w.index) > 0:
                    lis_mk = list(set(df_w['MF017'].tolist())) # 該品號製程所有加工單位(不含PCS)
                lis_err3 = list(filter(lambda e: e not in lis_pmd, lis_mk))
                if len(lis_err3) > 0:
                    dic_err3[r['gid']] = lis_err3

            elif r['pd_type'] in 'Q': # 不展BOM銷售件
                dic_p = {
                'MW002': '銷售', # 製程
                'MF006': '', # 供應商代號
                'MF007': 'YEOSEH', # 供應商簡稱'
                'MF017': r['sales_unit'],   # 銷售單位
                'MF018': r['sales_price_1'] # 售價定價一
                }
                df_m = df_m.append(dic_p, ignore_index=True)

            dic_all[r['pdno']] = df_m
        info['err3'] = dic_err3

        # step 3 清洗 複製 SS001, SS002, SS003
        ed4={}; ed5={}; ed6={}; ed7={}; ed9={}; ed10={}; ed11={}; ed12={}; ed13={} # error dictionary
        dic_all_new = {}
        for pdno, df_s in dic_all.items():
            df_new = df_s.copy()
            em4={}; em5={}; em6={}; em7={}; em9={}; em10={}; em11={}; em12={}; em13={}  # error message
            for i, r in df_s.iterrows():
                f_ss001 = 0 # 製程單價
                f_ss031 = 0 # 最新進價(鼎新)
                f_ss042 = 0 # 最新進價(自算)
                if r['MW002'] in ['採購','銷售']:
                    f_ss001 = r['MF018'] # 第一順位 最新進價(本國幣別NTD) or 售價一
                    # print('f_ss001:', f_ss001) # 製程單價
                    # 檢查採購 最新單位進價(以計價單位為主的單價)pdno
                    if r['MW002'] in ['採購']:
                        #第二順位 採購計價
                        if f_ss001 == 0:
                            dic_pmb = pmb(pdno,r['MF006'],today)
                            if dic_pmb is not None:
                                if dic_pmb['MB004'] == r['MF017']: # 計價單位與製程單位相符
                                    f_ss001 = dic_pmb['MB011'] # 取採購計價資料

                        # print('pdno11:', pdno)
                        last_pu = pdpu_dic(pdno) # 採購最後進貨
                        f_ss031 = last_pu.get('TH018',0) # 採購最後進貨 原幣單位進價
                        th001 = last_pu.get('TH001','') # 單別
                        th002 = last_pu.get('TH002','') # 單號
                        th007 = last_pu.get('TH007','') # 進貨數量
                        th008 = last_pu.get('TH008',0)  # 原幣單位進價(單價)
                        th016 = last_pu.get('TH016',0) # 計價數量
                        th019 = last_pu.get('TH019',0) # 原幣進貨金額
                        th056 = last_pu.get('TH056',0) # 計價單位
                        tg005 = last_pu.get('TG005','') # 廠商代號
                        tg007 = last_pu.get('TG007','') # 幣別
                        tg008 = last_pu.get('TG008',1) # 匯率

                        try:
                            f_ss042 =round(th016*f_ss031*tg008/th016,3) # 採購最後進貨單價(自算本弊) 等同
                                          #計價數量*原幣單位進價*匯率/計價數量
                        except:
                            f_ss042 = 0 # 沒有進貨

                        # 有進貨就檢查
                        curr_md002 = pmd(pdno,'MD002') # 目前換算單位
                        curr_md003 = pmd(pdno,'MD003') # 目前換算分子 (單位換算)
                        is_pdkg_same = False # 規格kg與換算單位是否相同
                        if all([f_ss042!=0, r['MF017']=='PCS', th056=='PCS']): # 單位與計價單位皆為PCS
                            if all([curr_md002 != '',curr_md002 != 'PCS']): # 卻有換算單位
                                pd_kg = pkg(pdno)
                                if pd_kg == '': # 規格沒有kg數
                                    em13['mk_i']=i; em13['mssage']=f"採購單位與最新計價單位皆為PCS\n卻有單位換算{curr_md002}\n有可能是複製品號時造成，建議刪除"
                                else: # 規格有kg數
                                    if ceil(curr_md003) == ceil(float(pd_kg)):
                                        is_pdkg_same = True
                                        em13 = {} # 規格kg 與 換算kg相同 清除錯誤

                        if f_ss042 != 0: # 有進貨
                            last_md003 = round(th016/th007, 3)  # 最新換算分子 (最新進貨)
                            if all([curr_md003!='', last_md003!='']):
                                if all([ceil(curr_md003)!=ceil(last_md003), not is_pdkg_same]):
                                    em12['mk_i']=i; em12['mssage']=f"未更新單位換算:{last_md003}{th056}"

                        if all([f_ss042!=0, ceil(f_ss042)!=ceil(r['MF018'])]):
                            em10['mk_i']=i
                            em10['mssage'] = f'進貨{th001}-{th002}\n最新進價:{f_ss042}\n計價數量{th016},計價單位{th056}\n原幣進貨金額{th019}{tg007}({tg008})\n進貨數量{th007}{th008}'
                            if r['MF017'] != 'PCS':
                                if curr_md002 == '':
                                    em11['mk_i']=i; em11['mssage']=f"未設定單位換算:{r['MF017']}"

                        # 檢查採購加工單位
                        # last_mf017 = pdpu(pdno, 'TH056') # 採購最後進貨計價單位
                        last_mf017 = last_pu.get('TH056',0) # 採購最後進貨計價單位
                        if last_mf017:
                            if last_mf017 != r['MF017']:
                                em9['mk_i']=i; em9['mssage']=f'不符合採購最後進貨計價單位:{last_mf017}' # 第二順位
                        st_mf017 = stpk(r['MF006']) # 標準廠商採購工單位
                        if st_mf017:
                            if st_mf017 != r['MF017']:
                                em9['mk_i']=i; em9['mssage']=f'不符合標準廠商採購工單位:{st_mf017}' # 第一順位
                else:
                    # 製造加工
                    st_mf017 = strk(r['MF004']) # 標準途程加工單位
                    if st_mf017:
                        if st_mf017 != r['MF017']:
                            em7['mk_i']=i; em7['mssage']=f'不符合標準途程加工單位:{st_mf017}'

                    if r['MF005'] == '1':
                        if re_f(r['MF023']).strip() != '':
                            f_ss001 = re_f(r['MF023']) # 1自製抓 MF023 備註

                    elif r['MF005'] == '2':
                        # 加工單價(第一順位)
                        f_ss001 = r['MF018']       # 2托外抓 MF018

                        # 加工計價資料(第二順位)
                        if f_ss001 == 0:
                            dic_cma = cma(pdno,r['MF004'],r['MF006'],today)
                            if dic_cma is not None:
                                if dic_cma['MA004'] == r['MF017']: # 計價單位與製程單位相符
                                    f_ss001 = dic_cma['MA005'] # 取加工計價資料

                        f_ss031 = pdct(pdno, r['MF004'],'TI024') # 托外最新進價
                        if not f_ss031:
                            f_ss031 = 0

                        st_mf003 = stmk(r['MF006'], r['MF004']) # 標準廠商加工單價
                        if st_mf003: # 有標準廠商加工單價
                            if f_ss001 != st_mf003: # 不符合標準廠商加工單價
                                em6['mk_i']=i; em6['mssage'] = f'不符合標準廠商加工單價{st_mf003}'
                        else: # 沒有標準廠商加工單價
                            if all([float(f_ss031)!=float(f_ss001),
                                    ceil(float(f_ss031))!=float(f_ss001),
                                    f_ss031 != 0]): # 已有最新進價
                                em5['mk_i'] = i
                                ti001 = pdct(pdno, r['MF004'],'TI001')
                                ti002 = pdct(pdno, r['MF004'],'TI002')
                                ti023 = pdct(pdno, r['MF004'],'TI023')
                                th005 = pdct(pdno, r['MF004'],'TH005')
                                em5['mssage'] = f'進貨{ti001}-{ti002}\n已有最新進價:{f_ss031}\n計價單位{ti023}'
                    else:
                        f_ss001 = 0
                        f_ss031 = 0

                if pmd(pdno, 'MD002') == r['MF017']: # 有換算單位
                    f_ss002 = float(f_ss001) * pmd(pdno, 'MD003')
                else:
                    f_ss002 = float(f_ss001) if f_ss001 else 0
                df_new.at[i,'SS001'] = f_ss001 # 製程單價 該製程加工單位的單價  例如1KG多少錢
                df_new.at[i,'SS002'] = f_ss002 # 單價 該製程換算為PCS的單價
                df_new.at[i,'SS031'] = f_ss031 if f_ss031 else 0 # 最新進價

                if r['MF006'].strip() in self.lis_puk: # 供應商代號在禁止交易名單內
                    em4['mk_i']=i; em4['mssage']= f"{r['MF007']}禁止交易"

            #  type object to numeric
            df_new['SS001'] = pd.to_numeric(df_new['SS001'], errors='coerce')
            df_new['SS002'] = pd.to_numeric(df_new['SS002'], errors='coerce')
            df_new['SS031'] = pd.to_numeric(df_new['SS031'], errors='coerce')
            df_new[['SS001','SS002','SS031']] = df_new[['SS001','SS002','SS031']].fillna(value=0) # 空白補0
            # print(df_new.dtypes)
            dic_all_new[pdno] = df_new

            # error
            if len(em4)>0: ed4[pdno] = em4    # 禁止交易供應商 未更新
            if len(em5)>0: ed5[pdno] = em5    # 托外最新加工單價 未維護
            if len(em6)>0: ed6[pdno] = em6    # 不吻合 標準廠商加工單價
            if len(em7)>0: ed7[pdno] = em7    # 不吻合 標準途程加工單位
            if len(em9)>0: ed9[pdno] = em9    # 不吻合 採購加工單位
            if len(em10)>0: ed10[pdno] = em10 # 未更新 採購最新進價
            if len(em11)>0: ed11[pdno] = em11 # 採購件 未設定換算單位
            if len(em12)>0: ed12[pdno] = em12 # 採購件 未更新換算分子
            if len(em13)>0: ed13[pdno] = em13 # 採購件 採購單位與計價單位皆為PCS卻有換算單位

        info['err4']=ed4
        info['err5']=ed5
        info['err6']=ed6
        info['err7']=ed7
        info['err9']=ed9
        info['err10']=ed10
        info['err11']=ed11
        info['err12']=ed12
        info['err13']=ed13
        self.dic_bmk = dic_all_new

        # debug
        # for k, v in dic_all_new.items():
        #     # print(k)
        #     # df = v[['MF001','MW002','MF017','MF018','MF023','SS001','SS002','SS031']]
        #     df = v[['MF001','MW002','MF017','MF006','MF004','SS031']]
        #     print(df)

    def comp_3(self): # 檢查資料異常
        info = self.error_information # 異常信息
        df = self.df_bom

        #檢查 1 最下階應為P, Q件，或應再建立P件為子件
        lis_err1 = []
        df_w = df.loc[(df['bom_level_lowest'] == True) &
            (df['pd_type'] != 'P') & (df['pd_type'] != 'Q')] # P件 且 子件數量大於0
        if len(df_w.index) > 0:
            lis_err1 = df_w['gid'].tolist()
        info['err1'] = lis_err1

        #檢查 2 P件不應該有BOM架構，或有BOM應為S件or M件
        lis_err2 = []
        df_w = df.loc[(df['pd_type'] == 'P') & (df['child_quantity']>0)] # P件 且 子件數量大於0
        if len(df_w.index) > 0:
            lis_err2 = df_w['gid'].tolist()
        info['err2'] = lis_err2

        # 最後N筆交易紀錄 是否已改變供應商
        islast = self.is_pui_last_change
        df_w = df.loc[df['pd_type'] == 'P']
        ed8={} # error dictionary
        if len(df_w.index)>0:
            for pdno, supply in zip(df_w['pdno'].tolist(), df_w['supply'].tolist()):
                em8={} # error message
                result, message = islast(pdno, supply)
                if result == True:
                    em8['mssage']= message
                    ed8[pdno] = em8
        info['err8']=ed8

        # 單位換算不可以為1
    def find_kg(self, pd_spec): # 正則找KG
        mRegex = re.compile(r"(?<=@)\d*\.*\d*(?=KG)")
        match = mRegex.search(pd_spec)
        return match.group() if match else 0

    def comp_4(self): # 解析產品規格KG
        df = self.df_bom
        fkg = self.find_kg
        df_w = df.loc[(df['pd_type'] == 'P') &
                      (df['pdno'].str.contains(r'^2.*')) & # 2開頭
                      (df['pd_spec'].str.contains(r'@\s*.*\s*KG'))] # 有@ KG者

        df1 = df_w[['gid','pdno','pd_spec']].copy()
        df1.reset_index(inplace=True) #重置索引
        lis = []
        for i, r in df1.iterrows():
            lis.append(fkg(r['pd_spec']))
        df1.insert(len(df1.columns), 'p_kg', lis, True) #插在最後
        self.df_pkg = df1

def test1():
    # bom = COST('3AAB1A3205')
    # bom = COST('3B01ABA05150')
    bom = COST('4A602060')
    # print(bom)
    # bom = COST('5A160600033')
    # bom = COST('6AA09N180100004', pump_lock = True)
    # bom = COST('6AE0300002', pump_lock = True)
    # bom = COST('8AC002', pump_lock = True)
    # bom = COST('8CC006', pump_lock = True)

    # dic = bom.dlookdic_pdpu('3B01ABA05150') # 找採購最新進貨
    # print(dic)

    dic_err=bom.error_dic()
    print(dic_err)

    # for pdno, e in dic_err['err13'].items():
    #     print(pdno)
    #     print(e)
    # bom = COST('4A428003')
    # lis = bom.group_to_list()
    # print(lis)

if __name__ == '__main__':
    test1()