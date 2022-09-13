if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import re
import array
import tool_db_yst
import tool_bom, tool_bmk

class COST(): # 基於bom 與 製程bmk 合併產生出 cost data
    def __init__(self, pdno, pump_lock = False):
        #tool
        self.yst = tool_db_yst.db_yst()
        self.df_bom = tool_bom.BOM(pdno, pump_lock).to_df() # BOM
        self.df_md = self.yst.wget_imd(','.join(list(set(self.df_bom['pdno'].tolist())))) #所有品號單位換算
        self.dic_err = {} # 異常
        self.comp_1() # 計算成本群組 rerutn  self.cost_group_list
        self.comp_2() # 計算所有製程成本 rerutn  self.dic_bmk, self.dic_gid_pdno
        self.comp_3() # 檢查資料異常

    def spec_md_dic(self):
        # 特定製程必為特定單位
        return {
            '熱處理':     'KG',
            '真空熱處理': 'KG',
            '氮化處理':   'KG',
            }
    def error_dic(self):
        return self.dic_err

    def group_to_list(self):
        return self.cost_group_list

    def dlookup_bmk_pdno(self, pdno): # 產品製程 
        return self.dic_bmk[pdno] # dict key: pdno,   value: ddaaframe 

    def dlookup_pdmd(self, pdno, column):
        # 從 pdno 找對應的 column欄位名稱的值value
        # MD001     MD002  MD003  MD004
        # 4A306025  KG     4.160    1.0
        result = ''
        df = self.df_md
        if pdno in df['MD001'].tolist():
            result = df.loc[df['MD001']==pdno][column].item()
        return result

    def dlookup_pdmd_tolist(self, pdno):
        # 從 pdno 找所有換算單位
        result = []
        df = self.df_md
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
 
    def isRegMath_float(self,findStr): #返回第一個被找到的數字值
        if any([findStr == None, findStr =='']):
            return ''
        mRegex = re.compile(r'\d+\.?\d*')
        match = mRegex.search(findStr)
        return match.group() if match else '' 

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
        # step 2 由材料找 
        df = self.df_bom
        df_w = df[df['pdno'].str.contains(r'^2.*')] # 材料件 品號2開頭
        for i, r in df_w.iterrows():
            if not (r['gid'] in arr_c):
                arr_c.append(r['gid'])
        
        lis_c = arr_c.tolist(); lis_c.sort() # 排序
        arr_c = array.array('i', lis_c) 
        # print('arr_c:', arr_c.tolist())

        # step 3 產生 group list
        # [[90], [89], [88, 87], [87], [86], [85, 84, 83]]
        lis_g = []
        arr_gid = array.array('i', self.df_bom['gid'].tolist()) # all gid
        arr_gid.reverse() # 反序遍歷
        # print(arr_gid.tolist()) 
        while len(arr_gid)>0:
            lis_t = [] # 暫存
            curr_id = arr_gid[0]
            arr_gid.remove(curr_id)
            loop_id = curr_id
            while loop_id in arr_c: # 是材料
                arr_c.remove(loop_id) # 移除arr_c 直到沒有 才不會無線循環
                if curr_id not in lis_t:
                        lis_t.append(loop_id)

                parent_id = gid(loop_id,'pid')
                if parent_id not in lis_t:
                        lis_t.append(parent_id)
                loop_id = parent_id
            else:
                if curr_id not in lis_t:
                    lis_t.append(curr_id)
            lis_g.append(lis_t)

        lis_g.reverse() # 恢復反序
        # print(lis_g)
        for lis in lis_g:
            if len(lis) > 1:
                lis.reverse()

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

        self.cost_group_list = lis_g

    def comp_2(self): # 計算宏觀產品製程成本
        df = self.df_bom
        re_f = self.isRegMath_float # re 文字找數字
        pmd = self.dlookup_pdmd # 找換算單位
        pmd_tolist = self.dlookup_pdmd_tolist # 從 pdno 找所有換算單位
        dic_spmd = self.spec_md_dic() #特定單位

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
            }

        lis_err4 = [] # gid-mk_i 的string array
        for i, r in df.iterrows():
            # 宏觀的產品製程( 包含產品製程資料 及 採購資料)  to_df() 方法
            df_m = pd.DataFrame(None, columns = list(dic_columns.keys())) # None dataframe
            if r['pd_type']=='P': # 品號屬性 P.採購件
                dic_p = {
                'MW002': '採購', # 製程
                'MF006': r['supply'], # 供應商代號
                'MF007': self.yst.get_pur_ma002(r['supply']), # 供應商簡稱'
                'MF017': 'PCS', # 加工單位
                'MF018': r['last_price'] # 最新進價(本國幣別NTD)
                }
                df_m = df_m.append(dic_p, ignore_index=True)


            elif r['pd_type'] in 'MSY': # 品號屬性 M.自製件,S.託外加工件,Y.虛設品號
                df_make = dic_msy[r['pdno']] # 該品號的製程
                df_m = df_m.append(df_make, ignore_index = True)

                # 檢查 3 有製程單位卻無換算單位
                dic_err3 = {}; lis_err3 = []
                lis_pmd = pmd_tolist(r['pdno']) # 該品號所有換算單位
                df_w = df_make.loc[df_make['MF017']!='PCS'] # 非PCS的加工單為
                lis_mk = []
                if len(df_w.index) > 0:
                    lis_mk = list(set(df_w['MF017'].tolist())) # 該品號製程所有加工單位(不含PCS)
                lis_err3 = list(filter(lambda e: e not in lis_pmd, lis_mk))
                if len(lis_err3) > 0:
                    dic_err3[r['gid']] = lis_err3

                # 檢查 4 特定製程的單位必為特定單位
                for mk_i, mk_r in df_make.iterrows():
                    if mk_r['MW002'] in list(dic_spmd.keys()):
                        if dic_spmd[mk_r['MW002']] != mk_r['MF017']:
                            lis_err4.append(f"{r['gid']}-{mk_i}")

            elif r['pd_type'] in 'Q': # 不展BOM銷售件
                dic_p = {
                'MW002': '銷售', # 製程
                'MF006': '', # 供應商代號
                'MF007': 'YEOSEH', # 供應商簡稱'
                'MF017': 'PCS', # 加工單位
                'MF018': r['sales_price_1'] # 最新進價(本國幣別NTD)
                }
                df_m = df_m.append(dic_p, ignore_index=True)

            dic_all[r['pdno']] = df_m
        self.dic_err['err3'] = dic_err3
        self.dic_err['err4'] = lis_err4

        # step 3 清洗 複製 SS001, SS002
        dic_all_new = {}
        for pdno, df_s in dic_all.items():
            df_new = df_s.copy()
            for i, r in df_s.iterrows():
                if r['MW002'] in ['採購','銷售']:
                    f_ss001 = r['MF018']
                else:
                    # 製造加工
                    if r['MF005'] == '1':
                        f_ss001 = re_f(r['MF023']) # 1自製抓 MF023 備註
                    elif r['MF005'] == '2':
                        f_ss001 = r['MF018']       # 2托外抓 MF018 加工單價
                    else:
                        f_ss001 = 0

                if pmd(pdno, 'MD002') == r['MF017']: # 有換算單位
                    f_ss002 = f_ss001 * pmd(pdno, 'MD003')
                else:
                    f_ss002 = f_ss001
                df_new.at[i,'SS001'] = f_ss001
                df_new.at[i,'SS002'] = f_ss002

            #  type object to numeric
            df_new['SS001'] = pd.to_numeric(df_new['SS001'], errors='coerce')
            df_new['SS002'] = pd.to_numeric(df_new['SS002'], errors='coerce')
            df_new[['SS001', 'SS002']] = df_new[['SS001','SS002']].fillna(value=0) # 空白補0
            # print(df_new.dtypes)
            dic_all_new[pdno] = df_new
        self.dic_bmk = dic_all_new

        # debug
        # print('step 3 new')
        # for k, v in dic_all_new.items():
        #     print(k)
        #     df = v[['MF001','MW002','MF017','MF018','MF023','SS001','SS002']]
        #     print(df)

    def comp_3(self): # 檢查資料異常
        df = self.df_bom
        # pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
        # pd.set_option('display.max_columns', None) #顯示最多欄位
        # df1 = df[['gid','pdno', 'pid','bom_level']]
        # print(df1)
        
        #檢查 1 最下階應為P, Q件，或應再建立P件為子件
        lis_err1 = []
        df_w = df.loc[(df['bom_level_lowest'] == True) & 
            (df['pd_type'] != 'P') & (df['pd_type'] != 'Q')] # P件 且 子件數量大於0
        # print(df_w)
        if len(df_w.index) > 0:
            lis_err1 = df_w['gid'].tolist()
        self.dic_err['err1'] = lis_err1
        
        #檢查 2 P件不應該有BOM架構，或有BOM應為S件or M件
        lis_err2 = []
        df_w = df.loc[(df['pd_type'] == 'P') & (df['child_quantity']>0)] # P件 且 子件數量大於0
        if len(df_w.index) > 0:
            lis_err2 = df_w['gid'].tolist()
        self.dic_err['err2'] = lis_err2

def test1():
    bom = COST('4A306019')
    print(bom.error_dic())
    # bom = COST('7AA01001A01')
    # print(bom.dlookup_bmk_pdno('5A090100005'))
    # bom = COST('5A090600004')
    # lis = bom.group_to_list()
    # print(lis)
    # bom = COST('5E00970')

if __name__ == '__main__':
    test1()