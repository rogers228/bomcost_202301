if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import array
import tool_db_yst
import tool_bom, tool_bmk


class COST(): # 基於bom 與 製程bmk 合併產生出 cost data
    def __init__(self, pdno, pump_lock = False):
        #tool
        self.yst = tool_db_yst.db_yst()

        self.df_bom = tool_bom.BOM(pdno, pump_lock).to_df() # BOM
        self.df_md = self.yst.wget_imd(','.join(list(set(self.df_bom['pdno'].tolist())))) #所有品號單位換算

        df = self.df_bom
        pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
        pd.set_option('display.max_columns', None) #顯示最多欄位
        df = df[['gid','pid','pdno','pd_name']]
        print(df)

        print(self.df_md)

        self.comp_1() # rerutn  self.cost_group_list
        print(self.cost_group_list)
        # self.comp_2()

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
        # print(arr_c.tolist())

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
        self.cost_group_list = lis_g

    # def comp_2(self):


def test1():
    # bom = COST('6AA03JA001AL1A01', pump_lock = True)
    bom = COST('5A160800006')
    # bom = COST('5E00970')

    # print(df)
    # # lis_columns = list(bom.get_columns().keys())
    # # print(lis_columns)


if __name__ == '__main__':
    test1()                