if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import time
import openpyxl
from tool_excel2 import tool_excel
from tool_style import *
import tool_file
import tool_cost
from config import *

class Report_bcs01(tool_excel):
    def __init__(self, filename, pdno, pump_lock = False):
        self.fileName = filename
        self.pdno = pdno
        self.pump_lock = pump_lock
        self.report_name = 'bcs01' # BOM製程成本表
        self.report_dir = config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑

        self.file_tool = tool_file.File_tool() # 檔案工具並初始化資料夾
        
        if self.report_path is None:
            print('找不到路徑')
            sys.exit() #正式結束程式
        self.file_tool.clear(self.report_name) # 清除舊檔
        self.bcs = tool_cost.COST(self.pdno, self.pump_lock) # data
        self.dic_spmd = self.bcs.spec_md_dic() #特定單位
        # print(self.bcs.error_dic())

        self.create_excel()  # 建立
        self.output()
        self.save_xls()
        self.open_xls() # 開啟

    def create_excel(self):
        wb = openpyxl.Workbook()
        sh = wb.active
        sh.title = self.report_name
        self.xlsfile = os.path.join(self.report_path, self.fileName)
        wb.save(filename = self.xlsfile)
        super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
        self.set_page_layout_horizontal()

    def output(self):
        caption = 'BOM製程成本表' # 標題
        if True: # style, func
            f11 = font_11; f10 = font_10
            # func, method
            err_dic = self.bcs.error_dic()
            write = self.c_write
            gid = self.bcs.dlookup_gid     # bom品號資料
            md = self.bcs.dlookup_pdmd     # 單位換算
            mk = self.bcs.dlookup_bmk_pdno # 產品製程

        # write(1, 1, caption, f11) #標題
        # write(2, 1, f'查詢時間: now', f11) #標題
        columns = ['序號',     '品號',   '品名規格', '廠商代號', '簡稱',
                   '製程',     '單位',   '製程單價', '單價',     '工時批量',
                   '固定人時', '固定機時','總用量',  '總單價',    '金額',
                   '試算']
        column_width =[15,15,32, 10, 12,
                       17, 5 ,10, 7, 10,
                       10, 10, 9, 9, 9,
                       9] #欄寬
        self.c_column_width(column_width) # 設定欄寬
        for i, e in enumerate(columns):
            write(1, i+1, e, f11, alignment=ah_wr, border=bt_border, fillcolor=cf_gray) # 欄位名稱

        cr = 1
        total = 0 # 試算
        lis_group = self.bcs.group_to_list()
        for group_idx, lis_gid in enumerate(lis_group):
            price = 0 # 總單價
            cr+=1; write(cr, 1, group_idx+1, f11)

            cr-=1
            for pdno_idx, v in enumerate(lis_gid):
                pdno = gid(v,'pdno')
                cr+=1; write(cr, 2, pdno, f10, alignment=ah_wr) # 品號 
                write(cr+1, 2, v, font_A_10G, alignment=ah_right) # gid
                crm = cr  # 紀錄位置 crm
                if pdno_idx ==0:
                    cr_pd = cr #紀錄位置 首個cr_pd
                    # image
                    max_height = 58 if gid(v,'bom_level_lowest') == True else 116
                    self.c_image2(cr, 1, pdno ,0.1,0.4,max_height=max_height)

                write(cr, 3, gid(v,'pd_name'), f11, alignment=ah_wr) # 品名
                
                # 自身總用量 self_quantity
                if pdno_idx == 0:
                    quantity = float(gid(v,'self_quantity'))
                    write(cr, 13, quantity, f11, alignment=ah_right) # 自身總用量 self_quantity
                cr+=1; write(cr, 3, gid(v,'pd_spec'), f11, alignment=ah_wr) # 規格
                # 用量/底數
                str1 = f"{gid(v,'bom_level')}階({gid(v,'pd_type')}件),用量/底數:{gid(v,'bom_mol'):.3f}/{gid(v,'bom_den'):.3f}"
                # 用量換算率
                mol_den = float(gid(v,'bom_mol')/gid(v,'bom_den'))
                cr+=1; write(cr, 3, str1, f10)
                # 換算單位
                if md(pdno,'MD002') != '':
                    str2 = f"{md(pdno,'MD003')}/{md(pdno,'MD004')}{md(pdno,'MD002')}"
                    cr+=1; write(cr, 3, str2, f10)

                # 產品製程
                crm -= 1
                df_mk = mk(pdno) # 產品製程 dataframe
                for mk_i, mk_r in df_mk.iterrows():
                    crm+=1; write(crm, 4, mk_r['MF006'], f11) # 廠商代號 
                    write(crm, 5, mk_r['MF007'], f11) # 簡稱
                    write(crm, 6, mk_r['MW002'], f11, alignment=ah_wr) # 製程
                    write(crm, 7, mk_r['MF017'], f11) # 加工單位

                    write(crm, 8, float(mk_r['SS001']), f11, alignment=ah_right) # 製程單價(後製資料), 抓取順序1.備註 2.加工單價
                    
                    # 單價 =   製程單價(加工單價*單位換算)  * 用量換算率(用量/底數)
                    price_mol_den = float(mk_r['SS002']) * min(mol_den, 1) # 小於1時應換算(取最小且最大為1)
                    write(crm, 9, price_mol_den, f11, alignment=ah_right)

                    write(crm, 10, mk_r['MF019'], f11, alignment=ah_right) # 工時批量
                    write(crm, 11, mk_r['F_MF009'], f11) # 固定人時
                    write(crm, 12, mk_r['F_MF024'], f11) # 固定機時
                    price += price_mol_den # 總單價

                    # 檢查 單價為0
                    if price_mol_den == 0:
                        self.c_fill(crm, 8); self.c_fill(crm, 9)

                    # 檢查 加工單位 非 固定單位
                    if f"{v}-{mk_i}" in err_dic['err4']:
                        self.c_fill(crm, 7); self.c_comm(crm, 7, f"{mk_r['MW002']}加工單位應為 {self.dic_spmd[mk_r['MW002']]}")

                # 檢查
                if v in err_dic['err1']:
                    self.c_fill(cr_pd+2,3); self.c_comm(cr_pd+2,3, '最下階應為P件，或應再建立P件為子件')

                if v in err_dic['err2']:
                    self.c_fill(cr_pd+2,3); self.c_comm(cr_pd+2,3, 'P件不應該有BOM架構，或有BOM應為S件or M件')

                cr = max(cr, crm)
                # 群組底
                if pdno_idx == len(lis_gid)-1:
                    self.c_line_border(cr,1,16,border=bottom_border) # 畫線


            price=float(f'{price:0.2f}')
            v=f'{price:0.2f}'; v= v.rstrip('0'); v=v.rstrip('.'); write(cr_pd, 14, v, f11, alignment=ah_right) # 總單價

            money=float(f'{price*quantity:0.2f}')
            v=f'{money:0.2f}'; v=v.rstrip('0'); v=v.rstrip('.'); write(cr_pd, 15, v, f11, alignment=ah_right) # 金額

            total+=money; v=f'{total:,.1f}'; v=v.rstrip('0'); v=v.rstrip('.'); write(cr_pd, 16, v, font_11_green, alignment=ah_right) # 試算
        cr+=1; write(cr, 1, '-結束- 以下空白', alignment=ah_center_top); self.c_merge(cr,1,cr,len(columns))

def test1():
    fileName = 'bcs01' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    # Report_bcs01(fileName, '5A110100005')
    # Report_bcs01(fileName, '4A306019')
    # Report_bcs01(fileName, '5F00001')
    # Report_bcs01(fileName, '6AA03JA001AL1A01')
    Report_bcs01(fileName, '7AA01001A01', True)
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式