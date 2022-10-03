if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import time
import openpyxl
import PySimpleGUI as sg
from tool_excel2 import tool_excel
from tool_style import *
import tool_file
import tool_cost
import tool_db_yst
from tool_func import is_dec
from config import *

class Report_bcs01(tool_excel):
    def __init__(self, filename, pdno, pump_lock = False):
        sg.theme('SystemDefault')
        self.fileName = filename
        self.pdno = pdno
        self.pump_lock = pump_lock

        # if True: #debug
        #     sg.popup('\nPart number is not exist\nplease confirm whether your product number.\n\n維護中!', title='bom製程成本表')
        #     sys.exit()

        self.db = tool_db_yst.db_yst() # db
        if not self.db.is_exist_pd(self.pdno):
            sg.popup('\nPart number is not exist\nplease confirm whether your product number.\n\n品號不存在!', title='bom製程成本表')
            sys.exit()

        self.report_name = 'bcs01' # BOM製程成本表
        self.report_dir = config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑

        self.file_tool = tool_file.File_tool() # 檔案工具並初始化資料夾
        
        if self.report_path is None:
            print('找不到路徑')
            sys.exit() #正式結束程式
        self.file_tool.clear(self.report_name) # 清除舊檔
        self.comp_base()
        self.bcs = tool_cost.COST(self.pdno, self.pump_lock) # data
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

    def comp_base(self):
        # 基礎設定
        # name, width, sql_column_name
        lis_base = []; a = lis_base.append
        a('序號,    15,  none') 
        a('品號,    15,  none')
        a('品名規格, 38, none') 
        a('廠商,     8, MF006') 
        a('簡稱,    12, MF007')
        a('製號,     6, MF004')
        a('製程,    17, MW002') 
        a('製程敘述,15, MF008') 
        a('單位,    6, MF017') 
        a('製程單價, 10, SS001') 
        a('單價,      7, SS002') 
        a('工時批量,   10, MF019') 
        a('固定人時,   10, F_MF009') 
        a('固定機時,   10, F_MF024') 
        a('總用量,     9, none') 
        a('總單價,    9,  none') 
        a('金額,      9,  none') 
        a('試算,     14,  none') 
        lis_e1, lis_e2, lis_e3 = [],[],[]
        for e in lis_base:
            [e1, e2, e3]= e.split(',')
            lis_e1.append(e1.strip())
            lis_e2.append(int(e2.strip()))
            lis_e3.append(e3.strip())
        self.xls_index =dict(zip(lis_e1, [e for e in range(1,len(lis_e1)+1)]))
        self.xls_width =dict(zip(lis_e1, lis_e2))
        self.xls_sqlcn =dict(zip(lis_e1, lis_e3))
    def output(self):
        caption = 'BOM製程成本表' # 標題
        if True: # style, func
            f9g =font_9_Calibri_g
            f10 = font_10_Calibri
            f11 = font_11_Calibri
            f11gr = font_11_Calibri_green
            ahr = ah_right
            # func, method
            write=self.c_write; fill=self.c_fill; comm=self.c_comm; img=self.c_image2; column_w=self.c_column_width
            bmr = self.bcs.dlookup_gid     # bom品號資料
            md = self.bcs.dlookup_pdmd     # 單位換算
            mk = self.bcs.dlookup_bmk_pdno # 產品製程
            pkg = self.bcs.dlookup_gid_pkg # 產品材料規格KG
            # number_format
            nf_total = '$#,##0.0'

        x_i = self.xls_index
        x_width = self.xls_width
        x_sqlcn = self.xls_sqlcn
        cr=1; column_w(list(x_width.values())) # 設定欄寬
        for name, index in x_i.items():
            write(cr, index, name, f11, alignment=ah_wr, border=bt_border, fillcolor=cf_gray) # 欄位名稱

        total = 0 # 試算
        err_dic = self.bcs.error_dic()
        lis_group = self.bcs.group_to_list()
        for group_idx, lis_gid in enumerate(lis_group):
            price = 0 # 總單價
            cr+=1; write(cr, x_i['序號'], group_idx+1, f11) # 序號
            cr-=1
            for pdno_idx, gid in enumerate(lis_gid):
                pdno = bmr(gid,'pdno')
                cr+=1; write(cr, x_i['品號'], pdno, f10, alignment=ah_wr) # 品號 
                crm = cr  # 紀錄位置 crm
                # 自身總用量 self_quantity
                if pdno_idx == 0:
                    quantity = float(bmr(gid,'self_quantity'))
                    write(cr, x_i['總用量'], quantity, f11, alignment=ahr) # 自身總用量

                if pdno_idx ==0:
                    cr_pd = cr #紀錄位置 首個cr_pd
                    max_height = 58 if bmr(gid,'bom_level_lowest') == True else 116
                    img(cr, 1, pdno ,0.1,0.4,max_height=max_height)           # img
                cr+=1; write(cr, x_i['品號'], gid, f9g, alignment=ahr)     # gid
                cr-=1
                cr+=0; write(cr, x_i['品名規格'], bmr(gid,'pd_name'), f11, alignment=ah_wr) # 品名
                cr+=1; write(cr, x_i['品名規格'], bmr(gid,'pd_spec'), f11, alignment=ah_wr) # 規格
                mol_den = float(bmr(gid,'bom_mol')/bmr(gid,'bom_den'))       # 用量換算率
                cr+=1; write(cr, x_i['品名規格'], self.fmat_1(gid), f10)  # x階(x件),用量/底數
                if md(pdno,'MD002') != '':
                    # tmpstr = f"{md(pdno,'MD003')}/{md(pdno,'MD004')}{md(pdno,'MD002')}"
                    cr+=1; write(cr, x_i['品名規格'], self.fmat_md(pdno), f10) #換算單位 0.5KG /1.0PCS

                # 產品製程
                crm -= 1
                df_mk = mk(pdno) # 產品製程 dataframe
                for mk_i, mk_r in df_mk.iterrows():
                    crm+=1;
                    sn='廠商';    write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f9g)
                    sn='簡稱';    write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f11)
                    sn='製號';    write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f9g)
                    sn='製程';    write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f11, alignment=ah_wr)
                    sn='製程敘述'; write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f10, alignment=ah_wr)
                    sn='單位';    write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f11)
                    sn='製程單價'; write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f11, alignment=ahr)
                    # 顯使材料kg
                    if all([bmr(gid,'pd_type')=='P', len(df_mk.index)==1,
                        mk_r[x_sqlcn['單位']]=='PCS']): # P採購件 且 宏觀製程僅1筆 且單位為PCS
                        pd_kg = pkg(gid)
                        if all([mk_r[x_sqlcn['製程單價']]>0, pd_kg != '', pd_kg != 0]): #有單價 有KG資料
                            price_kg = round(float(mk_r[x_sqlcn['製程單價']])/float(pd_kg), 2)
                            write(crm+1, x_i[sn], f'({price_kg}/KG)', f11gr, alignment=ahr)

                    # 單價 =   製程單價(加工單價*單位換算)  * 用量換算率(用量/底數)
                    price_mol_den = float(mk_r['SS002']) * min(mol_den, 1) # 小於1時應換算(取最小且最大為1)
                    sn='單價';    write(crm, x_i[sn], price_mol_den, f11, alignment=ahr)
                    sn='工時批量'; write(crm, x_i[sn], mk_r[x_sqlcn[sn]], f11, alignment=ahr)
                    if mk_r['MF005'] == '1': # 自製
                        sn='固定人時';v=mk_r[x_sqlcn[sn]];v=str(v).replace('00:00:00','');v=v.replace('nan','');write(crm,x_i[sn],v,f11)
                        sn='固定機時';v=mk_r[x_sqlcn[sn]];v=str(v).replace('00:00:00','');v=v.replace('nan','');write(crm,x_i[sn],v,f11)
                    price += price_mol_den # 總單價
                    # debug
                    # write(crm, 18, mk_r['SS031'], f11) # 托外最新進價

                    # 檢查 單價為0
                    if price_mol_den == 0:
                        sn='製程單價'; fill(crm, x_i[sn])
                        sn='單價';    fill(crm, x_i[sn])

                    # 檢查 禁止交易供應商
                    if pdno in list(err_dic['err4'].keys()):
                        if err_dic['err4'][pdno]['mk_i'] == mk_i:
                            sn='簡稱'; cj=x_i[sn]; fill(crm, cj); comm(crm, cj, err_dic['err4'][pdno]['mssage'])

                    # 檢查 已有最新托外進價
                    if pdno in list(err_dic['err5'].keys()):
                        if err_dic['err5'][pdno]['mk_i'] == mk_i:
                            sn='製程單價'; cj=x_i[sn]; fill(crm, cj, fillcolor=cf_khaki); comm(crm, cj, err_dic['err5'][pdno]['mssage'])

                    # 檢查 不吻合 標準廠商加工單價
                    if pdno in list(err_dic['err6'].keys()):
                        if err_dic['err6'][pdno]['mk_i'] == mk_i:
                            sn='製程單價'; cj=x_i[sn]; fill(crm, cj); comm(crm, cj, err_dic['err6'][pdno]['mssage'])

                    # 檢查 不吻合 標準途程加工單位
                    if pdno in list(err_dic['err7'].keys()):
                        if err_dic['err7'][pdno]['mk_i'] == mk_i:
                            sn='單位'; cj=x_i[sn]; fill(crm, cj); comm(crm, cj, err_dic['err7'][pdno]['mssage'])

                    # 檢查 # 最後N筆交易紀錄 是否已改變供應商
                    if pdno in list(err_dic['err8'].keys()):
                        sn='簡稱'; cj=x_i[sn]; fill(crm, cj, fillcolor=cf_khaki); comm(crm, cj, err_dic['err8'][pdno]['mssage'])

                    # 檢查 # 採購加工單位
                    if pdno in list(err_dic['err9'].keys()):
                        if err_dic['err9'][pdno]['mk_i'] == mk_i:
                            sn='單位'; cj=x_i[sn]; fill(crm, cj); comm(crm, cj, err_dic['err9'][pdno]['mssage'])

                    # 檢查 # 採購加工單價
                    if pdno in list(err_dic['err10'].keys()):
                        if err_dic['err10'][pdno]['mk_i'] == mk_i:
                            sn='製程單價'; cj=x_i[sn]; fill(crm, cj, fillcolor=cf_khaki); comm(crm, cj, err_dic['err10'][pdno]['mssage'])

                    # 檢查 # 未設定單位換算
                    if pdno in list(err_dic['err11'].keys()):
                        if err_dic['err11'][pdno]['mk_i'] == mk_i:
                            sn='品號'; cj=x_i[sn]; fill(crm, cj); comm(crm, cj, err_dic['err11'][pdno]['mssage'])

                    # 檢查 # 未更新單位換算
                    if pdno in list(err_dic['err12'].keys()):
                        if err_dic['err12'][pdno]['mk_i'] == mk_i:
                            sn='品名規格'; cj=x_i[sn]; fill(crm+3, cj); comm(crm+3, cj, err_dic['err12'][pdno]['mssage'])

                # 檢查
                if gid in err_dic['err1']:
                    sn='品名規格'; cj=x_i[sn]; fill(cr_pd+2,cj); comm(cr_pd+2,cj, '最下階應為P件，或應再建立P件為子件')

                if gid in err_dic['err2']:
                    sn='品名規格'; cj=x_i[sn]; fill(cr_pd+2,cj); comm(cr_pd+2,cj, 'P件不應該有BOM架構，或有BOM應為S件or M件')

                cr = max(cr, crm)
                # 群組底
                if pdno_idx == len(lis_gid)-1:
                    self.c_line_border(cr,1,len(x_i.keys()),border=bottom_border) # 畫線

            price=float(f'{price:0.3f}') # 強制取小數至3位數即停
            nf = '0.##' if is_dec(price) else 'General'
            sn='總單價'; write(cr_pd, x_i[sn], price, f11, alignment=ahr, number_format=nf)

            money=price*quantity
            nf = '0.##' if is_dec(price) else 'General'
            sn='金額'; write(cr_pd, x_i[sn], money, f11, alignment=ahr, number_format=nf) 

            total+=money
            sn='試算'; write(cr_pd, x_i[sn], total, f11gr, alignment=ahr, number_format=nf_total) 

        cr+=1; write(cr, 1, '-結束- 以下空白', alignment=ah_center_top); self.c_merge(cr,1,cr,len(x_i.keys()))

    def fmat_1(self, gid):
        # 格式化 1
        bmr = self.bcs.dlookup_gid
        bom_level = bmr(gid,'bom_level')
        pd_type = bmr(gid,'pd_type')
        a = f"{bmr(gid,'bom_mol'):.3f}"; a=a.rstrip('0'); a=a.rstrip('.'); bom_mol=a
        b = f"{bmr(gid,'bom_den'):.3f}"; b=b.rstrip('0'); b=b.rstrip('.'); bom_den=b
        return f'{bom_level}階({pd_type}件), 用量/底數:{bom_mol}/{bom_den}'

    def fmat_md(self, pdno):
        # 格式化 單位換算
        md = self.bcs.dlookup_pdmd
        md02 = md(pdno,'MD002') # 換算單位
        md03 = md(pdno,'MD003') # 換算率分子
        md04 = md(pdno,'MD004') # 換算率分母
        if not is_dec(md04):
            return f'{md03}{md02}/{md04:.0f}PCS'
        else:
            return f'{md03}{md02}/{md04}PCS'

def test1():
    fileName = 'bcs01' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    # Report_bcs01(fileName, '4A505051')
    # Report_bcs01(fileName, '4B104018-01')
    # Report_bcs01(fileName, '5A220100004')
    Report_bcs01(fileName, '6EB0028')
    # Report_bcs01(fileName, '8FC026', True)
    
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式