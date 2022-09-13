if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)
    
# 權限
import tool_db_hr

class Authorization():
    def __init__(self):
        self.hr = tool_db_hr.db_hr()

    def isqs(self, qsno): # 檢查是否擁有權限
        # qsno: 權限id
        computer_name = os.getenv('COMPUTERNAME')
        # print('computer_name:', computer_name)
        if computer_name in ['VM-TESTER']:
            return True # 開發者環境 擁有權限

        result = False
        
        no = self.hr.ps18Getps01(computer_name) #電腦名稱 搜尋使用者
        # print('no1:', no)
        if no == '':
            no = self.hr.pc02Getpc01(computer_name) # 設備名稱搜尋使用者
        # print('no2:', no)

        lis_qs = self.hr.cpGer_qs_lis(no) # 全限列表
        # print('lis_qs:', lis_qs)
        
        if qsno in lis_qs:
            # print(f'{computer_name} 擁有 {qsno} 的權限')
            result = True
        # else:
        #     print(f'{computer_name} 沒有 {qsno} 權限!')

        return result

def test1():
    au = Authorization()
    # au.isqs(701)
    print(au.isqs(701))
    

if __name__ == '__main__':
    test1()