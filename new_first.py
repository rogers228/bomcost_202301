if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import PySimpleGUI as sg

def test1():
    # print(path_list)
    # print(sys.path)
    sg.theme('SystemDefault')
    sg.popup(sys.path)  # Shows red error button

if __name__ == '__main__':
    test1()