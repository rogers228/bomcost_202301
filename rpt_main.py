if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['bomcost_202301'] # 目前路徑
    sys.path.append(config_path)

import time
import click
import tool_auth
import rpt_bcs01

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning) # 抑制未來警告

@click.command() # 命令行入口
@click.option('-report_name', help='report name', required=True, type=str) # required 必要的
@click.option('-pdno', help='product no.', required=True, type=str) # required 必要的
@click.option('-pump_lock', help='pump lock', type=bool) # required 必要的
def main(report_name, pdno, pump_lock= False):
    au = tool_auth.Authorization()
    if not au.isqs(12): # 檢查 12 權限
        click.echo('無權限!')
        return # 無權限 退出

    filename = report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    rpt_bcs01.Report_bcs01(filename, pdno, pump_lock)

if __name__ == '__main__':
    main()
    # cmd
    # C:\python_venv\python.exe \\220.168.100.104\pdm\python_program\bomcost_202301\rpt_main.py -report_name bcs01 -pdno 4A306019
