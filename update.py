# update to server
# 由server開發環境方能執行同步 
# 開發資料夾 同步至server 以利工作站執行

from config import *
import dirsync

def main():
    args = {
        'purge': True,   # 同步清除
        'create' : True, # 資料夾不存在時則建立
        'ignore' : ['\.git', '\.gitignore', 'old', 'test.*', 'update.py'] # 忽略
    }
    dirsync.sync(config_develop_program, config_servr_program, 'sync', **args)

    print('update is finished')

if __name__ == '__main__':
    main()