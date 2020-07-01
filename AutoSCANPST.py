#!/usr/bin/python
# -*- coding: big5 -*-

import sys, os
pwd = os.path.dirname(os.path.realpath(__file__))+"\\"
sys.path.append(pwd)

import glob
import autoit

def ErrorHandle():
    from traceback import format_exc
    print(format_exc())
    input("Exit with error! Press any key to exit...")
    #sys.exit(1)

def user_control_wait_text(title, control, text):
    while True:
        try:
            if text in autoit.control_get_text(title, control):
                break
        except:
            #ErrorHandle()
            continue

scanpstPath = 'C:\Program Files (x86)\Microsoft Office\Office15\SCANPST.EXE' # path of SCANPST
scanpstTitle = '[CLASS:#32770]' # title of SCANPST

for pstPath in glob.glob('D:\mail\*.pst'):
    print("Processing %s" % pstPath)

    autoit.run(scanpstPath) # open exe
    user_control_wait_text(scanpstTitle, "Static2", '請輸入您要掃描的檔案名稱') # wait win active (background handle)
    #autoit.win_set_state(scanpstTitle, autoit.autoit.Properties.SW_SHOWNOACTIVATE)
    autoit.control_set_text(scanpstTitle, "Edit1", pstPath) # set pst path
    autoit.control_click(scanpstTitle, "Button2") # 啟動
    
    user_control_wait_text(scanpstTitle, "Static2", pstPath) # wait scan finish
    
    # judgement result
    rstMsg = autoit.control_get_text(scanpstTitle, "Static3")
    print(rstMsg)
    if '沒有在這個檔案中找到錯誤' in rstMsg:
        autoit.control_click(scanpstTitle, "Button5") # 確定
    else:
        # if need to fix
        autoit.control_click(scanpstTitle, "Button1") # 取消勾選備份
        autoit.control_click(scanpstTitle, "Button4") # 修復
        fixMsg = '修復完成'
        user_control_wait_text(scanpstTitle, "Static2", fixMsg)
        autoit.control_click(scanpstTitle, "Button1") # 確定
        print(fixMsg)

    print("") # new line
