from datetime import date, datetime, timezone, timedelta
import os
import os.path as ospath
import shutil
import getpass
from tkinter import messagebox

def os_move_rename():
    print('start-os_move_rename')
    results = False
    #================== 날짜 ==================
    filenstr = "스마트스토어_전체주문발주발송관리_"+ str(date.today()).replace("-","")
    #================== 모든 경로 ==================
    orderpath = './data/'
    dowunlodepath = 'C:/Users/%s/Downloads/'%(getpass.getuser())
    #================== 파일 이름 ==================
    files = [file for file in os.listdir(dowunlodepath) if file.__contains__(filenstr)]
    #================== 파일 확인 ==================
    if len(files) > 0:
        createTime = datetime.fromtimestamp(ospath.getctime(dowunlodepath+files[-1]))
        amOrpm = "오후" if createTime.hour >= 14 else "오전"
        if(datetime.now().hour == createTime.hour):
            shutil.move(dowunlodepath+files[-1], orderpath+amOrpm+files[-1])
            messagebox.showinfo("파일이동성공", "파일이동이 완료되었습니다.\n%s"%(files[-1]))
            results = True
        else:
            messagebox.showwarning("파일이동실패", "현시간과 맞는 파일이 아닙니다.\n%s"%(files[-1]))

    else:
        messagebox.showwarning("파일이동실패", "찾는 파일이 없습니다.")
        print('file is no serch\nend-os_move_rename')
    return results

os_move_rename()