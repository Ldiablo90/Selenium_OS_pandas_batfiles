import pandas as pd
import LeeFunction as lf
import os
from datetime import date, datetime
import naver_enum as nem
import msoffcrypto
import io
from tkinter import messagebox

def excelpassok(path):
    decrypted = io.BytesIO()
    with open(path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=nem.EXCELPASSWORD)  # Use password
        file.decrypt(decrypted)
    return pd.read_excel(decrypted, header=1)
def excelpassnot(path):
    return pd.read_excel(path, header=1)

def resultFile(path):
    if os.path.isfile(path):
        print(path)
        try:
            read_data = excelpassok(path)
        except:
            read_data = excelpassnot(path)
        onlyShoeBox, onlyRubby = lf.divideList(read_data)
        if len(onlyShoeBox) > 0:
            onlyShoeBox = lf.SixDivideList(onlyShoeBox)
            sortShoeBox = lf.GrassSubmit(onlyShoeBox)
            concatData = pd.concat([sortShoeBox, onlyRubby])
            endData = lf.typeChack(concatData)
            endData = endData.fillna("")
        else:
            typeData = lf.typeChack(onlyRubby)
            endData = typeData.fillna("")
        userPath = "../블랙리스트(슈케이브).xlsx"
        if os.path.isfile(userPath):
            orderUser = lf.userCheck(endData,userPath)
            if(orderUser):messagebox.showwarning("유저발견","파일을 확인해 주세요.")
        return endData
    else:
        return []

def saveFile(datas):
    datas["상품주문번호"] = datas["상품주문번호"].apply(lambda x: '{:d}'.format(x))
    datas["주문번호"] = datas["주문번호"].apply(lambda x: '{:d}'.format(x))
    amOrpm = "오후" if datetime.now().hour >= 14 else "오전"
    datas.to_excel("./result/스마트스토어_%s_%s.xlsx"%(amOrpm, date.today()), index=False)

orderpath = os.getcwd().replace("\\","/")+'/data/'

filename = str(date.today()).replace("-","")
files = files = [filetitle for filetitle in os.listdir(orderpath) if filetitle.__contains__(filename)]
print(files)

if len(files) > 0 :
    createTime = datetime.fromtimestamp(os.path.getctime(orderpath+files[-1]))
    if datetime.now().hour == createTime.hour:
        datas = resultFile(orderpath+files[-1])
        saveFile(datas)
        messagebox.showinfo("저장완료", "전처리파일이 저장 완료되었습니다.%s"%(files[-1]))
    else:
        print()
        messagebox.showwarning("저장실패", "전처리할 시간이 아닙니다.")
    messagebox.showinfo("저장완료", "전처리파일이 저장 완료되었습니다.%s"%(files[-1]))
else:
    messagebox.showwarning("저장실패", "전처리할 파일이 없습니다.")
# else:
#     messagebox.showwarning("저장실패", "전처리할 파일이 없습니다.")