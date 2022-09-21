import pandas as pd
import numpy as np

def divideList(dataList:pd.DataFrame):
    data = dataList.copy()
    aList = pd.DataFrame()
    bList = pd.DataFrame()
    data.insert(22, "dType", "")
    ordersNumbers = data["주문번호"].unique()
    for orNum in ordersNumbers:
        temp = data[data["주문번호"].isin([orNum])]
        shoeBox = temp[temp["상품명"].str.contains("슈박스")]
        if len(shoeBox) > 0:
            aList = pd.concat([aList, temp])
        else:
            bList = pd.concat([bList, temp])
    return aList, bList

def SixDivideList(dataList:pd.DataFrame):
    data = dataList.copy()
    _result = pd.DataFrame()
    for i in range(len(data)):
        temp = data.iloc[i:i+1,:]
        filter_01 = temp["상품명"].str.contains("1세트").iloc[0]
        filter_03 = (temp["수량"] > 6).iloc[0]
        if ~filter_01 & filter_03:
            subtemp = temp.copy()
            elsebox = (temp["수량"] % 6).iloc[0]
            sixbox = (temp["수량"] // 6).iloc[0]
            if elsebox > 0:
                elsebox += 6
                if elsebox % 2 == 0:
                    subtemp["수량"] = elsebox // 2
                    for i in range(2):
                        _result = pd.concat([_result, subtemp])
                else:
                    for i in range(2):
                        subtemp["수량"] = (elsebox // 2) + i
                        _result = pd.concat([_result, subtemp])
                if sixbox > 1:
                    for i in range(sixbox-1):
                        subtemp["수량"] = 6
                        _result = pd.concat([_result, subtemp])
            else:
                for i in range(sixbox):
                    subtemp["수량"] = 6
                    _result = pd.concat([_result, subtemp])
        else:
            _result = pd.concat([_result, temp])
    return _result

def GrassSubmit(gList:pd.DataFrame):
    data = gList.copy()
    unique = gList["주문번호"].unique()
    _result = pd.DataFrame()
    for order in unique:
        temp = data[data["주문번호"].isin([order])]
        shoeBox = temp[temp["상품명"].str.contains("슈박스")]
        orders = temp[~temp["옵션정보"].str.contains("잔디매트: 8mm") & ~temp["옵션정보"].str.contains("잔디매트: 20mm")]
        g08 = temp[temp["옵션정보"].str.contains("잔디매트: 8mm")] 
        g20 = temp[temp["옵션정보"].str.contains("잔디매트: 20mm")]
        _result = pd.concat([_result, orders])
        
        if len(g08) - len(shoeBox) > 0:
            sg08 = g08.sort_values(by=["수량"]).copy()
            sg08.iloc[1,21] = sg08.iloc[0:2,:]["수량"].sum()
            sg08 = sg08.iloc[1:,:]
            _result = pd.concat([_result, sg08])
        else:
            _result = pd.concat([_result, g08])
        if len(g20) - len(shoeBox) > 0:
            sg20 = g20.sort_values(by=["수량"]).copy()
            sg20.iloc[1,21] = sg20.iloc[0:2,:]["수량"].sum()
            sg20 = sg20.iloc[1:,:]
            _result = pd.concat([_result, sg20])
        else:
            _result = pd.concat([_result, g20])
    return _result

def typeChack(tList:pd.DataFrame):
    _result = tList.copy()
    pl = _result["상품명"].str.contains("프리미엄")
    pl_f1 = _result["수량"] == 2
    pl_f2 = _result["수량"] > 2
    pl_f3 = _result["수량"] < 5
    pl_f4 = _result["수량"] >= 5
    bl = _result["상품명"].str.contains("신발 슈케이스")
    bl_f1 = _result["수량"] > 1
    bl_f2 = _result["수량"] < 4
    bl_f3 = _result["수량"] > 3
    _result.loc[(pl&pl_f1)|(bl&(bl_f1&bl_f2)) ,'dType'] = "b"
    _result.loc[(pl&(pl_f2&pl_f3))|(bl&bl_f3) ,'dType'] = "c"
    _result.loc[(pl&pl_f4) ,'dType'] = "d"
    _result.loc[(~pl&~bl) ,'dType'] = "s"
    return _result

def userCheck(orderData:pd.DataFrame):
    orderUser = pd.read_excel("./data/블랙리스트(슈케이브).xlsx", header=2)
    orderPhone = orderData["수취인연락처1"].values
    userPhone = orderUser["연락처"].values
    result = any(np.in1d(orderPhone,userPhone))
    return result