from selenium import webdriver
from tkinter import messagebox
import time
import naver_enum as nem
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def chromedriver(_url):
    options = webdriver.ChromeOptions()
    # options.headless = True
    options.add_argument("--window-size=1900,900")
    options.add_argument("--disable-gpu")
    _driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    _driver.get(_url)
    _driver.implicitly_wait(10)
    return _driver

def xpath_send_keys(_driver,_xpath,_send_keys):
    time.sleep(1)
    _driver.implicitly_wait(10)
    _driver.find_element( By.XPATH,_xpath).send_keys(_send_keys)
        
def xpath_click(_driver, _xpath):
    time.sleep(1)
    _driver.implicitly_wait(10)
    _driver.find_element( By.XPATH,_xpath).click()
    
def xpath_enter(_driver, _xpath):
    time.sleep(1)
    _driver.implicitly_wait(10)
    _driver.find_element( By.XPATH,_xpath).send_keys("\n")

def xpath_text(_driver,_xpath):
    time.sleep(1)
    _driver.implicitly_wait(10)
    value = _driver.find_element( By.XPATH,_xpath).text
    print(value)
    return value

def changeframe(_driver, _id):
    time.sleep(3)
    _driver.implicitly_wait(10)
    _driver.switch_to.frame(_driver.find_element( By.ID, _id))
    return _driver

def backframe(_driver):
    _driver.default_content()
    return _driver

def alertaccept(_driver):
    _driver.implicitly_wait(10)
    time.sleep(5)
    alert = _driver.switch_to.alert
    alert.accept()

def checkingbtn(_driver, _xpath):
    time.sleep(3)
    _driver.implicitly_wait(10)
    btn = _driver.find_element(By.XPATH,_xpath).isDisplayed()
    if (btn):
        print("ok?")
#################### 실행부분 ######################

driver = chromedriver(nem.URL)
################### 로그인부분 ################################
xpath_click(driver, nem.LOGINPAGEMOVE)                       #
xpath_send_keys(driver, nem.SIGININPUTID, nem.NAVERID)       #
xpath_send_keys(driver, nem.SIGININPUTPASS, nem.NVAERPASS)   #
xpath_click(driver, nem.SIGNIN)                              #
##############################################################
try:
    xpath_click(driver,nem.TEMPBUTTON)
except:
    print("빅파워없음")

################### 화전전환부분 #############################
if xpath_text(driver, nem.STORENAME) == "제이에이치디자인":  #
    xpath_click(driver, nem.STOREMOVE)                      #
    xpath_click(driver, nem.STOREMOVEBTN)                   #
    time.sleep(5)                                           #
#############################################################

xpath_click(driver, nem.SALESMANAGEMENT)
xpath_click(driver, nem.ORDERANDSEND)
# iframe 이동하기
driver = changeframe(driver, nem.IFRAME)

preordercount = int(xpath_text(driver, nem.PREORDERCOUNT))
newordercount = int(xpath_text(driver, nem.NEWORDERCOUNT))

if  preordercount: #예약구매 있는지 확인하기
    xpath_click(driver, nem.PREORDERPAGE) # 예약구매 페이지로 이동
    xpath_click(driver, nem.ALLCHECK) # 상품 전체선택
    xpath_click(driver, nem.SENDCHECK) # 상품 확인선택
    alertaccept(driver) # 알람 확인 버튼 누르기

if newordercount:# 신규주문 확인
    try:
        xpath_click(driver,nem.TEMPBUTTON) # 짜증나는 빅파워
    except:
        print("빅파워없음")
    if newordercount > 99:
        xpath_click(driver, nem.LISTCOUNTSELECT) # 셀렉트 선택
        xpath_click(driver, nem.LISTCOUNT500) # 주문갯수 선택
        time.sleep(10) # 정보불러올떄까지 기다리기
    xpath_click(driver, nem.ALLCHECK) # 상품 전체선택
    xpath_click(driver, nem.SENDCHECK) # 상품 최종단계선택
    alertaccept(driver) #알람 확인
    time.sleep(10)
    alertaccept(driver) #알람 확인
    xpath_click(driver, nem.ORDERCHECKPAGE) # 주문 페이지로 이동
    if newordercount > 99:
        xpath_click(driver, nem.LISTCOUNTSELECT) # 셀렉트 선택
        xpath_click(driver, nem.LISTCOUNT500) # 주문갯수 선택
        time.sleep(10) # 정보불러올떄까지 기다리기
    xpath_click(driver, nem.EXCELDOWNLOAD) # 엑셀 버튼클릭
    try:
        xpath_send_keys(driver, nem.DOWNLOADINPUT1, nem.EXCELPASSWORD)
        xpath_send_keys(driver, nem.DOWNLOADINPUT2, nem.EXCELPASSWORD)
        xpath_click(driver, nem.EXCELPASSBTN)
    except:
        xpath_enter(driver, nem.EXCELAPPLY) # 엑셀다운로드버튼 클릭

messagebox.showinfo("인터넷작업","크롤링작업을 마쳤습니다.\n버전 v[0.5.5]")
time.sleep(5)
driver.close()
