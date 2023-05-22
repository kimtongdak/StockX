# Rev0 ( 2022-01-19 ) : Rev0 완성
# Rev1
#   1. DataFrame Header 오타수정 (완료)
#   2. time.sleep(-) 거의 다 wait until 로 수정 (완료)
#   3. StockX 위치 및 단위변경 selector -> full xpath로 변경 / 갑자기 바뀜 (완료)
#   4. Kream은 더이상 수정필요x 줜나빨라짐 ( 완료 )
#   5. DataFrame 1행 고정필요 (panel freeze = (1,1) ) ( 완료 )
#   6. Stockx 검색속도 향상방법 검토 ( 몰라 더이상 )

# Rev2 ( 2022-02-09) 
#   1. Xpath 위치가 자꾸 변경됨, 문제 있을때마다 full xpath로 수정필요
#   2. chrome 32/64 bit에 따라 설치경로 찾아가기 추가


#######################################################################################################################
from os import name, replace
from re import I
from this import d
from turtle import end_fill
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wait 
from selenium.webdriver.common.by import By 
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from tkinter import *
from tkinter import filedialog


import chromedriver_autoinstaller # 크롬 버전 업그레이드 되도 자동으로 설치
import subprocess
import time
import pandas
import sys
import os
#######################################################################################################################

# 파일경로에서 불러와서 존재여부 확인 후 프로그램 종료여부 결정
file_name = filedialog.askopenfile(title = '거위밥을 선택하세요', filetypes = (('xlsx files', '*.xlsx'),('xls files', '*.xls'),('csv files', '*.csv')))
if file_name == 'None':
    sys.exit()
else:
    file_path = file_name.name

#######################################################################################################################
current_time = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')

if os.path.isdir(r'C:\Program Files\Google'):
    subprocess.Popen(r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"') # 디버거 크롬 구동
elif os.path.isdir(r'C:\Program Files (x86)\Google'):
    subprocess.Popen(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"') # 디버거 크롬 구동

option = Options()
option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
try:
    browser = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
except:
    chromedriver_autoinstaller.install(True)
    browser = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
browser.implicitly_wait(10)

# 크롬 설치위치 확인 후 디버거모드로 접속 ( 봇탐지 우회 )
# 추가적으로 봇 확인하는건 보통 자바스크립트가 실행되었는지 확인 ( 사람이 할 경우에는 키보드나 마우스 사용, 자바스크립트 안씀)
#######################################################################################################################

#######################################################################################################################
# 0. 네이버 환율 확인
browser.maximize_window() # 창 최대화
browser.get("https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=%ED%99%98%EC%9C%A8")
Exchange_rate = browser.find_element_by_xpath('//*[@id="_cs_foreigninfo"]/div/div[2]/div/div[1]/div[1]/div[2]/div[3]/div/div[2]/div[2]/span[2]').text
Exchange_rate = float(int(str(Exchange_rate[0:5]).replace(',',''))) #1,xxx ( float 타입으로 바꿔줘됨, float으로 바꿀려면 ','제거필요 / replace할려면 str도 필요 )


#######################################################################################################################

# 1. 공통 변수 선언
stockx_fee = 1.03
card_fee = 1.01
delivery_charge = 5.58

stockx_pandas = pandas.read_excel(f'{file_path}', usecols='B') # B : Stockx
kream_pandas = pandas.read_excel(f'{file_path}', usecols='C') # C : Kream

stockx_list = list()
for i in range(len(stockx_pandas)):
    if str(stockx_pandas.iat[i,0]) == 'nan':
        continue
    else:
        stockx_list.append(str(stockx_pandas.iat[i,0]))

kream_list = list()
for i in range(len(kream_pandas)):
    if str(kream_pandas.iat[i,0]) == 'nan':
        continue
    else:
        kream_list.append(str(kream_pandas.iat[i,0]))

#######################################################################################################################
# 2. StockX 접속
browser.get("https://stockx.com/ko-kr")

## 2-1. scroll 제일 아래 
prev_height = browser.execute_script("return document.body.scrollHeight")
while True:
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(2) 
    current_height = browser.execute_script("return document.body.scrollHeight") # 현재 문서 높이를 가져와서 저장
    if current_height == prev_height:
        break
    prev_height = current_height

## 2-2. 위치 및 단위변경
browser.find_element_by_xpath('//*[@id="site-footer"]/div/div[2]/div/div[1]/button').click() # 지역 / 언어 / 화폐 옵션창 클릭
browser.find_element_by_id("region-select").click()
browser.find_element_by_xpath('/html/body/div[4]/div/div[4]/div/section/div/div[1]/select/option[1]').click() # Region : 미국
browser.find_element_by_id("currency-select").click()
browser.find_element_by_xpath('//*[@id="currency-select"]/option[11]').click() # Currency : $ USD 
browser.find_element_by_xpath('/html/body/div[4]/div/div[4]/div/section/footer/div/button[2]').click() # 변경사항저장 ( xpath사용시 자꾸 변경되서 full xpath로 사용 )
time.sleep(3)

## 2-3. 제품 리스트 및 배열 초기화
all_item_list = [[]*1 for _ in range(len(stockx_list)*6)] # 2차원 리스트 선언 및 초기화
index = 0 # index 0부터 시작해서 사이즈 하나씩 바뀔때마다 +1

for a in range(len(stockx_list)):
    current_item = stockx_list[a]
    research = browser.find_element_by_id('site-search')
    research.clear() # 검색창 초기화
    research.send_keys(current_item)
    research.send_keys(Keys.ENTER)
    wait(browser, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME,'css-xzkzsa')))

## 2-4. 검색 후 검색어와 동일한 이미지로 접속
    for b in range(1,10): # 10개 이상은 안나오겠찡...
        research_xpath = f'/html/body/div[1]/div[1]/main/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[{b}]/div/a/div/div[2]/p'
        research = browser.find_element_by_xpath(research_xpath)
        if research.get_attribute('innerText') == f'{current_item}':
            research.click()
            break
        else:
            pass
        
## 2-5. 사이즈 갯수 및 리스트 확인
    browser.find_element_by_xpath('//*[@id="menu-button-pdp-size-selector"]').click() # 사이즈 리스트 : 이거를 눌러줘야 아래 class가 나옴
    time.sleep(3)
    all_size_xpath = '//*[@id="menu-list-pdp-size-selector"]/div[2]/div'
    size_str_type = browser.find_element_by_xpath(all_size_xpath) # 사이즈 목록
    size_lst_type = str.split(size_str_type.text) 

    size_list = list()
    for e in range(len(size_lst_type)): # 0,1,2,3,4,5,6,7 ( 0 ~ range-1 까지)
        if e == 0 : # 0
            size_list.insert(e,size_lst_type[e]) # append는 마지막에 추가 / insert는 설정 index에 추가
        elif e%2 == 1: # 홀수
            pass
        elif e%2 == 0: # 짝수
            size_list.insert(e-1,size_lst_type[e])

    for f in range(len(size_list)): # size 갯수만큼 for문 돌려서 사이즈 / 판매 / 구매 가격순으로 list
        size_xpath = f'{all_size_xpath}/button[{f+1}]' # button[1~4]
        browser.find_element_by_xpath(size_xpath).click() # button[1] 클릭
        selling_us_price = browser.find_element_by_xpath('//*[@id="main-content"]/div/section[1]/div[3]/div[2]/div[1]/div[2]/div[1]/div/dl/dd').text
        if selling_us_price == '--': # 구매 혹은 판매가가 없을경우 '--'로 나옴 0으로 처리
            selling_us_price = float(0)
        else:
            selling_price = str(selling_us_price)[3:] # US$ 제거
        buying_us_price = browser.find_element_by_xpath('//*[@id="main-content"]/div/section[1]/div[3]/div[2]/div[1]/div[2]/div[2]/div/dl/dd').text
        if buying_us_price == '--':
            buying_us_prece = float(0)
        else:
            buying_price = str(buying_us_price)[3:] # US$ 제거
        
        print(buying_price)
        print(selling_price)

        all_item_list[index].append(current_item) # index0 : 제품명
        all_item_list[index].append(size_list[f]) # index1 : 사이즈
        all_item_list[index].append(float(int(selling_price.replace(',','')))) # index2 : 즉시판매가
        all_item_list[index].append(float(int(buying_price.replace(',','')))) # index3 : 즉시구매가
        all_item_list[index].append(int((all_item_list[index][2]*stockx_fee+delivery_charge)*card_fee*Exchange_rate)) # index4 : StockX 최종구매가 ( 판매가 * 수수료 + 배송비) * 환율 * 카드수수료
        index += 1 # 마지막에 넣어줘야 0부터 시작
        browser.find_element_by_xpath('//*[@id="menu-button-pdp-size-selector"]').click() # 사이즈 리스트 클릭

    for i in range(5):
        all_item_list[index].append('') # 아이템 하나 끝나고 공란 ( stockx랑 , kream이랑 index에 맞게 넣어야됨 )
    index += 1 
#######################################################################################################################

#######################################################################################################################
# 3. Kream 접속 ( 최초 1회 접속 후 로그인 필요 )

browser.get("https://kream.co.kr")

## 3-1. 제품 리스트 및 배열 초기화
index = 0
browser.refresh() # Kream 접속 후 새로고침 한번 해줘야 검색창에서 검색가능
browser.implicitly_wait(10) # 새로고침 완료될때까지 대기

# Stockx랑 Kream이랑 제품명 표기 다름
# ex) stock : supreme the north face summit series outer tape seam jacket black
# ex) Kream : Supreme x The North Face Summit Series Outer Tape Seam Jacket Black - 21SS
# 제품명 모아서 stock랑 Kream이랑 index만 맞춰주면 될 듯

## 3-2. 새로고침 및 검색창 활성화
for a in range(len(kream_list)):
    current_item = kream_list[a]
    browser.find_element_by_xpath('/html/body/div/div/div/div[1]/div[2]/div/div[1]/div/a').click() # 검색 버튼 클릭
    research = browser.find_element_by_xpath('/html/body/div/div/div/div[1]/div[2]/div/div[4]/div/div/div[2]/div[1]/div/div/div/input') # 검색창 클릭
    research.clear()
    research.send_keys(current_item)
    research.send_keys(Keys.ENTER)
    wait(browser, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME,'search_result_list'))) # 전체 아이템을 찾기위해서 부모 class 찾아내야됨!?

## 3-3. 제품명과 동일한놈으로 클릭
    research = browser.find_elements_by_class_name('name')
    for b in range(len(research)): # 'name'으로 검색해서 current_item과 동일한 품목은 바로 클릭
        if research[b].text == current_item:
            research[b].click()
            break
        else:
            continue

## 3-4. 사이즈 갯수 확인
    browser.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[1]/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/a').click() # 사이즈 버튼 클릭
    size_list = list()
    size = browser.find_elements_by_class_name('size')  # 사이즈 최초1회 클릭되면 뭔가 변경됨 

    for c in range(len(size)-2): # 모든 사이즈 2개 + 사이즈 
        size_list.append(size[c+2].text) 
    size_burton_xpath = '//*[@id="__layout"]/div/div[2]/div[5]/div/div/div[2]/div/ul/'

## 3-5. 사이즈별로 가격 확인
    for d in range(len(size_list)):
        browser.find_element_by_xpath(f'{size_burton_xpath}/li[{d+2}]').click()
        time.sleep(0.5) # 사이즈 누르고 이거 없으면 가격 지마음대로 나옴
        price = browser.find_elements_by_class_name('num')
        for e in range(3):
            if price[e].text == '-': # 크림에서 금액이 없을경우 '-'로 나옴
                all_item_list[index].append(int(float(0)))
            else:
                all_item_list[index].append(int(str(price[e].text).replace(',',''))) # index5,6,7 : 최근 거래가 / 즉시 구매가 / 즉시 판매가 추가 
        
        all_item_list[index].append(all_item_list[index][5]-all_item_list[index][4]) # index8 : Stocxk 최종구매가 - 크림 최근 거래
        all_item_list[index].append(all_item_list[index][7]-all_item_list[index][4]) # index9 : Stocxk 최종구매가 - 크림 즉시 판매
       
        index += 1
        
        browser.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[1]/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/a').click() # 사이즈 버튼 클릭

    for i in range(5):
        all_item_list[index].append('') # 아이템 하나 끝나고 공란 ( stockx랑 , kream이랑 index에 맞게 넣어야됨 )
    
    index += 1
    browser.find_element_by_xpath('//*[@id="__layout"]/div/div[2]/div[5]/div/div/a').click() # X버튼 클릭

current_time = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
pandas.DataFrame.columns = ['품명','Size','StockX 구매(usd)','StockX 판매(usd)','StockX 구매(won)','크림 최근거래(won)','크림 구매(won)','크림 판매(won)','최근거래 차액(G-F)','즉시판매 차액(I-F)']
pandas.DataFrame(all_item_list).to_excel(f'The Goose with the Golden Eggs_{current_time}.xlsx',index_label=f'{Exchange_rate}',freeze_panes=(1,1))




