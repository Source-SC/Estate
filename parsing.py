import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import argparse
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime

import openpyxl#패키지 불러오기

from selenium.webdriver.common.action_chains import ActionChains

# request에러 잡기위한 import
import requests

scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]

json_file_name = 'second_client.json'

credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)


spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1_tEv2GDvD1_EYbP3xTlnhN19Pmmigzdltzgfwea_Gr8/edit#gid=515257963'

# # 스프레스시트 문서 가져오기 
sh = gc.open_by_url(spreadsheet_url)

# # 시트 선택하기
worksheet = sh.get_worksheet(5)

# worksheet.update_acell('A2',"a")

# sh = gc.open('2020estate')



upload_date = "-"
address = "-"
price = "-"
size = "-"
floor = "-"
in_date = "-"
usage = "-"
ok_date = "-"
detail = "-"
serial_num = "-"

def reset():
    global upload_date
    global address
    global price
    global size
    global floor
    global in_date
    global usage
    global ok_date
    global detail
    global serial_num

    upload_date = "-"
    price = "-"
    size = "-"
    floor = "-"
    in_date = "-"
    usage = "-"
    ok_date = "-"
    detail = "-"
    serial_num = "-"

def transform_date(date):
    output = ""
    for i in range(len(date)):
        if(date[i] == "."):
            output = output + "-"
        else:
            output = output + date[i]
    if(output[len(output)-1] == "-"):
        output = output[:-1]
    return output

def matching_title_contents(title, contents):
    if(title == "대지/연면적" or title == "대지면적"):
        global size
        size = contents
        return 1
    elif(title == "지상층/지하층"):
        global floor
        floor = contents
        return 1
    elif(title == "입주가능일"):
        global in_date
        in_date = contents 
        return 1
    elif(title == "현재용도"):
        global usage
        usage = contents 
        return 1
    elif(title == "사용승인일"):
        global ok_date
        ok_date = contents 
        return 1 
    elif(title == "매물설명"):
        global detail
        detail = contents 
        return 1
    elif(title == "매물번호"):
        global serial_num
        serial_num = contents 
        return 1
    else:
        return 0

def switch(a):
    if(a == 1):
        return 2
    else:
        return 1	

def send_excel():

    total = int(worksheet.acell('W1').value)

    global upload_date
    global address
    global price
    global size
    global floor
    global in_date
    global usage
    global ok_date
    global detail
    global serial_num

    
    tod = " | " + str(datetime.today().month)+"/"+str(datetime.today().day)
    upload_date = str(transform_date(upload_date)) + tod



    link = 'https://new.land.naver.com/offices?ms=37.4101,126.7378,16&a=TJ:GM&e=RETAIL&articleNo=' + serial_num

    input_list = [[]]
    input_list[0].append(str(total+1))
    input_list[0].append(str(serial_num))
    input_list[0].append(str(address))
    input_list[0].append(upload_date)
    input_list[0].append(str(price))
    input_list[0].append("")
    input_list[0].append(str(size))
    input_list[0].append(str(floor))
    input_list[0].append(str(transform_date(in_date)))
    input_list[0].append(str(usage))
    if(len(str(ok_date)) < 3):
        ok_date = "--"
    input_list[0].append(str(transform_date(ok_date)))
    input_list[0].append(str(detail))
    input_list[0].append(str(link))
    input_list[0].append('-')

    # requests.exceptions.ReadTimeout: HTTPSConnectionPool(host='sheets.googleapis.com', port=443): Read timed out. (read timeout=120)
    # 해당 에러해결 위한 try문
    try:
        worksheet.update(('A' + str(total+2) + ':N' + str(total+2)),input_list)
        worksheet.update('W1',str(total+1))
        # requests.post(url, headers, timeout=10)
    except requests.exceptions.Timeout:
        print("Timeout occurred")

    # worksheet.update_acell('A'+str(total+2),str(total+1))
    # worksheet.update_acell('B'+str(total+2),str(serial_num))
    # worksheet.update_acell('C'+str(total+2),str(transform_date(upload_date)))
    # worksheet.update_acell('D'+str(total+2),str(price))
    # worksheet.update_acell('E'+str(total+2),str(size))
    # worksheet.update_acell('F'+str(total+2),str(floor))
    # worksheet.update_acell('G'+str(total+2),str(transform_date(in_date)))
    # worksheet.update_acell('H'+str(total+2),str(usage))
    # worksheet.update_acell('I'+str(total+2),str(transform_date(ok_date)))
    # worksheet.update_acell('J'+str(total+2),str(detail))
    # worksheet.update_acell('O2',str(total+1))
    




# dr.get('https://new.land.naver.com/offices?ms=37.4101,126.7378,16&a=TJ:GM&e=RETAIL&articleNo=2103040642')
#        'https://new.land.naver.com/offices?ms=37.3967,126.7001,16&a=TJ:GM&e=RETAIL&articleNo=2103350163'


# 고잔동 'https://new.land.naver.com/offices?ms=37.3967,126.7001,16&a=TJ:GM&e=RETAIL&articleNo=2103350163'
# 논현동 'https://new.land.naver.com/offices?ms=37.4101,126.7378,16&a=TJ:GM&e=RETAIL'
# 남촌동 'https://new.land.naver.com/offices?ms=37.4241,126.7127,16&a=TJ:GM&e=RETAIL'
print("------------------------------")
print("네이버 부동산 매물 검색 서비스 ")
print("-------------------------------")

for sel in range(0,3):
    success = 0
    exist = 0
    i = 1
    if(sel == 0): # 고잔동
        address = "고잔동"
        dr = webdriver.Chrome('./chromedriver-2') 
        dr.get('https://new.land.naver.com/offices?ms=37.3967,126.7001,16&a=TJ:GM:GJCG&b=A1&e=RETAIL')
        # dr.get('https://new.land.naver.com/offices?ms=37.3967,126.7001,16&a=TJ:GM&e=RETAIL')
    elif(sel == 1): # 논현동
        address = "논현동"
        dr = webdriver.Chrome('./chromedriver-2') 
        dr.get('https://new.land.naver.com/offices?ms=37.4101,126.7378,16&a=TJ:GM:GJCG&b=A1&e=RETAIL')
        # dr.get('https://new.land.naver.com/offices?ms=37.4101,126.7378,16&a=TJ:GM&e=RETAIL')
    else: # 남촌동
        address = "남촌동"
        dr = webdriver.Chrome('./chromedriver-2') 
        dr.get('https://new.land.naver.com/offices?ms=37.4241,126.7127,16&a=TJ:GM:GJCG&b=A1&e=RETAIL')
        # dr.get('https://new.land.naver.com/offices?ms=37.4241,126.7127,16&a=TJ:GM&e=RETAIL')

    exist_date = worksheet.col_values(2)

    dr.find_element_by_xpath("""//*[@id="list"]/div/div/div[1]/div/a[2]""").click()
    time.sleep(1)
    while(True):

        original = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i) + "]/div/a/div[1]/span"
        direct_naver = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i)+ "]/div/div[2]/a"
        reset()
        try:
            tmp = dr.find_element_by_xpath(original)
            time.sleep(0.2)
        except NoSuchElementException:
            print("none")
            break
        
        if(tmp.text == ""):
            print("빈칸때문에 한칸 뒤로 이동")
            original = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i) + "]/div/a/div[1]/span[2]"
            reset()
            try:
                tmp = dr.find_element_by_xpath(original)
                time.sleep(0.2)
            except NoSuchElementException:
                print("none")
                break
        print(i,"번째 탐색중 : ",tmp.text)
        if(tmp.text == "공장" or tmp.text == "창고" or tmp.text == "공장용지"):
            check = 1
            now = 1
            try:
                direct_link = dr.find_element_by_xpath(direct_naver)
                direct_link.click()
            except NoSuchElementException:
                tmp.click()
                pass
            # 2021.05.02
            # 가끔 대지/연면적 값이 안들어올때가 있었다. 확인해보니 매물정보중 5row부터는 정보를 가져오는것을 확인하였고
            # 해당 매물을 불러오기 이전에 값을 불러와서 에러가 나서 스킵이 되었던것같다.
            # 이같은 상황이 오늘 갑자기 생겼고 이전과 달라진 점은 매물중에 네이버에서 바로보기가 생겼다는 점이다.
            # 그래서 아래와같이 0.6초의 대기시간을 주어서 해결되긴했다.
            time.sleep(0.6)
            price_str = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i) + "]/div/a/div[2]/span[2]"
            date_str = "//*[@id='listContents1']/div/div/div/div[ " + str(i) + "]/div/div[2]/span/em[2]"

            tmp_price = dr.find_element_by_xpath(price_str) 
            price = tmp_price.text
            tmp_upload_date = dr.find_element_by_xpath(date_str) 										
            upload_date = tmp_upload_date.text

            last_index = 0
            while(check == 1):
                for point in range(3):
                    if point == 0: # 매물정보가 한줄로 되어있을때
                        searching_title = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/th"
                        searching_contents = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/td"
                    elif point == 1: # 매물정보가 두줄로 되어있을때 그중 왼쪽
                        searching_title = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/th[1]"
                        searching_contents = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/td[1]"
                    else: # 매물정보가 두줄로 되어있을때 그중 오른쪽
                        searching_title = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/th[2]"
                        searching_contents = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/td[2]"
                
            
                    try:
                        now_title = dr.find_element_by_xpath(searching_title) 
                        now_contents = dr.find_element_by_xpath(searching_contents)
                        matching_title_contents(now_title.text, now_contents.text)

                        if(now_title.text == "중개사"):
                            check = 0
                    except NoSuchElementException:
                        pass

                now = now + 1
                last_index = switch(last_index)

            

            # 이미 시트에 기록된 매물인지 확인
            if serial_num in exist_date: 
                exist = exist + 1
                # print('시트에 값이 있으므로 추가하지않음.') 
            else: 
                send_excel()
                success = success + 1
                print('리스트에 값이 없으므로 새로 추가. 매물번호 : ', serial_num)



        # print("-----------------------------------")
        # print("업로드날짜 : ",upload_date)
        # print("매매금액 : ",price)
        # print("대지연면적 : ",size)
        # print("지상층 / 지하층 : ",floor)
        # print("입주가능일 : ",in_date)
        # print("현재용도 : ",usage)
        # print("사용승인일 : ",ok_date)
        # print("매물설명 : ",detail)
        # print("-----------------------------------")

        time.sleep(0.2)
        i = i + 1

        actions = ActionChains(dr)
        actions.move_to_element(tmp).perform()
        dr.execute_script("arguments[0].scrollIntoView();", tmp)

    dr.close()
    print("-------------------------")
    if(sel == 0):
        print("고잔동")
    elif(sel == 1):
        print("논현동")
    else:
        print("남촌동")
    print("입력완료 매물 : ", success," 개")
    print("이미입력되어 있는 매물 : ",exist," 개")
    print("-------------------------")
    


# //*[@id="ct"]/div[2]/div[2]/div/div[2]/div[1]/div[1]/div[1]/span[2]/em[2]
# 업로드날짜				//*[@id="ct"]/div[2]/div[2]/div/div[2]/div[1]/div[1]/div[1]/span[1]/em[2]
# 매매금액				//*[@id="ct"]/div[2]/div[2]/div/div[2]/div[1]/div[1]/div[3]/span[2]
# 대지연면적				//*[@id="detailContents1"]/div[1]/table/tbody/tr[2]/td
# 지상흥/지하층			//*[@id="detailContents1"]/div[1]/table/tbody/tr[4]/td[1]

# 입주가능일				//*[@id="detailContents1"]/div[1]/table/tbody/tr[5]/td[2]
# 현재용도				//*[@id="detailContents1"]/div[1]/table/tbody/tr[7]/td[2]
# 사용승인일				//*[@id="detailContents1"]/div[1]/table/tbody/tr[12]/td[1]
# 매물설명				//*[@id="detailContents1"]/div[1]/table/tbody/tr[13]/td/span
# //*[@id="detailContents1"]/div[1]/table/tbody/tr[9]/td/span/pre

    # sub_screen.click()


    # tmp2.click()

    # tmp2.send_keys(Keys.ARROW_DOWN)

    # tmp2.send_keys(Keys.PAGE_DOWN)

    # tmp2.send_keys(Keys.PAGE_DOWN)
    # tmp2.send_keys(Keys.PAGE_DOWN)
    


    # //*[@id="listContents1"]/div/div/div[1]/div[1]/div/a/div[2]/span[2]
    # //*[@id="listContents1"]/div/div/div[1]/div[2]/div/a/div[2]/span[2]
    # //*[@id="listContents1"]/div/div/div[1]/div[3]/div/a/div[2]/span[2]