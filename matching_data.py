# 주요 공지 사항
# 현재 시트에서 행 순서를 바꿔버리면 예상매물의 이름에서는 상관없지만 링크가 변경되는 문제가 있다.

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import argparse
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

def parsing_size(size_list):
    final_list = []
    tmp = []
    final_list.append(tmp)
    for i in range(1,len(size_list)):
        tmp = []
        switch = 1
        front = ""
        second = ""
        print("now",size_list[i])
        if(size_list[i] == "-"):
            tmp.append(0)
            tmp.append(0)
            final_list.append(tmp)
            # final_list[i].append(0)
            # final_list[i].append(0)
        else:
            for j in range(len(size_list[i])):
                if(switch == 1):
                    if(size_list[i][j] == "/"):
                        switch = 2
                    else:
                        front = front + size_list[i][j]
                else:
                    second = second + size_list[i][j]
            if(switch == 1): # 대지/연면적 두값이 있지않고 한값만 있을경우 이부분을 통해 에러 해결 date. 21.02.26
                tmp.append(float(front[:-1]))
                tmp.append(0)
                final_list.append(tmp)
            else:
                tmp.append(float(front[:-1]))
                tmp.append(float(second[:-1]))
                final_list.append(tmp)
            # final_list[i].append(float(front[:-1]))
            # final_list[i].append(float(second[:-1]))
    return final_list



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
worksheet4 = sh.get_worksheet(4)
worksheet5 = sh.get_worksheet(5)
	
total_list = worksheet5.col_values(23)
total = int(total_list[0])

print("-----------------------------")
print("네이버부동산 정보 - 사용승인일자 매칭")
print("현재 네이버부동산 셀 개수 : ", total)
print("매칭을 원하는 셀 번호 시작점과 끝점을 입력하세요. 모든셀을 선택하려면 0을 입력하세요")
start = int(input("시작점 : "))
if(start == 0):
    start = 1
    end = total
else:
    end = int(input("끝점 : "))
start = start + 1
end = end + 2


# 위치를 지정하여 행렬 형태로 데이터 가져오기

ok_date_list = worksheet5.col_values(11)
size_list = worksheet5.col_values(7)

date_list = worksheet4.col_values(19)
ground_list = worksheet4.col_values(11)
buliding_list = worksheet4.col_values(12)
name_list =  worksheet4.col_values(2)
# print(values_list)

parse_size_list = parsing_size(size_list)
print(parse_size_list)
index = 0
answer = []
print(ground_list)
for now in range(start,end):
    sub_list = []
    # print(now-1,"번째 셀 검색중...")
    ok_date = ok_date_list[now-1]
    size = parse_size_list[now-1][0]
    if(len(str(ok_date)) != 1):
        for i in range(len(date_list)):
            # print("ii = ",i)
            if(date_list[i] == ok_date): # 날짜로 조회
                cell = 'S' + str(i+1)
                cell = "https://docs.google.com/spreadsheets/d/1_tEv2GDvD1_EYbP3xTlnhN19Pmmigzdltzgfwea_Gr8/edit#gid=515257963&range=" + cell
                if(len(str(name_list[i])) == 0): # 상호가 여러개가 있는 곳이면 맨 위에만 상호가 등록되어있으므로 찾을때까지 위로 올라감
                    # print("복수상호 발견")
                    up = 1
                    while(len(str(name_list[i-up])) == 0):
                        up = up+1
                    sub_list.append(name_list[i-up])
                else:
                    sub_list.append(name_list[i])
                
                sub_list.append(cell)
            # 면적 정보가 존재할때
            # print(size)
            # print(ground_list[i])
    if(size > 0):
        mini = size/10000 * 9999
        maxi = size/10000 * 10001 
        for i in range(1,len(ground_list)):
            # 왜 그런지 모르겠는데 대지연면적 확인쪽에서 위의 len(date_list)로 되어있어서 인덱스 오류가 발생했다.
            # 아마도 존재하는 대지연면적와 사용승인일 날짜의 데이터수가 같다고 생각했었는데 시트를 수정하면서 그값이 달라진것 같다.

            # print(len(date_list))
            # print("now i = ",i)
            # print("str : ",str(ground_list[i]))
            if(len(str(ground_list[i])) > 2):
                if(float(ground_list[i]) >= mini and float(ground_list[i]) <= maxi):
                    cell = 'K' + str(i+1)
                    cell = "https://docs.google.com/spreadsheets/d/1_tEv2GDvD1_EYbP3xTlnhN19Pmmigzdltzgfwea_Gr8/edit#gid=515257963&range=" + cell
                    if(len(str(name_list[i])) == 0): # 상호가 여러개가 있는 곳이면 맨 위에만 상호가 등록되어있으므로 찾을때까지 위로 올라감
                        # print("복수상호 발견")
                        up = 1
                        while(len(str(name_list[i-up])) == 0):
                            up = up+1
                        sub_list.append(name_list[i-up])
                    else:
                        sub_list.append(name_list[i])
                    
                    sub_list.append(cell)
                    

        # 65 - A 76 - L
        if(len(sub_list) == 0): # 검색실패시
            sub_list.append("X")
        else:
            sub_list.append("-")

        # if(len(ok_date) == 0):
        #     print("null")
        # else:
        #     print(ok_date)
    else:
        sub_list.append("-")
        print("날짜가 비어있습니다")
    answer.append(sub_list)
    index = index + 1
input_cell = chr(78) + str(start) + ":BF" + str(end-1)
print(input_cell)
print(answer)
worksheet5.update(input_cell,answer)
    
















# while(True):
# 	original = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i) + "]/div/a/div[2]/span[2]"
# 	reset()
# 	try:
# 		tmp = dr.find_element_by_xpath(original)
# 		time.sleep(0.2)
# 	except NoSuchElementException:
# 		break
# 	print(i,"번째 탐색중 : ",tmp.text)

# 	check = 1
# 	now = 1
# 	tmp.click()
# 	price_str = "//*[@id='listContents1']/div/div/div[1]/div[" + str(i) + "]/div/a/div[2]/span[2]"
# 	date_str = "//*[@id='listContents1']/div/div/div/div[ " + str(i) + "]/div/div[2]/span/em[2]"

# 	tmp_price = dr.find_element_by_xpath(price_str) 
# 	price = tmp_price.text
# 	tmp_upload_date = dr.find_element_by_xpath(date_str) 										
# 	upload_date = tmp_upload_date.text

# 	last_index = 0
# 	while(check == 1):

# 		# //*[@id="detailContents1"]/div[1]/table/tbody/tr[4]/th[1]
# 		# //*[@id="detailContents1"]/div[1]/table/tbody/tr[4]/td[1]
# 		# //*[@id="detailContents1"]/div[1]/table/tbody/tr[4]/th[2]
# 		# //*[@id="detailContents1"]/div[1]/table/tbody/tr[4]/td[2]

# 		# print(now, " - ", last_index)
# 		# time.sleep(0.2)
# 		searching_title = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/th[ " + str(last_index) + "]"
# 		searching_contents = "//*[@id='detailContents1']/div[1]/table/tbody/tr[" + str(now) + "]/td[ " + str(last_index) + "]"

# 		try:
# 			now_title = dr.find_element_by_xpath(searching_title) 
# 			now_contents = dr.find_element_by_xpath(searching_contents)
# 			print(now_title.text," : ",now_contents.text)
# 			matching_title_contents(now_title.text, now_contents.text)
# 			if(now_title.text == "중개사"):
# 				check = 0
# 		except NoSuchElementException:
# 			print("out")
				
# 		if(last_index == 2):
# 			now = now + 1
# 		last_index = switch(last_index)

# 	send_excel()

# 	print("-----------------------------------")
# 	print("업로드날짜 : ",upload_date)
# 	print("매매금액 : ",price)
# 	print("대지연면적 : ",size)
# 	print("지상층 / 지하층 : ",floor)
# 	print("입주가능일 : ",in_date)
# 	print("현재용도 : ",usage)
# 	print("사용승인일 : ",ok_date)
# 	print("매물설명 : ",detail)
# 	print("-----------------------------------")

# 	time.sleep(0.2)
# 	i = i + 1

# 	actions = ActionChains(dr)
# 	actions.move_to_element(tmp).perform()
# 	dr.execute_script("arguments[0].scrollIntoView();", tmp)

# print("end")


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