import pandas as pd
import requests
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 

# 핵심 기입 정보
year = "2022" # 년
month = "01" #월
date  = "02" #일
날짜 = year+month+date
read = '22.01.02 Covid-19_.xlsx' #읽을 엑셀 파일
출력 = '주소 확인('+날짜+').xlsx'#출력 될것
#파이썬 3.9.7
#사전작업
#사전에 열어둔 크롬 페이지로 설정
#실행 커맨드 
#cd C:\Program Files\Google\Chrome\Application
#chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeTEST"
#https://covid19.kdca.go.kr/

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "chromedriver.exe" # Your Chrome Driver path
driver = webdriver.Chrome(chrome_driver, options=chrome_options)

driver.implicitly_wait(3)
#프레임 변경
driver.switch_to.frame('base')
driver.switch_to.frame('ifrm')

#사전 조건 
# https://covid19.kdca.go.kr/ 에 접속하여 공인인증서 인증후
# 코로나19 정보관리시스템 > 환자감시 > 감염병웹신고(병의원) > 신고내역 관리
# > 신고서 작성 창까지 들어온다.
# 주민번호 앞자리가 0일시 빠진다
#핵심 함수
def 자동입력(이름, 주소1, 주소2, 주민번호1, 주민번호2, 전화번호1, 전화번호2, 전화번호3):
        #이름
        driver.execute_script("document.getElementById(\"ptxtPatntNm\").value=\""+이름+"\"")

        #주민번호 ptxtPatntIhidnum1 ptxtPatntIhidnum2
        driver.execute_script("document.getElementById(\"ptxtPatntIhidnum1\").value=\""+str(주민번호1)+"\"")
        driver.execute_script("document.getElementById(\"ptxtPatntIhidnum2\").value=\""+str(주민번호2)+"\"")
        # ptxtPatntIhidnum2의 앞자리가 5, 6, 7, 8 에 경우 외국인 전용 구간 필요


        #전화번호 ptxtPatntMbtlnum1 ptxtPatntMbtlnum2 ptxtPatntMbtlnum3
        driver.execute_script("document.getElementById(\"ptxtPatntMbtlnum1\").value=\""+str(전화번호1)+"\"")
        driver.execute_script("document.getElementById(\"ptxtPatntMbtlnum2\").value=\""+str(전화번호2)+"\"")
        driver.execute_script("document.getElementById(\"ptxtPatntMbtlnum3\").value=\""+str(전화번호3)+"\"")

        #주소(우편번호가 없어도 되는가? 팝업 제어가 되는가?)
        # ptxtPatntRnZip 우편 번호
        # ptxtPatntRdnmadr 도로명 주소
        # ptxtPatntRdnmadrDtl 추가 주소
        # driver.execute_script("document.getElementById(\"ptxtPatntRnZip\").value=\"010\"")
        # driver.execute_script("document.getElementById(\"ptxtPatntRdnmadr\").value=\"010\"")
        # driver.execute_script("document.getElementById(\"ptxtPatntRdnmadrDtl\").value=\"010\"")

        #juso 팝업 api 제어 
        driver.execute_script("document.getElementById(\"pbtnSearchRdnmadr\").click()")
        driver.implicitly_wait(5)
        driver.switch_to.window(driver.window_handles[1])
        #주소 입력
        driver.find_element_by_class_name('popSearchInput').send_keys(주소1)
        driver.execute_script("document.getElementById('keyword').value =\""+ 주소1 +"\"")
        #검색 클릭
        driver.implicitly_wait(5)
        driver.execute_script("javascript:$('#raFirstSortNone').prop('checked',true); searchUrlJuso();")
        #driver.find_element_by_xpath("/html/body/form[2]/div/div[1]/div/div[1]/fieldset/span/input[2]").click()
        driver.implicitly_wait(5)

        #검색 결과 클릭
        driver.execute_script("javascript:setMaping('1')")


        driver.implicitly_wait(5)
        #rtAddrDetail 추가 주소 입력
        driver.execute_script("document.getElementById('rtAddrDetail').value =\""+ 주소2 +"\"")
        #주소 입력 및 팝업 닫기
        end = driver.find_element_by_class_name("btn-bl")
        driver.execute_script("arguments[0].click();", end)
        driver.implicitly_wait(5)
        #원래 페이지로 돌아가기
        driver.switch_to.window(driver.window_handles[0])
        #프레임 변경
        driver.switch_to.frame('base')
        driver.switch_to.frame('ifrm')


        #기본 기입 사항
        # pcmbPatntOccpCd 직업 기타 : Z
        driver.execute_script("document.getElementById(\"pcmbPatntOccpCd\").value=\"Z\"")

        # ptxtNwkndIcdStarmySymptms 검사 이유 입력
        driver.execute_script("document.getElementById(\"ptxtNwkndIcdStarmySymptms\").value=\"코로나 검사\"")

        # 날짜 입력 ptxtAtfssDe1 ptxtAtfssDe2 ptxtAtfssDe3 
        #  ptxtDgnssDe1 ptxtDgnssDe2 ptxtDgnssDe3

        driver.execute_script("document.getElementById(\"ptxtAtfssDe1\").value=\""+ year +"\"")
        driver.execute_script("document.getElementById(\"ptxtAtfssDe2\").value=\""+ month +"\"")
        driver.execute_script("document.getElementById(\"ptxtAtfssDe3\").value=\""+ date +"\"")

        driver.execute_script("document.getElementById(\"ptxtDgnssDe1\").value=\""+ year +"\"")
        driver.execute_script("document.getElementById(\"ptxtDgnssDe2\").value=\""+ month +"\"")
        driver.execute_script("document.getElementById(\"ptxtDgnssDe3\").value=\""+ date +"\"")


        # 체크리스트
        # prdoDsndgnssInspctResultTyCd3 검사 진행중
        # prdoDsndgnssInspctResultTyCd2 음성
        # prdoPatntClCd2 의사환자
        driver.execute_script("document.getElementById(\"prdoDsndgnssInspctResultTyCd2\").click()")#음성
        driver.execute_script("document.getElementById(\"prdoPatntClCd2\").click()")#의사환자


        driver.execute_script("document.getElementById(\"ptxtRmInfo\").value=\"조사대상 유증상자 3 : 국내 집단발생 관련 유증상자\"")

        # 의사 이름
        driver.execute_script("document.getElementById(\"ptxtSttemntDoctrNm\").value=\"노리히사요꼬\"")


        driver.execute_script("document.getElementById(\"pchkNA0012ErrCheck\").click()")


        #외국인 전용 시퀸스
        # pchkFrgnrAt
        # driver.execute_script("document.getElementById(\"pchkFrgnrAt\").click()")
        # driver.execute_script("document.getElementById(\"pchkErrCheck\").click()")
        if(주민번호2>4999999) :
                driver.execute_script("document.getElementById(\"pchkFrgnrAt\").click()")
                driver.execute_script("document.getElementById(\"pchkErrCheck\").click()")   
                

        #신고 버튼 눌렸을때 반응
        # btn-blue btn-check 
        driver.execute_script("document.getElementById(\"pbtnCreateReport\").click()")
        time.sleep(3)
        try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                driver.implicitly_wait(5)
                # 취소하기(닫기)
                #alert.dismiss()
                
                # 확인하기
                alert.accept()
                print("신고완료")
        except:
                print("no alert")



#엑셀을 읽어서 추가 자료들을 정제
#전화번호 -를 기준으로 slice
#주소 카카오 api를 사용하여 추가 정제
#현재 키는 590488a94a19d10b3e9a6e876738dc4e 이며 바꿔야될수도?
#추가 사항 juso api를 추가 적으로 사용하여 추가 정제
#현재 키는 devU01TX0FVVEgyMDIxMTIzMDE3MzM1ODExMjA4NDg= 이며 유효기간은 90일로 상당히 짧음
#만약 주소가 없다면 기존의 주소로 저장하고 넘겨버림

def 주소정제juso(주소1):
    url = 'https://www.juso.go.kr/addrlink/addrLinkApi.do' 
    params = {'keyword': 주소1,'confmKey': "devU01TX0FVVEgyMDIxMTIzMDE3MzM1ODExMjA4NDg=",'resultType' : "json"} 
    places = requests.get(url, params=params).json()
    #print(places)
    #print (places['results']['juso'][0]['roadAddrPart1'])
    result = ""
    try:
        result = places['results']['juso'][0]['roadAddrPart1']
    except IndexError as e : 
        if(result == ''):
            result = 주소1
    return result

def 주소정제kakao(주소1):
    url = 'https://dapi.kakao.com/v2/local/search/keyword.json' 
    params = {'query': 주소1,'page': 1} 
    headers = {"Authorization": "KakaoAK 590488a94a19d10b3e9a6e876738dc4e"}
    places = requests.get(url, params=params, headers=headers).json()['documents']
    y=""
    try:
        #print(places[0]["road_address_name"])
        y=places[0]["road_address_name"]
    except IndexError as n :
        y=주소1
    #만약 검색은 되는데 도로명 주소가 안나오는 경우가 존재
    if y == '':
        y=주소1
    return y


def 주소정제(주소1):
    re = 주소정제juso(주소1)
    if re != 주소1:
        return re
    else:
        re = 주소정제kakao(주소1)
    return re



# 진짜로 읽을 파일
# 파일 이름 변수로 변환 read
#df = pd.read_excel('주소수정완료(12.26).xlsx')
df = pd.read_excel(read)

x = df.values.tolist()
print(x)
#[][1] : 이름
#[][2] : 주민1
#[][3] : 주민2
#[][4] : 전화번호
#[][5] : 주소1
#[][6] : 주소2

정제전번= []
입력실패번호=[]
for n in range(0,len(x)):
    #주민1정제
    #제로필을 사용하여 정제
    x[n][2]= str(x[n][2]).zfill(6)
    #주소 정제
    #주소정제가 실패한경우 입력 실패로 간주하고 다음으로 넘어감
    주소1 =  x[n][5]
    정제주소 = 주소정제(주소1)
    if 정제주소==주소1:
        입력실패번호.append(n)
        print(str(n+1)+"/"+str(len(x))+" 주소 검색 불가")
        continue
    else:
        x[n][5] = 정제주소
    #print(x[n][5])

    #전화번호 정제
    정제전번 = str(x[n][4]).split('-')
    #print(정제전번)
    
    # 진짜로 입력 하는 함수
    try:
        자동입력(str(x[n][1]), str(x[n][5]), str(x[n][6]), x[n][2], x[n][3], 정제전번[0], 정제전번[1], 정제전번[2])
        print(str(n+1)+"/"+str(len(x))+" 5초 대기")
        time.sleep(5)
    except:
        입력실패번호.append(n)
        print(str(n+1)+"/"+str(len(x))+" 입력실패")


if 입력실패번호 == [] :
    print("전부 입력완료")
    quit()
    
   
#입력 이후 정제 실패한것들로 다시 엑셀 생성
y = []
for n in range(0,len(x)):
    for z in 입력실패번호:
        if n == z:
            y.append(x[n])

# 엑셀 파일 출력
# 출력 파일 

#df = pd.DataFrame(y,columns=['등록번호', '이름', '주민1','주민2','전화번호','주소1','주소2'])
#df.to_excel('주소수정필요(12.27).xlsx', sheet_name=str(날짜), index=False, header=True)
#df.to_excel(출력, index=False, header=True)
df = pd.DataFrame(y)
df.to_excel(출력, index=False,header=True)
print("주소수정 필요 파일 확인필요")

