from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import tkinter
import time
import tkinter.ttk
import pandas as pd
import requests
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
import chromedriver_autoinstaller
chromedriver_autoinstaller.install()


driver = webdriver.Chrome()
driver.get('https://covid19.kdca.go.kr/')


root= Tk()
root.title("코로나 자동 신고 0.6 제작자 : 윤호찬")
root.geometry("600x400")
root.resizable(False,False)

#받아야 할 값
#들어갈 엑셀 파일
#날짜파일
#단 날짜가 당일이면 안됨
#에러시 나올 파일이름

#입력할 엑셀 파일 지정
global inputfile
inputfile =""
global year
global month
global date
global savefile
#최종확인 카운터
global f_test
f_test= 0

def open():
    global inputfile
    root.filename = filedialog.askopenfilename(initialdir='', title='파일선택', filetypes=(
    ('xlsx files', '*.xlsx'), ('all files', '*.*')))
    Label(root, text=root.filename).grid(row=5,column=2) # 파일경로 view
    inputfile = str(root.filename)
    print(inputfile)

#날짜, 출력파일 이름 받기
def check():
    global year
    global month
    global date
    global savefile
    try:
        year = str(input1.get())
        month = int(input2.get())
        date = int(input3.get())
    except:
        messagebox.showinfo("안내","년 월 일이 이상합니다.")
        return
    if month >= 13  or month <= 0 or month=='':
        messagebox.showinfo("안내","년 월 일이 이상합니다.")
        return    
    if date >= 32  or date <= 0 or date=='':
        messagebox.showinfo("안내","년 월 일이 이상합니다.")
        return
    month = str(month).zfill(2)
    date = str(date).zfill(2)
    날짜 = year+month+date
    savefile = '신고 실패('+날짜+').xlsx'
    if savefile == "":
        messagebox.showinfo("안내","저장할 파일 이름을 작성하세요")
        return
    print(year+month+date)
    print(savefile)
    messagebox.showinfo("입력 확인", year+"년 "+month+"월 "+date+"일\n걸러질파일: "+savefile)

def final():
    global inputfile
    global year
    global month
    global date
    global savefile
    global f_test
    try:
        if inputfile == "":
            messagebox.showinfo("에러","엑셀 데이터를 다시 확인하세요") 
            return
        if month == "":
            messagebox.showinfo("에러","날짜를 확인해주세요")   
            return
    except:
        messagebox.showinfo("에러","날짜를 확인하세요")
        return
    messagebox.showinfo("최종확인",year+"년 "+month+"월 "+date+"일\n걸러질파일: "+savefile+"\n불러올 파일경로: "+inputfile)
    messagebox.showinfo("안내","커맨드로 열린 크롬에서 인증서 로그인을 한 후 신고창을 열어주세요")
    f_test = 1

P_var  = DoubleVar()
def main():
    global inputfile
    global year
    global month
    global date
    global savefile
    global f_test
    if f_test != 1:
        messagebox.showinfo("안내", "최종확인을 하고 오세요")
        return
    button3 = tkinter.DISABLED
    #실제 모든 작업이 집중될 장소
    
    #1.임의로 실행중인 크롬이랑 연결
    #사전에 열어둔 크롬 페이지로 설정
    #실행 커맨드 cmd 실행기준
    #cd C:\Program Files\Google\Chrome\Application
    #chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeTEST"
    #https://covid19.kdca.go.kr/
    global driver
    driver.implicitly_wait(3)
    #프레임 변경
    driver.switch_to.frame('base')
    driver.switch_to.frame('ifrm')
    
    #핵심 기믹
    # 진짜로 읽을 파일
    # 파일 이름 변수로 변환 read
    #df = pd.read_excel('주소수정완료(12.26).xlsx')
    df = pd.read_excel(inputfile)
    x = df.values.tolist()
    print(x)
   
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
            자동입력(driver,str(x[n][1]), str(x[n][5]), str(x[n][6]), x[n][2], x[n][3], 정제전번[0], 정제전번[1], 정제전번[2])
            #print(str(x[n][1]), str(x[n][5]), str(x[n][6]), x[n][2], x[n][3], 정제전번[0], 정제전번[1], 정제전번[2])
            print(str(n+1)+"/"+str(len(x))+" 5초 대기")
            time.sleep(5)
        except:
            입력실패번호.append(n)
            print(str(n+1)+"/"+str(len(x))+" 입력실패")
        P_var.set((n+1)*(100/len(x)))
        progressbar.update() 
    if 입력실패번호 == [] :
        print("전부 입력완료")
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
    df.to_excel(savefile, index=False,header=True)
    #print("주소수정 필요 파일 확인필요")

    #정규 출력
    #y=153
    #for x in range(1,y):
        #P_var.set((x+1)*(100/y))
        #print(P_var.get())
        #progressbar.update()
        #time.sleep(1)
    messagebox.showinfo("완료하였습니다.","신고가 완료되었습니다.\n"+savefile+"파일을 확인하세요")
    #button3['state']=tkinter.DISABLED
    driver.close()
    global root
    root.quit()
    quit()    

#추가적으로 driver를 추가
def 자동입력(driver,이름, 주소1, 주소2, 주민번호1, 주민번호2, 전화번호1, 전화번호2, 전화번호3):
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
        #try문은 for문에 있으므로 쓰지 않는다.
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        driver.implicitly_wait(5)
        # 취소하기(닫기)
        #alert.dismiss()
        #print(alert.text)
        if alert.text == "정상적으로 등록되었습니다.\n다른 신고서를 계속 입력하시겠습니까?":
            alert.accept()
            print("신고완료") 
        else:
            print(alert.text)
            alert.accept()
            raise NotImplementedError
        # 확인하기
        
         

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

    
 
#my_btn = Button(root, text='입력할엑셀열기', command=open).pack(side=TOP)
label1 = Label(root, text='년: ',anchor="e")
label1.grid(row=0, column=0)

input1 = Entry(root)
input1.grid(row=0, column=1)
input1.insert(0, "2022")
input1['state'] = tkinter.DISABLED

label2 = Label(root, text='월: ',anchor="e")
label2.grid(row=1, column=0)

input2 = Entry(root)
input2.grid(row=1, column=1)

label3 = Label(root, text='일: ',anchor="e")
label3.grid(row=2, column=0)

input3 = Entry(root)
input3.grid(row=2, column=1)

#label4= Label(root,text="걸러질 파일의 이름: ",anchor="e")
#label4.grid(row=3,column=0)

#input4 = Entry(root)
#input4.grid(row=3, column=1)

button = Button(root, text="출력확인",command=check)
button.grid(row=4, column=1, sticky=W+E+N+S)

my_btn = Button(root, text='입력할엑셀열기', command=open).grid(row=1,column=3)

label5 = Label(root,text="입력할 엑셀 주소: ").grid(row=5,column=1)

button2=Button(root,text="최종 확인",command=final)
button2.grid(row=6, column=1, sticky=W+E+N+S)

button3=Button(root,text="시작",command=main)
button3.grid(row=6, column=2, sticky=W+E+N+S)

#results = Label(root, text="결과가 여기 표시될 예정")
#results.grid(row=7,rowspan=2,column=0,columnspan=5,sticky=W+E+N+S)

progressbar=tkinter.ttk.Progressbar(root, maximum=100, variable=P_var)
progressbar.grid(row=7,rowspan=2,column=0,columnspan=5,sticky=W+E+N+S)


root.mainloop()