from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders

import os
import time
from datetime import date
import requests
import subprocess
import re

# TODO : Do not use time.sleep(). time.sleep() is not a good practice.
# =============================================================================
# 현금영수증 발급
class CashReceipt:
    def __init__(self):
        pass

    def open_chrome(self):
        '''
        ### 예전 방법
        # chrome.exe 위치 확인
        if os.path.isfile(r'C:\Program Files\Google\Chrome\Application\chrome.exe'):
            chrome_path = r'wmic datafile where name="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" get Version /value'
        elif os.path.isfile(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'):
            chrome_path = r'wmic datafile where name="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" get Version /value"'
        # 현재 chrome 버전 알아내기
        version = subprocess.check_output(chrome_path, shell=True)
        version = version.decode('utf-8').strip()
        version = re.split('=', version)[1]

        # 현재 chrome 버전에 맞게 ChromeDriver 자동 설치
        release = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
        version = requests.get(release).text
        print(version)
        service = Service(ChromeDriverManager(version=version).install())
        '''
        # chrome_options: 실행이 끝난 후에 크롬창 종료 여부 선택
        chrome_options = Options()
        chrome_options.add_experimental_option('detach', True)
        chrome_options.add_argument('headless') # Hide a browser
        # Execute ChromeDriver
        # driver = webdriver.Chrome(service=service, chrome_options=chrome_options)
        self.driver = webdriver.Chrome()
        return self.driver

    def login_fn(self, id, pw):
        # 페이지 로딩될 때까지 최대 30초 기다림
        wait = WebDriverWait(self.driver, 30) # "driver.implicitly_wait(30)"보다 더 좋은듯듯
        
        # 로그인 페이지로 이동
        self.driver.get('https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3')

        # Click ID 로그인 button (자꾸 됐다 안됐다 해서 코드가 길어짐 : 요소 존재여부 확인 -> 스크롤을 요소로 이동 -> 클릭 가능 여부 확인 -> 클릭)
        login_box = wait.until(EC.presence_of_element_located((By.ID, 'mf_txppWframe_loginboxFrame_anchor24')))
        self.driver.execute_script("arguments[0].scrollIntoView(true);", login_box)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_loginboxFrame_anchor24')))
        login_box.click()
        
        # Enter ID, PW & Click 로그인 button
        wait.until(EC.visibility_of_element_located((By.ID, 'mf_txppWframe_loginboxFrame_iptUserId'))).send_keys(id)
        wait.until(EC.visibility_of_element_located((By.ID, 'mf_txppWframe_loginboxFrame_iptUserPw'))).send_keys(pw)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_loginboxFrame_wq_uuid_884'))).click()

    def apply_receipt_fn(self, issue_purpose, amount, business_reg_num): # 현금영수증 발급
        wait = WebDriverWait(self.driver, 30)

        # Click 탭(전체메뉴-발급-건별발급)
        element = wait.until(EC.presence_of_element_located((By.ID, 'mf_wfHeader_hdGroup005'))) # 요소가 존재하는지 확인
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element) # 스크롤을 해당 요소로 이동 (주로 요소가 팝업 메뉴에 가려져서 클릭이 안되는 경우)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_wfHeader_hdGroup005'))).click()
        wait.until(EC.element_to_be_clickable((By.ID, 'grpMenuLi_46_4606020000'))).click()
        wait.until(EC.element_to_be_clickable((By.ID, 'grpMenuAtag_46_4606020100'))).click()

        # Enter 거래정보
        # 사업자등록번호
        time.sleep(3) # 페이지 활성화될 때까지 기다림
        business_reg_num_box = wait.until(EC.presence_of_element_located((By.ID, 'mf_txppWframe_wframe3_spstCnfrNoEncCntn')))
        self.driver.execute_script("arguments[0].scrollIntoView(true);", business_reg_num_box)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_wframe3_spstCnfrNoEncCntn')))
        business_reg_num_box.send_keys(business_reg_num)
        time.sleep(1) 

        # 거래금액
        amount_box = wait.until(EC.presence_of_element_located((By.ID, 'mf_txppWframe_wframe3_trsAmt')))
        self.driver.execute_script("arguments[0].scrollIntoView(true);", amount_box)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_wframe3_trsAmt')))
        amount_box.send_keys(amount)
        time.sleep(1) 

        # 거래구분
        issue_purpose_box = wait.until(EC.presence_of_element_located((By.ID, 'mf_txppWframe_wframe3_selectbox8')))
        self.driver.execute_script("arguments[0].scrollIntoView(true);", issue_purpose_box)
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_wframe3_selectbox8')))
        Select(issue_purpose_box).select_by_visible_text(issue_purpose)        
        time.sleep(1) 

        # # Click 발급요청 button
        wait.until(EC.element_to_be_clickable((By.ID, 'mf_txppWframe_trigger4'))).click()
        time.sleep(1) # Alert 팝업이 뜨기까지 기다림

        # Click 확인 button on Alert Popup
        popup = wait.until(EC.alert_is_present())
        popup.accept() # 1차 확인
        popup = wait.until(EC.alert_is_present()) 
        popup.accept() # 2차 확인

        # Show 완료 메시지
        # if '완료' in popup.text:
        #     print(popup.text)
        # else:
        #     print('현금영수증 발급 실패')

    def print_fn(self, send_email = True): # 현금영수증 출력

        # chrome_options: 실행이 끝난 후에 크롬창 종료 여부 선택
        chrome_options = Options()
        chrome_options.add_experimental_option('detach', True)
        chrome_options.add_argument('headless') # Hide a browser
        # driver = webdriver.Chrome(service=service, chrome_options=chrome_options)
        driver = webdriver.Chrome()

        # 로그인 페이지로 이동
        driver.get('https://www.hometax.go.kr/websquare/websquare.wq?w2xPath=/ui/pp/index_pp.xml')
        driver.implicitly_wait(30) # 페이지 로딩될 때까지 기다림

        # Click 로그인 button
        driver.switch_to.frame('txppIframe') # Go into 프레임
        driver.find_element(By.XPATH, '//*[@id="textbox976"]').click()
        driver.implicitly_wait(30)

        # Click ID 로그인 button
        driver.switch_to.frame('txppIframe') # Go into 프레임
        # login_btn = driver.find_element(By.XPATH, '//*[@id="anchor15"]')
        # driver.execute_script('arguments[0].click();', login_btn)
        # time.sleep 사용하기
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="anchor15"]').click()
        driver.implicitly_wait(30)

        # Enter ID, PW & Click 로그인 button
        driver.find_element(By.ID, 'iptUserId').send_keys(id)
        driver.find_element(By.ID, 'iptUserPw').send_keys(pw)
        driver.find_element(By.XPATH, '//*[@id="anchor25"]').click()
        driver.implicitly_wait(30)

        # Click 전자계산서, 현금영수증, 신용카드 탭
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="hdTextbox548"]').click()
        driver.implicitly_wait(30)

        # Click 현금영수증 발급결과조회
        driver.switch_to.frame('txppIframe')
        driver.find_element(By.XPATH, '//*[@id="textbox_4606020500"]').click()
        driver.implicitly_wait(30)

        # Click 월별 조회하기
        driver.switch_to.frame('txppIframe')
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="radio1_input_2"]').click()
        # time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="trigger1"]').click()
        driver.implicitly_wait(30)

        # Print 발급결과조회
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="G_grid1___checkbox_chk_0"]').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="trigger11"]').click()
        driver.implicitly_wait(30)

        # Convert to 팝업 창 1
        popups = driver.window_handles
        driver.switch_to.window(popups[1])

        # Click PDF Settings
        driver.find_element(By.XPATH, '//*[@id="mskYn_input_1"]').click()
        driver.find_element(By.XPATH, '//*[@id="trigger1"]').click()

        # Convert to 팝업 창 2
        time.sleep(3)
        popups = driver.window_handles
        driver.switch_to.window(popups[1])

        # Print PDF
        driver.find_element(By.XPATH, '/html/body/div/div/div[1]/table/tbody/tr/td/div/nobr/button[2]').click()
        time.sleep(10) # 다운로드 완료까지 기다림

        # Prepare PDF 이름 변경
        files = [os.path.join(file_path, i) for i in os.listdir(file_path)] # 경로에 있는 모든 파일(abs path)
        new_file = max(files, key = os.path.getctime) # 가장 최근에 생성(수정)된 파일   
        
        # Change PDF 이름
        today = date.today()
        today_formatted = today.strftime('%y%m%d') + '-월세_현금영수증.pdf'
        idx = 1
        while(True):
            if os.path.isfile(os.path.join(file_path, today_formatted)):
                today_formatted = today.strftime('%y%m%d') + '-월세_현금영수증(' + str(idx) + ').pdf'
                idx += 1
            else:
                break
        pdf_output = os.path.join(file_path, today_formatted)
        os.rename(new_file, pdf_output)

        if send_email == True:
            # SMTP(SMTP 서버 url, 포트 정보) 객체 생성
            smtp = smtplib.SMTP('smtp.gmail.com', 587)
            # TLS 암호화
            smtp.starttls()
            
            # 로그인
            smtp.login(sender, pw_email)

            # Write Email
            email = MIMEMultipart()
            email['From'] = sender
            email['To'] = receipient
            month_formatted = today.strftime('%m') + '월 '
            email['Subject'] = Header(s=month_formatted + title, charset='utf-8')
            body = MIMEText(month_formatted + content, _charset='utf-8')
            email.attach(body)

            # Attach File
            attach_file = MIMEBase('application', 'octet-stream')
            with open(pdf_output, 'rb') as f:
                attach_file.set_payload(f.read())
            encoders.encode_base64(attach_file)
            attach_file.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_output))
            email.attach(attach_file)


            # Send Email
            smtp.sendmail(sender, receipient, email.as_string())
            smtp.quit()
    
def main():
    id = 'kristiansand' # 홈택스 ID
    pw = '1q2w#E$R%T' # 홈택스 PW

    amount = '1430000' # 거래금액(원)
    business_reg_num = '7761702078' # 사업자등록번호/전화번호/주민등록번호
    issue_purpose = '지출증빙'
    file_path = '/receipts' # 파일 경로(다운로드)
    # file_path = r'C:\Users\hwlee\Downloads'

    send_email = False # Email 발송 여부
    pw_email = 'czgc mbsj jola rgjr' # gmail 앱 비밀번호
    sender = 'kristiansand555@gmail.com' # 메일 주소
    receipient = 'jyahn@xenoclan.co.kr'
    title = '월세 현금영수증 발급 완료' # 메일 제목
    content = '임대료에 대한 현금영수증 발급내역 보내드립니다.\
        \n\n감사합니다.\n이형우 드림' # 메일 내용

    # for loop로 여러개 발급도 가능함.
    run = CashReceipt()
    run.open_chrome()
    run.login_fn(id, pw)
    run.apply_receipt_fn(issue_purpose, amount, business_reg_num)

if __name__ == '__main__':
    main()
        