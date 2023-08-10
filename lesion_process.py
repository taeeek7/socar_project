from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

options = webdriver.ChromeOptions()

prefs = {
    'profile.default_content_setting_values': {
        'automatic_downloads': 1, 'images': 2, 'plugins': 2, 'popups': 2, 'geolocation': 2, 'notifications': 2, 'auto_select_certificate': 2,
        'fullscreen': 2, 'mouselock': 2, 'mixed_script': 2, 'media_stream': 2, 'media_stream_mic': 2, 'media_stream_camera': 2, 'protocol_handlers': 2, 'ppapi_broker': 2,
        'midi_sysex': 2, 'push_messaging': 2, 'ssl_cert_decisions': 2, 'metro_switch_to_desktop': 2, 'protected_media_identifier': 2, 'app_banner': 2, 'site_engagement': 2,
        'durable_storage': 2
    }
}

options.add_experimental_option('prefs', prefs)

# options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument("disable-gpu")
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)


driver.implicitly_wait(15)
# self.maximize_window()
driver.get("https://partnerop.socar.kr/partner/lesion_index")

# 로그인
search_box = driver.find_element_by_name("adminid")
search_box.send_keys("****")
search_box = driver.find_element_by_name("password")
search_box.send_keys("****")

# 검색 버튼 클릭
from selenium.webdriver.common.keys import Keys
import time

search_button = driver.find_element_by_class_name("btnB")
search_button.click()
time.sleep(1)

try:
    alert = driver.switch_to.alert
    alert.accept()
    print("Alert accepted")
except:
    pass


import pandas as pd
import gspread
import pydata_google_auth
import time
from selenium.webdriver.support.ui import Select, WebDriverWait

scope = ['https://www.googleapis.com/auth/spreadsheets']

# json_file_name = json_file_name
credentials = pydata_google_auth.get_user_credentials(scope, auth_local_webserver=True)

#구글스프레드시트 불러오기 
gs_c = gspread.authorize(credentials)
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1noePclfL8_TpRtZxuNscpc5HQ7d_lUjGHhyzM7kgY-g/edit#gid=0'
doc = gs_c.open_by_url(spreadsheet_url)
worksheet = doc.worksheet('장애카드처리_python')
df_a = pd.DataFrame(worksheet.get_all_records())


#표 출력 옵션
pd.set_option('display.width', 100)  # 표 최대 폭을 1000으로 설정
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.max_rows', None)  # 모든 행 출력




###반복문 실행 시 준비 구문###
last_row = df_a.tail(1).index[0] + 1
i = 0



# 장애카드번호 반복문
for i in range(1,last_row) : 
    
    #변수 설정
    lesion_id = df_a.loc[i,'장애카드번호']
    content = df_a.loc[i,'점검 상세내용']
    result = df_a.loc[i,'조치 유형']
    
    #장애카드링크 열기
    driver.get(f"https://partnerop.socar.kr/car_lesion/detail_lesion?id={lesion_id}")
    
    #업무배정 선택
    select_element_job = driver.find_element(By.NAME, 'job_assignment')
    select_job = Select(select_element_job)
    select_job.select_by_value("OUTSIDE")
    
    #처리구분 선택
    select_element_type = driver.find_element(By.NAME, 'process_type')
    select_type = Select(select_element_type)
    select_type.select_by_value("FIELD")
    
    #등록 클릭 (최하단)
    driver.find_element(By.XPATH, '//*[@id="car_accident_record"]').click()
    
    time.sleep(2)
    
    #코멘트 입력
    driver.find_element(By.NAME, 'memo').send_keys(content)
    
    #등록 클릭 (메모)
    driver.find_element(By.XPATH, '//*[@id="memo_ok"]').click()
    
    time.sleep(2)
    
    #상태변경
    if result == '조치완료' :
        driver.find_element(By.XPATH, '//*[@id="input_form"]/fieldset/table[2]/tbody/tr/td[5]/button[2]').click()
    else : 
        driver.find_element(By.XPATH, '//*[@id="input_form"]/fieldset/table[2]/tbody/tr/td[5]/button[3]').click()
    
    i+=1
    time.sleep(2)


  driver.quit()

