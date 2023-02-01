from selenium import webdriver
from time import sleep
import pandas as pd
import getpass
from tqdm.notebook import tqdm
import warnings

warnings.filterwarnings(action="ignore")
import argparse
import pathlib

parser = argparse.ArgumentParser()
parser.add_argument("-id", type=str, help="id", required=False)
parser.add_argument("-pw", type=str, help="pw", required=False)
parser.add_argument("-c", type=str, help="pw", required=False)
parser.add_argument("-f", type=str, default="./sample_data.xlsx", required=False)
args = vars(parser.parse_args())
ver = "# version 0.0.6"
print(f"그룹웨어 적요 채우기 Personal Version: {ver}")

my_id = args["id"]
# input("아이디를 입력해주세요 : ")
my_passwords = args["pw"]
# getpass.getpass("비밀번호를 입력해주세요 (masking되어 입력됩니다) : ")
chromedriver_path = args["c"]
options = webdriver.ChromeOptions()
print("C:/에 chromedriver 폴더를 만드시고, 자신의 크롬 버전에 맞는 실행파일을 설치하세요.")
import os 
from pathlib import Path
chrom_driver_path = Path(os.path.join(chromedriver_path,"chromedriver.exe"))

if chrom_driver_path.is_file() :
    pass 
else :
    raise FileNotFoundError(f"해당 폴더에 chromedriver.exe는 존재하지 않습니다. : {chrom_driver_path=}")
try:
    driver = webdriver.Chrome(str(chrom_driver_path), options=options)
    br_ver = driver.capabilities["browserVersion"]
    dr_ver = driver.capabilities["chrome"]["chromedriverVersion"].split(" ")[0]
    print("-------------------chrome current ver -------------------")
    print(f"Browser Version: {br_ver}\nChrome Driver Version: {dr_ver}")
except:  # selenium.common.exceptions.SessionNotCreatedException as e:
    print("error : chrome 버전을 반드시 확인해주세요. 업데이트가 필요할 수 있습니다.")

print("사이트로 접속합니다.")
driver.get("http://gw.agilesoda.ai/gw/uat/uia/egovLoginUsr.do")

print("경로에 데이터가 있는지 확인해주세요.")
acc_data = pd.read_excel(args["f"], sheet_name="Sheet1", dtype={"howmany": str, "etc": str, "amount": int})
acc_data.fillna(" ", inplace=True)


search_box_id = driver.find_element_by_xpath('//*[@id="userId"]')
search_box_passwords = driver.find_element_by_xpath('//*[@id="userPw"]')
login_btn = driver.find_element_by_xpath('//*[@id="login_b1_type"]/div[2]/div[2]/form/fieldset/div[2]/div')

print("로그인을 실시합니다")
search_box_id.send_keys(my_id)
search_box_passwords.send_keys(my_passwords)
login_btn.click()

print("---팝업창 유의---")
print("아래 하단 결재양식 > 톱니바퀴 > 휴가신청서와 지출결의서(개인경비)를 추가해주세요.")


driver.switch_to_window(driver.window_handles[0])  # main page

finance_templete = driver.find_element_by_xpath('//*[@id="26"]/a').click()  # 톱니바퀴에서 설정해주어야합니다.(index문제)
driver.switch_to_window(driver.window_handles[-1])

print("금월 적요 항목은 ", len(acc_data), "개 있습니다.")

for i in tqdm(range(len(acc_data))):

    # 항목추가
    driver.switch_to_window(driver.window_handles[-1])
    sleep(2)
    driver.find_element_by_xpath('//*[@id="btnExpendListAdd"]').click()

    sleep(3)
    # 표준적요 찾기
    driver.switch_to_window(driver.window_handles[-1])
    driver.find_element_by_xpath('//*[@id="btnListSummarySearch"]').click()  # 찾기 버튼
    sleep(0.5)
    driver.switch_to_window(driver.window_handles[-1])
    search_box_words = driver.find_element_by_xpath('//*[@id="cmmTxtSearchStr"]')  # 검색버튼
    search_box_words.send_keys(acc_data.loc[i, "type"])  # 데이터 프레임 검색 후 대입
    # 검색 버튼
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    # 제일 위에 있는 표준적요 코드
    sleep(0.5)
    driver.switch_to_window(driver.window_handles[-1])
    driver.find_element_by_xpath('//*[@id="tbl_codePopTbl"]/tbody/tr/td[2]').click()

    # 확인버튼
    driver.find_element_by_xpath('//*[@id="cmmBtnAccept"]').click()
    driver.switch_to_window(driver.window_handles[-1])

    # 적요 내용 읽고 삽입, string 연결
    acc_contents_box = driver.find_element_by_xpath('//*[@id="txtListNote"]')
    acc_contents_box.send_keys(
        str(acc_data.loc[i, "type"] + "( " + acc_data.loc[i, "howmany"] + ") " + acc_data.loc[i, "etc"])
    )

    # 증빙일자 8자리 입력
    prove_date_combobox = driver.find_element_by_xpath('//*[@id="txtListAuthDate"]')
    prove_date_combobox.click()
    sleep(0.5)
    prove_date_combobox.send_keys(str(acc_data.loc[i, "date"]))

    # 공급가액
    amt_box = driver.find_element_by_xpath('//*[@id="txtListStdAmt"]')
    amt_box.click()
    sleep(0.5)
    amt_box.send_keys(str(acc_data.loc[i, "amount"]))

    # 증빙유형 찾기 버튼
    driver.find_element_by_xpath('//*[@id="btnListAuthSearch"]').click()
    # 최근 페이지
    driver.switch_to_window(driver.window_handles[-1])
    sleep(2)
    # 제일 위에 있는 표준적요 코드
    driver.find_element_by_xpath('//*[@id="tbl_codePopTbl"]/tbody/tr[1]/td[2]').click()
    # 확인버튼
    driver.find_element_by_xpath('//*[@id="cmmBtnAccept"]').click()
    sleep(0.5)
    driver.switch_to_window(driver.window_handles[-1])

    # 확인버튼
    confirm_btn = driver.find_element_by_xpath('//*[@id="btnListSave"]').click()
    sleep(0.5)
    driver.switch_to_window(driver.window_handles[-1])
