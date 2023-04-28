from selenium import webdriver
import chromedriver_autoinstaller
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
from selenium.webdriver.support.ui import Select
from PyQt5.QtWidgets import QApplication, QProgressBar

class BizBoxCommon(object):
    def __init__(self, user="##", pw="##"):
        self.user = user
        self.pw = pw
        self.click_count = 0


    def open(self):
        options = webdriver.ChromeOptions()
        options.add_argument("headless")
        options.add_argument("--window-size=1920,1000")
        options.add_argument("disable-gpu")
        path = chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome(path, options=options)

    def close(self) :
        self.driver.close()

    def connect_site(self):
        self.open()
        print("사이트로 접속합니다.")
        self.driver.get("http://gw.agilesoda.ai/gw/uat/uia/egovLoginUsr.do")

    def login(self):
        search_box_id = self.driver.find_element_by_xpath('//*[@id="userId"]')
        search_box_passwords = self.driver.find_element_by_xpath('//*[@id="userPw"]')
        login_btn = self.driver.find_element_by_xpath(
            '//*[@id="login_b1_type"]/div[2]/div[2]/form/fieldset/div[2]/div'
        )

        print("로그인을 실시합니다")
        try:
            search_box_id.send_keys(self.user)
            search_box_passwords.send_keys(self.pw)
            login_btn.click()
            sleep(0.5)
            self.access_main()
            self.driver.find_element_by_xpath("//div[@class='fm_div fm_sc']").click()

            if self.click_count % 2 == 0:
                button = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@id="1dep"]/span'))
                )
                button.click()
                self.click_count += 1
            else:
                pass
        except Exception as e:
            print("로그인 정보가 올바르지 않은 경우, info.ini 파일생성 후, id와 pw가 잘 설정되어있는 지 확인해주세요.")
            import os
            from pathlib import Path

            folder_path = str(Path(__file__).parent.parent)
            os.startfile(folder_path)
            raise Exception(e)
    def access_main(self):
        self.driver.switch_to_window(self.driver.window_handles[0])  # main page

    def access_sub(self):
        self.driver.get("http://gw.agilesoda.ai/gw/userMain.do")
        self.driver.find_element_by_xpath("//div[@class='fm_div fm_sc']").click()
        self.click_count += 1

    def get_holiday_event(self) :
        # 링크를 찾을 때까지 기다립니다.
        print("change...")
        # link = WebDriverWait(self.driver, 20).until(
        #     EC.presence_of_element_located((By.XPATH, '//a[contains(@class, "jstree-anchor")][.//*[text()="내 일정 전체보기"]]'))
        # )
        # # 링크를 클릭합니다.
        # link.click()

        # link = self.driver.find_element_by_id("301080000_anchor")
        # link.click()
        # print("change...")
        sleep(1)
        iframe_element = self.driver.find_element_by_tag_name("iframe")

        self.driver.switch_to.frame(iframe_element)
        print("change...")
        date_element = self.driver.find_element_by_id("from_date")
        date= "2023-04-01"
        self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)
        date_element = self.driver.find_element_by_id("to_date")
        date= "2023-04-30"
        self.driver.execute_script(f"arguments[0].value = '{date}';", date_element)

        # 버튼을 찾을 때까지 기다립니다.
        button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@class="btn_search" and @type="button" and @value="검색"]'))
        )

        # 버튼을 클릭합니다.
        button.click()
        from bs4 import BeautifulSoup
        html = self.driver.page_source

        # BeautifulSoup 객체를 생성합니다.
        soup = BeautifulSoup(html, 'html.parser')

        # 테이블 요소를 찾습니다.
        table = soup.find("div", {"id": "grid_1"}).table

        # 테이블 행을 가져옵니다.
        rows = table.find_all("tr")

        # 테이블 헤더를 추출합니다.
        header = [th.get_text() for th in rows[0].find_all("th")]

        # 테이블 데이터를 추출합니다.
        data = []
        for row in rows[1:]:
            rowData = [td.get_text() for td in row.find_all("td")]
            data.append(rowData)

        # 데이터를 pandas DataFrame으로 변환합니다.
        df = pd.DataFrame(data, columns=header)

        # 웹 드라이버를 종료하

        from datetime import datetime, timedelta
        def apply_date(date_string) :
            start_date_str, end_date_str = date_string.split(" - ")

            # 날짜 형식을 지정하고 날짜를 파싱합니다.
            date_format = "%Y.%m.%d"
            start_date = datetime.strptime(start_date_str[:10], date_format)
            end_date = datetime.strptime(end_date_str[:10], date_format)

            # 시작 날짜와 끝 날짜 사이의 모든 날짜를 생성합니다.
            dates = [(start_date + timedelta(days=i)).day for i in range((end_date - start_date).days + 1)]
            return dates 

        list_of_lists = df['시간'].apply(lambda x: apply_date(x)).values.tolist()
        # 중첩된 리스트를 하나의 리스트로 결합합니다.
        merged_list = [item for sublist in list_of_lists for item in sublist]

        # 리스트를 집합으로 변환하여 중복 요소를 제거하고 다시 리스트로 변환합니다.
        unique_list = list(set(merged_list))
        holiday_events = "/".join([f"{day},휴가" for day in unique_list])
        return holiday_events

'''
bizbox = BizBoxCommon(user='##', pw='##')
bizbox.connect_site()
bizbox.login()
bizbox.access_main()
bizbox.access_sub()
event_holiday_text = bizbox.get_holiday_event()
'''
