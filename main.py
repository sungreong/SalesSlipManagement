# -*- coding: utf-8 -*-

import sys
import os
from urllib import response
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QComboBox,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QTextEdit,
    QTextBrowser,
    QLabel,
    QTableWidget,
)
from qtwidgets import PasswordEdit
from datetime import datetime
import configparser
from pathlib import Path
import pandas as pd
import json
import shutil
import re

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.append(root_dir)
from utils.utils import (
    make_sales_info,
    get_sales_check_of_credit_card,
    split_text,
    detect_encoding,
    get_img_list,
)
from utils.bizbox import BizBoxCommon
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *


def read_config():
    config = configparser.ConfigParser(interpolation=None)
    config.read("./info.ini")
    my_config_parser_dict = {s: dict(config.items(s)) for s in config.sections()}
    return my_config_parser_dict


def generate_error(e):
    log = f"Error Type : {type(e).__name__}\n File : {__file__}\n Line Number : {e.__traceback__.tb_lineno}\n Error msg={str(e)}"
    return log


# Creating tab widgets
class MyTabWidget(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)
        self.config = read_config()
        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tabs.resize(300, 400)

        # Add tabs
        self.tabs.addTab(self.tab1, "main")
        self.tabs.addTab(self.tab2, "table")
        self.tabs.addTab(self.tab3, "pattern")

        self.define_tab1()
        self.define_tab2()
        self.define_tab3()
        # Add tabs to widget
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)

    def define_tab1(self):
        # Create first tab
        self.tab1.layout = QVBoxLayout(self)
        ################################################
        currentYear = str(datetime.now().year)
        currentMonth = str(datetime.now().month)

        self.month_options = [str(i) for i in list(range(1, 13))]
        self.year_options = [str(i) for i in list(range(2022, 2027))]
        self.combo_year = QComboBox()
        self.combo_year.addItems(self.year_options)
        self.combo_year.setCurrentIndex(self.year_options.index(currentYear))
        self.combo_month = QComboBox()
        self.combo_month.addItems(self.month_options)
        self.combo_month.setCurrentIndex(self.month_options.index(currentMonth))


        self.tab1.layout.addWidget(QLabel("지출결의서 정리 프로그램"))
        ###########
        calender_layout = QHBoxLayout()
        calender_layout.addWidget(QLabel("Year "))
        calender_layout.addWidget(self.combo_year)
        calender_layout.addWidget(QLabel("Month"))
        calender_layout.addWidget(self.combo_month)
        self.tab1.layout.addLayout(calender_layout)
        ###########

        
        self.user_id = QTextEdit()
        self.user_id.setStyleSheet("font-size: 20px;")
        self.user_id.setFixedHeight(50)
        self.user_id.setPlaceholderText(r"USER의 ID를 입력하세요")
        
        self.pw_id = PasswordEdit()
        self.pw_id.setStyleSheet("font-size: 20px;")
        self.pw_id.setFixedHeight(50)

        login_layout = QHBoxLayout()
        login_layout.addWidget(QLabel("ID"))
        login_layout.addWidget(self.user_id)
        login_layout.addWidget(QLabel("PW "))
        login_layout.addWidget(self.pw_id)
        self.tab1.layout.addLayout(login_layout)
        #############
        self.options = ("Get Folder Dir", "Run")

        self.combo = QComboBox()
        self.combo.addItems(self.options)

        dir_layout = QHBoxLayout()
        btn = QPushButton("실행")
        btn.clicked.connect(self.getDirectory)
        
        
        self.textBrowser_dir = QTextBrowser()
        self.textBrowser_dir.setStyleSheet("font-size: 15px;")
        self.textBrowser_dir.setFixedHeight(40)

        dir_layout.addWidget(QLabel("영수증폴더"))
        dir_layout.addWidget(self.textBrowser_dir)
        dir_layout.addWidget(btn)
        self.tab1.layout.addLayout(dir_layout)

        self.event_text_edit = QTextEdit()
        self.event_text_edit.setStyleSheet("font-size: 15px;")
        self.event_text_edit.setFixedHeight(40)
        self.event_text_edit.setPlaceholderText(r"휴가 일자를 넣어 주세요 1,휴가/2,휴가/")
        
        event_layout = QHBoxLayout()
        event_layout.addWidget(QLabel('휴가일정   '))
        event_layout.addWidget(self.event_text_edit)
        btn = QPushButton("실행")
        btn.clicked.connect(self.getHolidayEvent)
        event_layout.addWidget(btn)
        self.tab1.layout.addLayout(event_layout)

        self.textBrowser = QTextBrowser()
        self.textBrowser.setOpenExternalLinks(True)
        self.textBrowser.setStyleSheet("font-size: 15px;")
        self.textBrowser.setFixedHeight(50)
        self.textBrowser.setFixedWidth(300)
        self.textBrowser.append("API KEY를 발급 받아 왼쪽에 입력")
        self.textBrowser.append("<a href=https://www.data.go.kr/data/15012690/openapi.do>한국천문연구원_특일 정보</a>")

        # api_layout = QHBoxLayout()
        # api_layout.addWidget(QLabel("참고  "))
        
        # self.tab1.layout.addLayout(api_layout)

        self.text_edit = PasswordEdit()

        self.text_edit.setStyleSheet("font-size: 15px;")
        self.text_edit.setFixedHeight(50)
        
        api_key_layout = QHBoxLayout()
        api_key_layout.addWidget(QLabel('API KEY    '))
        api_key_layout.addWidget(self.text_edit)
        api_key_layout.addWidget(self.textBrowser)
        self.tab1.layout.addLayout(api_key_layout)



        btn = QPushButton("실행")
        btn.clicked.connect(self.get_table)
        summary_layout = QHBoxLayout()
        summary_layout.addWidget(QLabel('영수증분석  '))
        summary_layout.addWidget(btn)
        self.tab1.layout.addLayout(summary_layout)

        self.textBrowser_msg = QTextBrowser()
        self.textBrowser_msg.setStyleSheet("font-size: 15px;")
        self.textBrowser_msg.setFixedHeight(150)
        self.tab1.layout.addWidget(self.textBrowser_msg)



        btn2 = QPushButton("실행")
        btn2.clicked.connect(self.run_acc_selenium)
        

        submission_layout = QHBoxLayout()
        submission_layout.addWidget(QLabel('지출결의서 제출  '))
        submission_layout.addWidget(btn2)
        self.tab1.layout.addLayout(submission_layout)

        self.tab1.setLayout(self.tab1.layout)
        self.update_info()

    def update_info(self) :

        if "USER" in self.config:
            if "id" in self.config["USER"]:
                if len(self.config["USER"]["id"]) > 1:
                    self.user_id.setText(self.config["USER"]["id"])
                else:
                    self.user_id.setPlaceholderText(r"(bizbox)USER의 ID를 입력하세요(info.ini 에 USER id를 입력하면 자동으로 들어갑니다.")
            else:
                self.user_id.setPlaceholderText(r"(bizbox)USER의 ID를 입력하세요(info.ini 에 USER id를 입력하면 자동으로 들어갑니다.")
        else:
            self.user_id.setPlaceholderText(r"(bizbox)USER의 ID를 입력하세요(info.ini 에 USER id를 입력하면 자동으로 들어갑니다.")
        if "USER" in self.config:
            if "pw" in self.config["USER"]:
                if len(self.config["USER"]["pw"]) > 1:
                    self.pw_id.setText(self.config["USER"]["pw"])
                else:
                    self.pw_id.setPlaceholderText(
                        r"(bizbox)USER의 PW를 입력하세요(info.ini 에 USER password를 입력하면 자동으로 들어갑니다."
                    )
            else:
                self.pw_id.setPlaceholderText(r"(bizbox)USER의 PW를 입력하세요(info.ini 에 USER password를 입력하면 자동으로 들어갑니다.")
        else:
            self.pw_id.setPlaceholderText(r"(bizbox)USER의 PW를 입력하세요(info.ini 에 USER password를 입력하면 자동으로 들어갑니다.")
        if "API" in self.config:
            if "key" in self.config["API"]:
                if len(self.config["API"]["key"]) > 1:
                    self.text_edit.setText(self.config["API"]["key"])
                else:
                    self.text_edit.setPlaceholderText(r"API KEY를 입력해주세요(info.ini 에 API KEY를 입력하면 자동으로 들어갑니다.")
            else:
                self.text_edit.setPlaceholderText(r"API KEY를 입력해주세요(info.ini 에 API KEY를 입력하면 자동으로 들어갑니다.")
        else:
            self.text_edit.setPlaceholderText(r"API KEY를 입력해주세요(info.ini 에 API KEY를 입력하면 자동으로 들어갑니다.")


    def getHolidayEvent(self) :
        try : 
            bizbox = BizBoxCommon(user=self.user_id.toPlainText().strip(), pw=self.pw_id.text())
            bizbox.connect_site()
            bizbox.login()
            bizbox.access_main()
            bizbox.access_sub()
            event_holiday_text = bizbox.get_holiday_event()
        except Exception as e :
            return QMessageBox.warning(self, "경고", "휴가 정보를 가져오는 중에 오류가 발생했습니다. \n" + str(e), QMessageBox.Ok)
        else :
            self.event_text_edit.setText(event_holiday_text)
            QMessageBox.information(self, "알림", "휴가 정보를 가져왔습니다.", QMessageBox.Ok)
        finally :
            return  bizbox.close()



    def define_tab2(self):

        self.tab2.layout = QFormLayout(self)
        today = datetime.now().strftime("%Y%m%d")
        msg = f"""
        [standardbriefscode.csv]

        위의 파일에 추가적으로 간소화된 표현을 작성하고,
        표준적요설명을 적어주시면 제출할 때 반영됩니다.
        
        (YYYYMMDD)+(TAG)+(PRICE)+(특이사항).jpg        
        예) {today}+점심+2_000+1명.jpg
        예) {today}+점심+50000+5명.jpg
        """

        readme = QLabel(msg)
        readme.setFixedWidth(500)  # +++
        readme.setMinimumHeight(30)
        font1 = readme.font()
        font1.setPointSize(5)
        self.tab2.layout.addRow(readme)
        self.add_new_tag_text_edit = QTextEdit()
        self.add_new_tag_text_edit.setPlaceholderText('표준적요 코드를 입력하세요.(간소화,표준적요설명,표준적요코드)')
        self.add_new_tag_text_edit.setFixedHeight(50)
        self.add_new_tag_button = QPushButton("추가")
        self.add_new_tag_button.clicked.connect(self.add_tag_button)
        addtag_layout = QHBoxLayout()
        addtag_layout.addWidget(QLabel("추가 입력"))
        addtag_layout.addWidget(self.add_new_tag_text_edit)
        addtag_layout.addWidget(self.add_new_tag_button)
        self.tab2.layout.addRow(addtag_layout)

        font = QFont()
        font.setPointSize(10)
        self.tableWidget = QTableWidget()
        encoding = detect_encoding("./standardbriefscode.csv")
        MAPPING_TABLE = pd.read_csv("./standardbriefscode.csv", encoding=encoding)
        self.tableWidget.setRowCount(MAPPING_TABLE.shape[0])
        self.tableWidget.setColumnCount(MAPPING_TABLE.shape[1])
        self.new_row_idx = 0
        for idx, one_row in MAPPING_TABLE.iterrows():
            for idx2, (_, v) in enumerate(one_row.to_dict().items()):
                v = QTableWidgetItem(str(v))
                v.setFont(font)
                self.tableWidget.setItem(idx, idx2, v)
            self.new_row_idx = idx 

        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.verticalScrollBar().setValue(0)
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # self.tableWidget.horizontalHeader().setFixedHeight(100)
        font = QFont()
        font.setPointSize(10)
        self.tableWidget.setHorizontalHeaderLabels(list(MAPPING_TABLE))
        # self.tableWidget.horizontalHeader().setDefaultSectionSize(15)

        self.tab2.layout.addRow(self.tableWidget)
        ##############
        self.tab2.setLayout(self.tab2.layout)

    def add_tag_button(self) :
        self.new_row_idx += 1
        new_tag = self.add_new_tag_text_edit.toPlainText().strip()
        if len(new_tag.split(",")) == 3 :
            pass 
        else :
            return QMessageBox.warning("경고", "표준적요 코드를 입력하세요.(간소화,표준적요설명,표준적요코드)", QMessageBox.Ok)
        font = QFont()
        font.setPointSize(10)
        self.tableWidget.insertRow(self.new_row_idx)
        for idx2 , v in enumerate(new_tag.split(",")) :
            v = QTableWidgetItem(str(v).strip())
            v.setFont(font)
            self.tableWidget.setItem(self.new_row_idx, idx2, v)
        if QMessageBox.question(self, "알림", "표준적요 코드를 추가했습니다. 저장하시겠습니까?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes :
            import csv
            with open("./standardbriefscode.csv", 'a', newline='', encoding='utf-8') as f:
                # 4. csv.writer 객체 생성
                writer = csv.writer(f)
                # 5. writerow() 메서드로 데이터 쓰기
                writer.writerow([i.strip() for i in new_tag.split(",")])
            QMessageBox.information(self, "알림", "저장했습니다.", QMessageBox.Ok)


    def define_tab3(self):

        self.tab3.layout = QFormLayout(self)

        today = datetime.now().strftime("%Y%m%d")
        msg = f"""
        (YYYYMMDD)+(TAG)+(PRICE).jpg        
        예시
        {today}+점심+2_000+1.jpg
        {today}+점심+5_000+2.jpg
        {today}+통신비+15_000.jpg
        """

        readme = QLabel(msg)
        readme.setFixedWidth(1000)  # +++
        readme.setMinimumHeight(30)
        font1 = readme.font()
        font1.setPointSize(10)
        self.tab3.layout.addRow(readme)

        self.file_pattern_id = QTextEdit()
        self.file_pattern_id.setStyleSheet("font-size: 20px;")
        self.file_pattern_id.setFixedHeight(50)
        self.file_pattern_id.setPlaceholderText(r"예시 파일 이름을 넣어보세요")
        label = QLabel("PATTERN  ")
        label.setFont(QFont("Arial", 10))
        self.tab3.layout.addRow(label, self.file_pattern_id)

        self.file_pattern_result = QTextBrowser()
        self.file_pattern_result.setStyleSheet("font-size: 20px;")
        self.file_pattern_result.setFixedHeight(50)
        label = QLabel("테스트결과  ")
        label.setFont(QFont("Arial", 10))
        self.tab3.layout.addRow(label, self.file_pattern_result)

        btn = QPushButton("테스트 실행")
        btn.clicked.connect(self.get_pattern)
        self.tab3.layout.addRow(btn)

        self.tab3.setLayout(self.tab3.layout)

    def get_pattern(self):
        file_name = self.file_pattern_id.toPlainText().strip()
        results = split_text(Path(file_name))
        self.file_pattern_result.clear()
        self.file_pattern_result.append(str(results))

    def get_chrome_driver(self):
        dir_name = QFileDialog.getExistingDirectory(self, caption="Select a ChromDriver Folder")
        response = dir_name
        try:
            if Path(dir_name).joinpath("chromedriver.exe").is_file() is False:
                raise FileNotFoundError("chromedriver.exe 존재하지 않습니다.")
        except FileNotFoundError as e:
            self.textBrowser_msg.append(f"Folder : {dir_name}")
            self.textBrowser_msg.append("해당 폴더에는 chromdriver.exe가 존재하지 않습니다.")
            return 404
        else:
            self.textBrowser_chrome_driver_msg.append(response)

    def check_exist_img_directory(self):
        if hasattr(self, "_dir_name") is False:
            self.textBrowser_msg.append("Warnings....")
            self.textBrowser_msg.append("폴더 경로를 정하지 않았습니다.")
            self.textBrowser_msg.append("매출전표가 있는 폴더 경로부터 지정해주세요.")
            raise NotADirectoryError("Directory를 설정해주셔야 합니다.")

    def get_table(self):
        try:
            self.textBrowser_msg.clear()
            api_key = self.text_edit.text()
            print("API KEY : ", api_key)
            self.check_exist_img_directory()
            img_file_list = get_img_list(self._dir_name)
            yyyymmdd = img_file_list[0].name.split("+")[0]
            print(str(img_file_list[0]), yyyymmdd)
            img_month = datetime.strptime(yyyymmdd, "%Y%m%d").month
            if int(img_month) != int(self.combo_month.currentText()):
                raise Exception(
                    f"image에 있는 날짜와 Month에 있는 날짜가 다릅니다. {int(img_month)} != {int(self.combo_month.currentText())}"
                )
            if api_key == "" or len(api_key) == 0:
                self.textBrowser_msg.append("Warnings....")
                self.textBrowser_msg.append("API KEY를 입력하지 않았습니다")
                self.textBrowser_msg.append("API KEY를 입력해주세요")
                raise Exception("참고를 참고하여 API를 동록해주셔야 합니다.")
            month = int(self.combo_month.currentText())
            year = int(self.combo_year.currentText())

            self.textBrowser_msg.append("테이블 작업 진행중...")
            result = get_sales_check_of_credit_card(self._dir_name, year, month, api_key)

            event_list = self.event_text_edit.toPlainText().strip().split("/")

            for event in event_list:
                if event == "":
                    continue
                day, event_name = event.split(",")
                date_time = pd.to_datetime(f"{year}/{month}/{day}", format="%Y/%m/%d")
                result.loc[result["day"] == date_time, "특이사항"] = event_name

            result_path = f"{self._dir_name}/result.xlsx"
            result_path2 = f"{self._dir_name}/sample_data.xlsx"
            writer = pd.ExcelWriter(result_path, engine="xlsxwriter")
            writer2 = pd.ExcelWriter(result_path2, engine="xlsxwriter")
            result2 = result.copy()
            result2["day"] = result2["day"].dt.strftime("%Y%m%d").astype(int)
            result2.to_excel(writer, sheet_name="지출내역")

            n_day_of_unique = result["day"].nunique()
            EVENT = ["휴가", "휴일"]
            n_day_of_not_event_unique = result.query("특이사항 not in @EVENT")["day"].nunique()

            event_total_df = result.groupby(["태그"], as_index=False).sum()
            event_total_df.to_excel(writer, sheet_name="총태그합")
            report_sales_info = make_sales_info(result_table=result)
            event_total_df = report_sales_info[["type", "amount"]].groupby(["type"], as_index=False).sum()
            event_total_df.to_excel(writer, sheet_name="총태그합(제출버전)")
            lunch_total = event_total_df.query("type == '중식대금(휴일포함)'")["amount"].values[0]

            number_of_used_file = result2["특이사항"].apply(lambda x: len(x) if type(x) == list else 0).sum()
            number_of_img_file = len(get_img_list(self._dir_name))
            if number_of_used_file != number_of_img_file:
                raise Exception("이미지를 전부 사용하지 않았습니다. 이미지 파일 패턴을 확인해주세요")
            summary_table = pd.DataFrame(
                [
                    {
                        "총근무일수": n_day_of_unique,
                        "총근무일수(이벤트제외)": n_day_of_not_event_unique,
                        "사용한 이미지 수": number_of_used_file,
                        "등록한 이미지 수": number_of_img_file,
                        "[점심]사용금액": lunch_total,
                        "[점심]총사용가능금액(이벤트제외)": n_day_of_not_event_unique * 10_000,
                        "[점심]잔여금액(이벤트제외)": n_day_of_not_event_unique * 10_000 - lunch_total,
                    }
                ]
            )
            summary_table.to_excel(writer, sheet_name="요약(체크)")
            writer.close()

            self.textBrowser_msg.clear()
            self.textBrowser_msg.append("결과 테이블 저장 완료...[1/2]")

            report_sales_info.to_excel(writer2, index=False)
            writer2.close()
            self.textBrowser_msg.append("sales 테이블 저장 완료...[2/2]")

            self.textBrowser_msg.append(f"폴더 경로 : {Path(result_path).parent}")

            Path(result_path2).name
            self.textBrowser_msg.append(f"경비 파일 : {Path(result_path).name}")
            self.textBrowser_msg.append(f"제출 파일 : {Path(result_path2).name}")
            # self.textBrowser_msg.append(f"제출 파일은 확인 및 일부 수정이 필요합니다.")
            folder_path = str(Path(os.path.abspath(result_path)).parent)
            self.textBrowser_msg.append(f"해당 파일이 있는 경로를 엽니다. 파일을 확인해주세요")
            os.startfile(folder_path)
            return QMessageBox.information(self, "알림", "영수증 정리가 완료되었습니다. \n 제출 파일(sample_data.xlsx)은 확인 및 일부 수정이 필요합니다.", QMessageBox.Ok)
        except Exception as e:
            # self.textBrowser_msg.clear()
            print(generate_error(e))
            self.textBrowser_msg.append("Error... 테이블 생성 에러...")
            self.textBrowser_msg.append(generate_error(e))
            return QMessageBox.warning(self, "오류", f"테이블 생성 에러 {generate_error(e)}", QMessageBox.Ok)

    def run_acc_selenium(self):
        try:
            self.check_exist_img_directory()
            result_path2 = f"{self._dir_name}/sample_data.xlsx"

            to_file_path = "./acc_contents_selenium/sample_data.xlsx"
            shutil.copyfile(result_path2, to_file_path)
            os.system(
                f"python ./acc_contents_selenium/acc_selenium.py -id {self.user_id.toPlainText().strip()} -pw {self.pw_id.text().strip()} -f {to_file_path}"
            )
        except Exception as e:
            self.textBrowser_msg.append(generate_error(e))
            return QMessageBox.warning(self, "오류", f"제출 중 에러 발생 {generate_error(e)}", QMessageBox.Ok)
        else:
            return QMessageBox.information(self, "알림", "영수증 제출이 완료되었습니다. \n 제출 파일(sample_data.xlsx)은 확인 및 일부 수정이 필요합니다.", QMessageBox.Ok)

    def getFileName(self):
        file_filter = "Data File (*.xlsx *.csv *.dat);; Excel File (*.xlsx *.xls)"
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select a data file",
            directory=os.getcwd(),
            filter=file_filter,
            initialFilter="Excel File (*.xlsx *.xls)",
        )
        return response[0]

    def getFileNames(self):
        file_filter = "Data File (*.xlsx *.csv *.dat);; Excel File (*.xlsx *.xls)"
        response = QFileDialog.getOpenFileNames(
            parent=self,
            caption="Select a data file",
            directory=os.getcwd(),
            filter=file_filter,
            initialFilter="Excel File (*.xlsx *.xls)",
        )
        return response[0]

    def getDirectory(self):
        self._dir_name = QFileDialog.getExistingDirectory(self, caption="Select a folder")
        response = self._dir_name

        self.textBrowser_dir.append(response)
        return response

    def getSaveFileName(self):
        file_filter = "Data File (*.xlsx *.csv *.dat);; Excel File (*.xlsx *.xls)"
        response = QFileDialog.getSaveFileName(
            parent=self,
            caption="Select a data file",
            directory="Data File.dat",
            filter=file_filter,
            initialFilter="Excel File (*.xlsx *.xls)",
        )
        return response[0]


class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = "매출전표 정리"
        self.left = 10
        self.top = 10
        self.setWindowTitle(self.title)
        self.window_width, self.window_height = 800, 300
        # self.setMinimumSize(self.window_width, self.window_height)
        self.setGeometry(self.left, self.top, self.window_width, self.window_height)

        self.tab_widget = MyTabWidget(self)
        self.setCentralWidget(self.tab_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(
        """
        QWidget {
            font-size: 20px;
        }
    """
    )

    myApp = MyApp()
    myApp.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Closing Window...")
