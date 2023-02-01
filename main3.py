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
from time import sleep
from qtwidgets import PasswordEdit

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.append(root_dir)
from utils.utils import make_sales_info, get_sales_check_of_credit_card, CODE_MAP_TABLE
from utils.ui import TableView

import csv

import configparser

config = configparser.ConfigParser()

from datetime import datetime, date
import json

from pathlib import Path
import pandas as pd

pattern = ""

SalesType = [
    "중식",
    "석식",
    "야근 택시비",
    "문화마일리지",
    "조식",
    "통신비",
    "국내출장 교통비",
    "국내출장 식대",
    "국내출장 숙박",
    "국내출장 주차비",
    "업무영 도서구입",
    "사무용품",
    "전산용품",
]


from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from datetime import datetime

# Creating tab widgets
class MyTabWidget(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout(self)

        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tabs.resize(300, 200)

        # Add tabs
        self.tabs.addTab(self.tab1, "main")
        self.tabs.addTab(self.tab2, "table")
        self.tabs.addTab(self.tab3, "Geeks")

        self.define_tab1()
        self.define_tab2()

        # Add tabs to widget
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)

    def define_tab1(self):
        # Create first tab
        self.tab1.layout = QFormLayout(self)
        ################################################
        currentYear = str(datetime.now().year)
        currentMonth = str(datetime.now().month)
        # self.buttonSave_tab = QPushButton("Save", self)
        # self.buttonSave_tab.clicked.connect(self.handleSavemon)
        self.month_options = [str(i) for i in list(range(1, 13))]
        self.year_options = [str(i) for i in list(range(2022, 2027))]
        self.combo_year = QComboBox()
        self.combo_year.addItems(self.year_options)
        self.combo_year.setCurrentIndex(self.year_options.index(currentYear))
        self.combo_month = QComboBox()
        self.combo_month.addItems(self.month_options)
        self.combo_month.setCurrentIndex(self.month_options.index(currentMonth))
        ###########
        label_yr = QLabel("Year ")
        label_mn = QLabel("Month")
        self.tab1.layout.addRow(label_yr, self.combo_year)
        self.tab1.layout.addRow(label_mn, self.combo_month)

        self.options = ("Get Folder Dir", "Run")
        # "Get File Name", "Get File Names", "Save File Name"

        self.combo = QComboBox()
        self.combo.addItems(self.options)

        btn = QPushButton("Get Folder Directory")
        btn.clicked.connect(self.getDirectory)
        label_dir = QLabel("Directory")
        self.tab1.layout.addRow(label_dir, btn)
        # TODO: 꾸미기
        # layout.addWidget(btn)

        # # layout.addWidget(self.combo)
        # self._dir_name = ""
        # ## folder directory
        self.textBrowser_dir = QTextBrowser()
        self.textBrowser_dir.setStyleSheet("font-size: 15px;")
        self.textBrowser_dir.setFixedHeight(70)
        self.tab1.layout.addRow(QLabel("폴더위치"), self.textBrowser_dir)
        # layout.addWidget(self.textBrowser_dir)

        # ### EVENT
        self.event_text_edit = QTextEdit()
        self.event_text_edit.setStyleSheet("font-size: 15px;")
        self.event_text_edit.setFixedHeight(50)
        self.event_text_edit.setPlaceholderText(r"이벤트 일자를 넣어 주세요 1,휴가;2,휴가;")
        self.tab1.layout.addRow(QLabel("이벤트기입  "), self.event_text_edit)
        # layout.addWidget(self.event_text_edit)

        # ### API KEY LINK
        self.textBrowser = QTextBrowser()
        self.textBrowser.setOpenExternalLinks(True)
        self.textBrowser.setStyleSheet("font-size: 15px;")
        self.textBrowser.setFixedHeight(100)
        self.textBrowser.append("API KEY를 발급 받아 아래 텍스트 박스에 넣어주세요.")
        self.textBrowser.append("<a href=https://www.data.go.kr/data/15012690/openapi.do>한국천문연구원_특일 정보</a>")

        self.tab1.layout.addRow(QLabel("참고  "), self.textBrowser)
        # layout.addWidget(self.textBrowser)
        self.text_edit = PasswordEdit()

        self.text_edit.setStyleSheet("font-size: 15px;")
        self.text_edit.setFixedHeight(50)
        if Path("./apikey.json").is_file():
            f = open("./apikey.json")
            self.text_edit.setText(str(json.load(f)["key"]))
        else:
            self.text_edit.setPlaceholderText(r"API KEY를 입력해주세요")
        self.tab1.layout.addRow(QLabel("API KEY  "), self.text_edit)
        # layout.addWidget(self.text_edit)

        btn = QPushButton("Run")
        btn.clicked.connect(self.get_table)
        self.tab1.layout.addRow(btn)
        # layout.addWidget(btn)
        self.textBrowser_msg = QTextBrowser()
        self.textBrowser_msg.setStyleSheet("font-size: 15px;")
        self.textBrowser_msg.setFixedHeight(150)
        self.tab1.layout.addRow(self.textBrowser_msg)
        # layout.addWidget(self.textBrowser_msg)
        ########### chrome driver
        btn = QPushButton("Find ChromDriver")
        btn.clicked.connect(self.get_chrome_driver)
        label_dir = QLabel("ChromDriver")
        self.tab1.layout.addRow(label_dir, btn)
        self.textBrowser_chrome_driver_msg = QTextBrowser()
        self.textBrowser_chrome_driver_msg.setStyleSheet("font-size: 15px;")
        self.textBrowser_chrome_driver_msg.setFixedHeight(50)
        label_dir = QLabel("Directory")
        self.tab1.layout.addRow(label_dir, self.textBrowser_chrome_driver_msg)

        # ### USER
        self.user_id = QTextEdit()
        self.user_id.setStyleSheet("font-size: 20px;")
        self.user_id.setFixedHeight(50)
        self.user_id.setPlaceholderText(r"USER의 ID를 입력하세요")
        label = QLabel("USER  ")
        label.setFont(QFont("Arial", 10))
        self.tab1.layout.addRow(label, self.user_id)
        # ### PW

        self.pw_id = PasswordEdit()
        # self.pw_id.setText("")
        # self.pw_id = QTextEdit()
        # self.pw_id.setStyleSheet("font-size: 20px;")
        self.pw_id.setFixedHeight(50)
        self.pw_id.setPlaceholderText(r"USER의 PW를 입력하세요")
        label = QLabel("PASSWORD ")
        label.setFont(QFont("Arial", 10))
        self.tab1.layout.addRow(label, self.pw_id)
        btn2 = QPushButton("제출 실행")
        btn2.clicked.connect(self.run_acc_selenium)
        self.tab1.layout.addRow(btn2)
        # font = QFont()
        # font.setPointSize(1)

        # password = PasswordEdit()
        # password.setText("hi")
        # self.tab1.layout.addRow(password)
        self.tab1.setLayout(self.tab1.layout)

    def define_tab2(self):

        self.tab2.layout = QFormLayout(self)
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
        self.tab2.layout.addRow(readme)

        font = QFont()
        font.setPointSize(10)
        self.tableWidget = QTableWidget()

        self.tableWidget.setRowCount(len(CODE_MAP_TABLE))
        self.tableWidget.setColumnCount(2)
        for idx, (key, value) in enumerate(CODE_MAP_TABLE.items()):
            # self.tableWidget.item(idx, 0).setFont(font)
            # self.tableWidget.item(idx, 1).setFont(font)
            v = QTableWidgetItem(key)
            v.setFont(font)
            self.tableWidget.setItem(idx, 0, v)
            v = QTableWidgetItem(value)
            v.setFont(font)
            self.tableWidget.setItem(idx, 1, v)

        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.verticalScrollBar().setValue(0)
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # self.tableWidget.horizontalHeader().setFixedHeight(100)
        font = QFont()
        font.setPointSize(15)
        self.tableWidget.setHorizontalHeaderLabels(["태그", "표준적요코드"])
        self.tableWidget.horizontalHeader().setDefaultSectionSize(15)
        self.tab2.layout.addRow(self.tableWidget)
        ##############
        self.tab2.setLayout(self.tab2.layout)

    def get_chrome_driver(self):
        dir_name = QFileDialog.getExistingDirectory(self, caption="Select a ChromDriver Folder")
        response = dir_name
        self.textBrowser_chrome_driver_msg.append(response)

    def get_table(self):
        try:
            api_key = self.text_edit.text()
            print("API KEY : ", api_key)
            if self._dir_name == "":
                self.textBrowser_msg.append("Warnings....")
                self.textBrowser_msg.append("폴더 경로를 정하지 않았습니다.")
                self.textBrowser_msg.append("매출전표가 있는 폴더 경로부터 지정해주세요.")
            if api_key == "":
                self.textBrowser_msg.append("Warnings....")
                self.textBrowser_msg.append("API KEY를 입력하지 않았습니다")
                self.textBrowser_msg.append("API KEY를 입력해주세요")
            month = int(self.combo_month.currentText())
            year = int(self.combo_year.currentText())
            result = get_sales_check_of_credit_card(self._dir_name, year, month, api_key)

            event_list = self.event_text_edit.toPlainText().strip().split(";")

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

            result.to_excel(writer, sheet_name="식대")

            n_day_of_unique = result["day"].nunique()
            event = ["휴가", "휴일"]
            n_day_of_not_event_unique = result.query("특이사항 not in @event")["day"].nunique()
            luncu_tag = ["점심", "중식"]
            lunch_total = result.query("태그 in @luncu_tag")["총합"].sum()
            lunch_event = pd.DataFrame(
                [
                    {
                        "총근무일수": n_day_of_unique,
                        "총근무일수(이벤트제외)": n_day_of_not_event_unique,
                        "현재까지 점심 사용 금액": lunch_total,
                        "총사용가능금액(이벤트제외)": n_day_of_not_event_unique * 10_000,
                        "잔여금액(이벤트제외)": n_day_of_not_event_unique * 10_000 - lunch_total,
                    }
                ]
            )
            lunch_event.to_excel(writer, sheet_name="점심식대")
            event_total_df = result.groupby(["태그"], as_index=False).sum()
            event_total_df.to_excel(writer, sheet_name="총태그합")
            writer.close()

            self.textBrowser_msg.clear()
            self.textBrowser_msg.append("결과 테이블 저장 완료...[1/2]")

            report_sales_info = make_sales_info(result_table=result)
            report_sales_info.to_excel(writer2, index=False)
            writer2.close()
            self.textBrowser_msg.append("sales 테이블 저장 완료...[2/2]")

            self.textBrowser_msg.append(f"폴더 경로 : {Path(result_path).parent}")

            Path(result_path2).name
            self.textBrowser_msg.append(f"경비 파일 : {Path(result_path).name}")
            self.textBrowser_msg.append(f"제출 파일 : {Path(result_path2).name}")
            self.textBrowser_msg.append(f"제출 파일은 확인 및 일부 수정이 필요합니다.")
            folder_path = str(Path(os.path.abspath(result_path)).parent)
            self.textBrowser_msg.append(f"해당 파일이 있는 경로를 엽니다. 파일을 확인해주세요")
            os.startfile(folder_path)
            return 200
        except Exception as e:
            # self.textBrowser_msg.clear()
            self.textBrowser_msg.append("Error... 테이블 생성 에러...")
            self.textBrowser_msg.append(str(e))
            return 404

    def run_acc_selenium(self):
        try:
            result_path2 = f"{self._dir_name}/sample_data.xlsx"
            # result_path2 = str(Path(result_path2)).replace("\\", "/").strip()
            #  -f {result_path2}
            import json
            import shutil

            to_file_path = "./acc_contents_selenium/sample_data.xlsx"
            shutil.copyfile(result_path2, to_file_path)
            chrom_driver_path = self.textBrowser_chrome_driver_msg.toPlainText()
            os.system(
                f"python ./acc_contents_selenium/acc_selenium.py -id {self.user_id.toPlainText().strip()} -pw {self.pw_id.text().strip()} -f {to_file_path} -c {chrom_driver_path}"
            )
        except Exception as e:
            self.textBrowser_msg.append(str(e))
            return 404
        else:
            return 200

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
        self.window_width, self.window_height = 800, 200
        # self.setMinimumSize(self.window_width, self.window_height)
        self.setGeometry(self.left, self.top, self.window_width, self.window_height)

        self.tab_widget = MyTabWidget(self)
        self.setCentralWidget(self.tab_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(
        """
        QWidget {
            font-size: 35px;
        }
    """
    )

    myApp = MyApp()
    myApp.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Closing Window...")
