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

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.append(root_dir)
from utils.utils import make_sales_info, get_sales_check_of_credit_card
import csv

import configparser

config = configparser.ConfigParser()

from datetime import datetime, date
import json

from pathlib import Path
import pandas as pd


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

        # self.widget_tab = QWidget()
        # self.buttonSave_tab = QPushButton("Save", self)
        # self.buttonSave_tab.clicked.connect(self.handleSavemon)

        # layout = QVBoxLayout()
        # self.setLayout(layout)
        wid1 = QWidget(self)
        self.setCentralWidget(wid1)
        self.month_options = [str(i) for i in list(range(1, 13))]
        self.year_options = [str(i) for i in list(range(2022, 2027))]

        currentYear = str(datetime.now().year)
        currentMonth = str(datetime.now().month)

        self.combo_year = QComboBox()
        self.combo_year.addItems(self.year_options)
        self.combo_year.setCurrentIndex(self.year_options.index(currentYear))
        self.combo_month = QComboBox()
        self.combo_month.addItems(self.month_options)
        self.combo_month.setCurrentIndex(self.month_options.index(currentMonth))
        mytext = QFormLayout()
        label_yr = QLabel("Year ")
        label_mn = QLabel("Month")
        mytext.addRow(label_yr, self.combo_year)
        mytext.addRow(label_mn, self.combo_month)
        wid1.setLayout(mytext)

        self.options = ("Get Folder Dir", "Run")
        # "Get File Name", "Get File Names", "Save File Name"

        self.combo = QComboBox()
        self.combo.addItems(self.options)

        btn = QPushButton("Get Folder Directory")
        btn.clicked.connect(self.getDirectory)
        label_dir = QLabel("Directory")
        mytext.addRow(label_dir, btn)
        # TODO: 꾸미기
        # layout.addWidget(btn)

        # # layout.addWidget(self.combo)
        # self._dir_name = ""
        # ## folder directory
        self.textBrowser_dir = QTextBrowser()
        self.textBrowser_dir.setStyleSheet("font-size: 15px;")
        self.textBrowser_dir.setFixedHeight(70)
        mytext.addRow(QLabel("폴더위치"), self.textBrowser_dir)
        # layout.addWidget(self.textBrowser_dir)

        # ### EVENT
        self.event_text_edit = QTextEdit()
        self.event_text_edit.setStyleSheet("font-size: 15px;")
        self.event_text_edit.setFixedHeight(50)
        self.event_text_edit.setPlaceholderText(r"이벤트 일자를 넣어 주세요 1,휴가;2,휴가;")
        mytext.addRow(QLabel("이벤트기입  "), self.event_text_edit)
        # layout.addWidget(self.event_text_edit)

        # ### API KEY LINK
        self.textBrowser = QTextBrowser()
        self.textBrowser.setOpenExternalLinks(True)
        self.textBrowser.setStyleSheet("font-size: 15px;")
        self.textBrowser.setFixedHeight(50)
        self.textBrowser.append("API KEY를 발급 받아 아래 텍스트 박스에 넣어주세요.")
        self.textBrowser.append("<a href=https://www.data.go.kr/data/15012690/openapi.do>한국천문연구원_특일 정보</a>")

        mytext.addRow(QLabel("참고  "), self.textBrowser)
        # layout.addWidget(self.textBrowser)
        self.text_edit = QTextEdit(self)
        self.text_edit.setStyleSheet("font-size: 15px;")
        self.text_edit.setFixedHeight(50)
        if Path("./apikey.json").is_file():
            f = open("./apikey.json")
            self.text_edit.setPlainText(str(json.load(f)["key"]))
        else:
            self.text_edit.setPlaceholderText(r"API KEY를 입력해주세요")
        mytext.addRow(QLabel("API KEY  "), self.text_edit)
        # layout.addWidget(self.text_edit)

        btn = QPushButton("Run")
        btn.clicked.connect(self.get_table)
        mytext.addRow(btn)
        # layout.addWidget(btn)

        # btn = QPushButton("Launch")
        # btn.clicked.connect(self.launchDialog)
        # layout.addWidget(btn)

        self.textBrowser_msg = QTextBrowser()
        self.textBrowser_msg.setStyleSheet("font-size: 15px;")
        self.textBrowser_msg.setFixedHeight(100)
        mytext.addRow(self.textBrowser_msg)
        # layout.addWidget(self.textBrowser_msg)

        # ### USER
        self.user_id = QTextEdit()
        self.user_id.setStyleSheet("font-size: 20px;")
        self.user_id.setFixedHeight(50)
        self.user_id.setPlaceholderText(r"USER의 ID를 입력하세요")
        label = QLabel("USER  ")
        label.setFont(QFont("Arial", 10))
        mytext.addRow(label, self.user_id)
        # ### PW
        self.pw_id = QTextEdit()
        self.pw_id.setStyleSheet("font-size: 20px;")
        self.pw_id.setFixedHeight(50)
        self.pw_id.setPlaceholderText(r"USER의 PW를 입력하세요")
        label = QLabel("PASSWORD ")
        label.setFont(QFont("Arial", 10))
        mytext.addRow(label, self.pw_id)
        btn2 = QPushButton("제출 실행")
        btn2.clicked.connect(self.run_acc_selenium)
        mytext.addRow(btn2)

    def handleSavemon(self):
        #        with open('monschedule.csv', 'wb') as stream:
        with open("monschedule.csv", "w") as stream:  # 'w'
            writer = csv.writer(stream, lineterminator="\n")  # + , lineterminator='\n'
            for row in range(self.tablewidgetmon.rowCount()):
                rowdata = []
                for column in range(self.tablewidgetmon.columnCount()):
                    item = self.tablewidgetmon.item(row, column)
                    if item is not None:
                        #                        rowdata.append(unicode(item.text()).encode('utf8'))
                        rowdata.append(item.text())  # +
                    else:
                        rowdata.append("")

                writer.writerow(rowdata)

    def launchDialog(self):
        option = self.options.index(self.combo.currentText())
        if option == 0:
            response = self.getDirectory()
        elif option == 1:
            response = self.get_table()
        else:
            print("Got Nothing")

    def get_table(self):
        try:
            api_key = self.text_edit.toPlainText()
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

            report_sales_info = make_sales_info(result_table=result)
            report_sales_info.to_excel(writer2, index=False)
            writer2.close()
            n_day_of_unique = result["day"].nunique()
            event = ["휴가", "휴일"]
            n_day_of_not_event_unique = result.query("특이사항 not in @event")["day"].nunique()
            luncu_tag = ["'점심'", "'중식'"]
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
            self.textBrowser_msg.append("결과 테이블 저장 완료...")

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
            self.textBrowser_msg.clear()
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
            os.system(
                f"python ./acc_contents_selenium/acc_selenium.py -id {self.user_id.toPlainText().strip()} -pw {self.pw_id.toPlainText().strip()} -f {to_file_path}"
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
        self.textBrowser_dir.append(self._dir_name)

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
