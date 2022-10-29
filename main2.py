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

import csv
from datetime import datetime

import calendar
import numpy as np
import itertools
from datetime import datetime, date
import json
from pandas import json_normalize
from pathlib import Path
import pandas as pd

__all__ = ["get_sales_check_of_credit_card"]


def get_holiday(api_key):
    import requests

    today = datetime.today().strftime("%Y%m%d")
    today_year = datetime.today().year
    # key = input("API 키를 입력해주세요 (https://www.data.go.kr/data/15012690/openapi.do)")
    key = api_key
    url = (
        "http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo?_type=json&numOfRows=50&solYear="
        + str(today_year)
        + "&ServiceKey="
        + str(key)
    )
    response = requests.get(url)

    if response.status_code == 200:
        json_ob = json.loads(response.text)
        holidays_data = json_ob["response"]["body"]["items"]["item"]
        dataframe = json_normalize(holidays_data)
    dataframe["locdate"] = pd.to_datetime(dataframe["locdate"], format="%Y%m%d")
    return dataframe


def check_holiday(result_table, api_key):

    holiday_df = get_holiday(api_key)

    check = list(set(result_table["day"]) & set(holiday_df["locdate"]))
    result_table[result_table.day.isin(check)] = result_table[result_table.day.isin(check)].fillna(
        dict(태그="-", 총합=0, 특이사항="휴일")
    )
    return result_table


from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *


def get_sales_check_of_credit_card(path, year, month, api_key):
    path = rf"{path}"
    folder = Path(path)
    result = []
    print(list(folder.glob("*.PNG")))
    max_col = 0
    for i in list(folder.glob("*.PNG")):
        result.append([i.replace("_", "") for i in i.parts[-1].split(".")[0].split("+")])
        max_col = max([max_col, len(result[-1])])
    else:
        if max_col == 3:
            result = pd.DataFrame(result, columns=["날짜", "태그", "가격"])
            result["비고"] = "-"
        elif max_col == 4:
            result = pd.DataFrame(result, columns=["날짜", "태그", "가격", "비고"])
        result["날짜"] = pd.to_datetime(result["날짜"])
        result["가격"] = result["가격"].astype(int)
        result["비고"] = result["비고"].fillna("-")

    day = calendar.monthcalendar(year, month)
    dayofweek = np.array(day)[:, -2]
    check_bool = np.where(dayofweek != 0, True, False)

    dayofweek_day = np.array(list(itertools.repeat(list("월화수목금"), dayofweek.shape[0])))[check_bool].ravel()
    dayofweek = dayofweek[check_bool].ravel()
    total_date = pd.DataFrame(dict(day=[date(year, month, i) for i in dayofweek], dayofweek=dayofweek_day))
    total_date["day"] = pd.to_datetime(total_date["day"])

    result2 = result.groupby(["날짜", "태그"], as_index=False).apply(
        lambda x: pd.Series(dict(총합=sum(x["가격"]), 특이사항=x["비고"].tolist()))
    )
    result_table = pd.merge(total_date, result2, left_on="day", right_on="날짜", how="outer").drop(columns=["날짜"])
    result_table = check_holiday(result_table, api_key)
    return result_table


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
        self.year_options = [str(i) for i in list(range(2022, 2025))]

        self.combo_year = QComboBox()
        self.combo_year.addItems(self.year_options)
        self.combo_month = QComboBox()
        self.combo_month.addItems(self.month_options)
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
        self.textBrowser_msg.setStyleSheet("font-size: 10px;")
        self.textBrowser_msg.setFixedHeight(100)
        mytext.addRow(self.textBrowser_msg)
        # layout.addWidget(self.textBrowser_msg)

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
        api_key = self.text_edit.toPlainText()
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
        writer = pd.ExcelWriter(result_path, engine="xlsxwriter")

        result.to_excel(writer, sheet_name="식대")

        n_day_of_unique = result["day"].nunique()
        event = ["휴가", "휴일"]
        n_day_of_not_event_unique = result.query("특이사항 not in @event")["day"].nunique()
        lunch_total = result.query("태그 == '점심'")["총합"].sum()
        lunch_event = pd.DataFrame(
            [{"총근무일수": n_day_of_unique, "총근무일수(이벤트제외)": n_day_of_not_event_unique, "현재까지 점심 사용 금액": lunch_total}]
        )
        lunch_event.to_excel(writer, sheet_name="점심식대")
        event_total_df = result.groupby(["태그"], as_index=False).sum()
        event_total_df.to_excel(writer, sheet_name="총태그합")
        writer.close()

        self.textBrowser_msg.clear()
        self.textBrowser_msg.append("결과 테이블 저장 완료...")
        self.textBrowser_msg.append(f"파일 위치 {result_path}")
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
        print(response)
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
        print(response)
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
