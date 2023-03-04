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
import configparser
from datetime import datetime
from pathlib import Path
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from datetime import datetime


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
        self.tabs.resize(300, 200)

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
        self.event_text_edit.setPlaceholderText(r"휴가 일자를 넣어 주세요 1,휴가/2,휴가/")
        self.tab1.layout.addRow(QLabel("휴가   "), self.event_text_edit)

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

        self.tab1.layout.addRow(label, self.user_id)
        # ### PW

        self.pw_id = PasswordEdit()
        self.pw_id.setFixedHeight(50)

        label = QLabel("PASSWORD ")
        label.setFont(QFont("Arial", 10))
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
        [standardbriefscode.csv]

        위의 파일에 추가적으로 간소화된 표현을 작성하고,
        표준적요설명을 적어주시면 제출할 때 반영됩니다.
        
        (YYYYMMDD)+(TAG)+(PRICE)+(특이사항).jpg        
        예) {today}+점심+2_000+1명.jpg
        예) {today}+점심+50000+5명.jpg
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
        encoding = detect_encoding("./standardbriefscode.csv")
        MAPPING_TABLE = pd.read_csv("./standardbriefscode.csv", encoding=encoding)
        self.tableWidget.setRowCount(MAPPING_TABLE.shape[0])
        self.tableWidget.setColumnCount(MAPPING_TABLE.shape[1])
        for idx, one_row in MAPPING_TABLE.iterrows():
            for idx2, (_, v) in enumerate(one_row.to_dict().items()):
                v = QTableWidgetItem(str(v))
                v.setFont(font)
                self.tableWidget.setItem(idx, idx2, v)

        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.verticalScrollBar().setValue(0)
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # self.tableWidget.horizontalHeader().setFixedHeight(100)
        font = QFont()
        font.setPointSize(15)
        self.tableWidget.setHorizontalHeaderLabels(list(MAPPING_TABLE))
        # self.tableWidget.horizontalHeader().setDefaultSectionSize(15)

        self.tab2.layout.addRow(self.tableWidget)
        ##############
        self.tab2.setLayout(self.tab2.layout)

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
            self.textBrowser_msg.append(f"제출 파일은 확인 및 일부 수정이 필요합니다.")
            folder_path = str(Path(os.path.abspath(result_path)).parent)
            self.textBrowser_msg.append(f"해당 파일이 있는 경로를 엽니다. 파일을 확인해주세요")
            os.startfile(folder_path)
            return 200
        except Exception as e:
            # self.textBrowser_msg.clear()
            print(generate_error(e))
            self.textBrowser_msg.append("Error... 테이블 생성 에러...")
            self.textBrowser_msg.append(generate_error(e))
            return 404

    def run_acc_selenium(self):
        try:
            self.check_exist_img_directory()
            result_path2 = f"{self._dir_name}/sample_data.xlsx"
            # result_path2 = str(Path(result_path2)).replace("\\", "/").strip()
            #  -f {result_path2}
            import json
            import shutil

            to_file_path = "./acc_contents_selenium/sample_data.xlsx"
            shutil.copyfile(result_path2, to_file_path)
            chrom_driver_path = self.textBrowser_chrome_driver_msg.toPlainText()
            if Path(chrom_driver_path).joinpath("chromedriver.exe").is_file() is False:
                raise FileNotFoundError(f"해당 폴더에는 chromedriver.exe 존재하지 않습니다. {chrom_driver_path}")
            os.system(
                f"python ./acc_contents_selenium/acc_selenium.py -id {self.user_id.toPlainText().strip()} -pw {self.pw_id.text().strip()} -f {to_file_path} -c {chrom_driver_path}"
            )
        except Exception as e:
            print(generate_error(e))
            self.textBrowser_msg.append(generate_error(e))
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
