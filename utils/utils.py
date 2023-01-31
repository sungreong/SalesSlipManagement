from datetime import datetime, date


import json
from pandas import json_normalize
import pandas as pd

import calendar
import numpy as np
import itertools
from pathlib import Path
import re
from enum import Enum
import os, sys

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.append(root_dir)
from utils.metadata import CODE_MAP_TABLE

# class syntax
class DayofWeek(Enum):
    월 = 0
    화 = 1
    수 = 2
    목 = 3
    금 = 4
    토 = 5
    일 = 6


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


"test.test2+test3".split(".|+")
import re


def split_text(text):
    text = text.parts[-1].split(".")[0]
    return re.split("\.|\+", text)


def get_sales_check_of_credit_card(path, year, month, api_key):
    path = rf"{path}"
    folder = Path(path)
    result = []
    img_format = "*[PNG$|jpg$|png$]"
    print(folder.glob(img_format))
    print(list(folder.glob(img_format)))
    assert len(list(folder.glob(img_format))) != 0, "파일 인식 문제 발생"
    max_col = 0
    for i in list(folder.glob(img_format)):
        result.append([i.replace("_", "") for i in split_text(i)])
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
    dayofweek = np.array(day)[:, :-2]
    check_bool = np.where(dayofweek != 0, True, False)

    dayofweek_day = np.array(list(itertools.repeat(list("월화수목금"), dayofweek.shape[0])))[check_bool].ravel()
    dayofweek = dayofweek[check_bool].ravel()
    try:
        total_date = pd.DataFrame(dict(day=[date(year, month, i) for i in dayofweek], dayofweek=dayofweek_day))
    except Exception as e:
        print(result.to_markdown())
        print(dayofweek)
        raise Exception(e)
    total_date["day"] = pd.to_datetime(total_date["day"])

    result2 = result.groupby(["날짜", "태그", "비고"], as_index=False).apply(
        lambda x: pd.Series(dict(총합=sum(x["가격"]), 특이사항=x["비고"].tolist()))
    )
    for i in list(set(result2["날짜"]).difference(set(total_date["day"]))):
        total_date = total_date.append({"day": i, "dayofweek": DayofWeek(i.weekday()).name}, ignore_index=True)

    result_table = pd.merge(total_date, result2, left_on="day", right_on="날짜", how="outer").drop(columns=["날짜"])
    result_table = check_holiday(result_table, api_key)
    return result_table


def make_sales_info(result_table: pd.DataFrame):
    result_table = result_table.query("총합 > 0")
    # tmp_replace_dict = {
    #     "저녁": "석식",
    #     "점심": "중식",
    #     "통신요금": "통신비",
    # }
    print(result_table)
    result_table["태그"] = result_table["태그"].replace(CODE_MAP_TABLE)
    result_table["day"] = result_table["day"].dt.strftime("%Y%m%d").astype(int)
    result_table["howmany"] = "1명"  # TODO:
    result_table = result_table[["day", "태그", "howmany", "총합", "특이사항"]]
    result_table.columns = ["date", "type", "howmany", "amount", "etc"]
    result_table["etc"] = result_table["etc"].apply(lambda x: re.sub(r"[^\uAC00-\uD7A30-9a-zA-Z\s]", "", str(x)))
    return result_table
