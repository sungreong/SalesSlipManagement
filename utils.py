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


def get_sales_check_of_credit_card(path, year, month, api_key):
    path = rf"{path}"
    folder = Path(path)
    result = []
    for i in list(folder.glob("*.PNG")):
        result.append([i.replace("_", "") for i in i.parts[-1].split(".")[0].split("+")])
    else:
        result = pd.DataFrame(result, columns=["날짜", "태그", "가격", "비고"])
        result["날짜"] = pd.to_datetime(result["날짜"])
        result["가격"] = result["가격"].astype(int)
        result["비고"] = result["비고"].fillna("-")
    print(result)

    day = calendar.monthcalendar(year, month)
    dayofweek = np.array(day)[:, :-2]
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
