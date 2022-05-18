#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
Date: 2022/2/24 14:50
Desc: 真气网-空气质量
https://www.zq12369.com/environment.php
空气质量在线监测分析平台的空气质量数据
https://www.aqistudy.cn/
"""
import getopt
import json
import os
import re
import sys

import numpy as np
import pandas as pd
import requests
from akshare.utils import demjson
from openpyxl.utils import get_column_letter
from py_mini_racer import py_mini_racer


def _get_js_path(name: str = None, module_file: str = None) -> str:
    """
    获取 JS 文件的路径(从模块所在目录查找)
    :param name: 文件名
    :type name: str
    :param module_file: 模块路径
    :type module_file: str
    :return: 路径
    :rtype: str
    """
    module_folder = os.path.abspath(
        os.path.dirname(os.path.dirname(module_file)))
    module_json_path = os.path.join(module_folder, "AqSpider", name)
    return module_json_path


def _get_file_content(file_name: str = "crypto.js") -> str:
    """
    获取 JS 文件的内容
    :param file_name:  JS 文件名
    :type file_name: str
    :return: 文件内容
    :rtype: str
    """
    setting_file_name = file_name
    setting_file_path = _get_js_path(setting_file_name, __file__)
    with open(setting_file_path) as f:
        file_data = f.read()
    return file_data


def has_month_data(href):
    """
    Deal with href node
    :param href: href
    :type href: str
    :return: href result
    :rtype: str
    """
    return href and re.compile("monthdata.php").search(href)


def air_city_table() -> pd.DataFrame:
    """
    真气网-空气质量历史数据查询-全部城市列表
    https://www.zq12369.com/environment.php?date=2019-06-05&tab=rank&order=DESC&type=DAY#rank
    :return: 城市映射
    :rtype: pandas.DataFrame
    """
    url = "https://www.zq12369.com/environment.php"
    date = "2020-05-01"
    if len(date.split("-")) == 3:
        params = {
            "date": date,
            "tab": "rank",
            "order": "DESC",
            "type": "DAY",
        }
        r = requests.get(url, params=params)
        temp_df = pd.read_html(r.text)[1].iloc[1:, :]
        del temp_df['降序']
        temp_df.reset_index(inplace=True)
        temp_df['index'] = temp_df.index + 1
        temp_df.columns = ['序号', '省份', '城市', 'AQI', '空气质量', 'PM2.5浓度', '首要污染物']
        temp_df['AQI'] = pd.to_numeric(temp_df['AQI'])
    return temp_df


def air_quality_watch_point(city: str = "杭州",
                            start_date: str = "20220408",
                            end_date: str = "20220409") -> pd.DataFrame:
    """
    真气网-监测点空气质量-细化到具体城市的每个监测点
    指定之间段之间的空气质量数据
    https://www.zq12369.com/
    :param city: 调用 ak.air_city_table() 接口获取
    :type city: str
    :param start_date: e.g., "20190327"
    :type start_date: str
    :param end_date: e.g., ""20200327""
    :type end_date: str
    :return: 指定城市指定日期区间的观测点空气质量
    :rtype: pandas.DataFrame
    """
    start_date = "-".join([start_date[:4], start_date[4:6], start_date[6:]])
    end_date = "-".join([end_date[:4], end_date[4:6], end_date[6:]])
    url = "https://www.zq12369.com/api/zhenqiapi.php"
    file_data = _get_file_content(file_name="crypto.js")
    ctx = py_mini_racer.MiniRacer()
    ctx.eval(file_data)
    method = "GETCITYPOINTAVG"
    ctx.call("encode_param", method)
    ctx.call("encode_param", start_date)
    ctx.call("encode_param", end_date)
    city_param = ctx.call("encode_param", city)
    ctx.call("encode_secret", method, city_param, start_date, end_date)
    payload = {
        "appId":
        "a01901d3caba1f362d69474674ce477f",
        "method":
        ctx.call("encode_param", method),
        "city":
        city_param,
        "startTime":
        ctx.call("encode_param", start_date),
        "endTime":
        ctx.call("encode_param", end_date),
        "secret":
        ctx.call("encode_secret", method, city_param, start_date, end_date),
    }
    headers = {
        "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36"
    }
    r = requests.post(url, data=payload, headers=headers)
    data_text = r.text
    data_json = demjson.decode(ctx.call("decode_result", data_text))
    temp_df = pd.DataFrame(data_json["rows"])
    return temp_df


def air_quality_hist(city: str = "杭州",
                     period: str = "day",
                     start_date: str = "20190327",
                     end_date: str = "20200427",
                     column: list = []) -> pd.DataFrame:
    """
    真气网-空气历史数据
    https://www.zq12369.com/
    :param city: 调用 ak.air_city_table() 接口获取所有城市列表
    :type city: str
    :param period: "hour": 每小时一个数据, 由于数据量比较大, 下载较慢; "day": 每天一个数据; "month": 每个月一个数据
    :type period: str
    :param start_date: e.g., "20190327"
    :type start_date: str
    :param end_date: e.g., "20200327"
    :type end_date: str
    :return: 指定城市和数据频率下在指定时间段内的空气质量数据
    :rtype: pandas.DataFrame
    """
    start_date = "-".join([start_date[:4], start_date[4:6], start_date[6:]])
    end_date = "-".join([end_date[:4], end_date[4:6], end_date[6:]])
    url = "https://www.zq12369.com/api/newzhenqiapi.php"
    file_data = _get_file_content(file_name="outcrypto.js")
    ctx = py_mini_racer.MiniRacer()
    ctx.eval(file_data)
    appId = "4f0e3a273d547ce6b7147bfa7ceb4b6e"
    method = "CETCITYPERIOD"
    timestamp = ctx.eval("timestamp = new Date().getTime()")
    p_text = json.dumps(
        {
            "city": city,
            "endTime": f"{end_date} 23:45:39",
            "startTime": f"{start_date} 00:00:00",
            "type": period.upper(),
        },
        ensure_ascii=False,
        indent=None,
    ).replace(' "', '"')
    secret = ctx.call("hex_md5",
                      appId + method + str(timestamp) + "WEB" + p_text)
    payload = {
        "appId": "4f0e3a273d547ce6b7147bfa7ceb4b6e",
        "method": "CETCITYPERIOD",
        "timestamp": int(timestamp),
        "clienttype": "WEB",
        "object": {
            "city": city,
            "type": period.upper(),
            "startTime": f"{start_date} 00:00:00",
            "endTime": f"{end_date} 23:45:39",
        },
        "secret": secret,
    }
    need = (json.dumps(payload,
                       ensure_ascii=False,
                       indent=None,
                       sort_keys=False).replace(' "', '"').replace(
                           "\\", "").replace('p": ',
                                             'p":').replace('t": ', 't":'))

    headers = {
        # 'Accept': '*/*',
        # 'Accept-Encoding': 'gzip, deflate, br',
        # 'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        # 'Cache-Control': 'no-cache',
        # 'Connection': 'keep-alive',
        # 'Content-Length': '1174',
        # 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        # 'Cookie': 'UM_distinctid=1800e5142c5b85-04b8f11aa852f3-1a343370-1fa400-1800e5142c6b7e; CNZZDATA1254317176=1502593570-1649496979-%7C1649507817; city=%E6%9D%AD%E5%B7%9E; SECKEY_ABVK=eSrbUhd28Mjo7jf8Rfh+uY5E9C+tAhQ8mOfYJHSjSfY%3D; BMAP_SECKEY=N5fGcwdWpeJW46eZRpR9GW3qdVnODGQwGm6JE0ELECQHJOTFc9MCuNdyf8OWUspFI6Xq4MMPxgVVr5I13odFOW6AQMgSPOtEvVHciC2NsQwb1pnmFtEaqyKHOUeavelt0ejBy6ETRD_4FXAhZb9FSbVIMPew7qwFX_kdPDxVJH-vHfCVhRx9XDZgb41B_T4D',
        # 'Host': 'www.zq12369.com',
        # 'Origin': 'https://www.zq12369.com',
        # 'Pragma': 'no-cache',
        # 'Referer': 'https://www.zq12369.com/environment.php?catid=4',
        # 'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Google Chrome";v="100"',
        # 'sec-ch-ua-mobile': '?0',
        # 'sec-ch-ua-platform': '"Windows"',
        # 'Sec-Fetch-Dest': 'empty',
        # 'Sec-Fetch-Mode': 'cors',
        # 'Sec-Fetch-Site': 'same-origin',
        'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }
    params = {"param": ctx.call("AES.encrypt", need)}
    params = {"param": ctx.call("encode_param", need)}

    r = requests.post(url, data=params, headers=headers)
    temp_text = ctx.call("decryptData", r.text)
    data_json = demjson.decode(ctx.call("b.decode", temp_text))
    print(data_json)
    temp_df = pd.DataFrame(data_json["result"]["data"]["rows"])
    temp_df.index = temp_df.index + 1
    # del temp_df["time"]
    # temp_df = temp_df.astype(float, errors="ignore")
    temp_df = pd.DataFrame(temp_df, columns=['time'] + column)
    return temp_df


def air_quality_rank(date: str = "") -> pd.DataFrame:
    """
    真气网-168 城市 AQI 排行榜
    https://www.zq12369.com/environment.php?date=2020-03-12&tab=rank&order=DESC&type=DAY#rank
    :param date: "": 当前时刻空气质量排名; "20200312": 当日空气质量排名; "202003": 当月空气质量排名; "2019": 当年空气质量排名;
    :type date: str
    :return: 指定 date 类型的空气质量排名数据
    :rtype: pandas.DataFrame
    """
    if len(date) == 4:
        date = date
    elif len(date) == 6:
        date = "-".join([date[:4], date[4:6]])
    elif date == '':
        date = '实时'
    else:
        date = "-".join([date[:4], date[4:6], date[6:]])

    url = "https://www.zq12369.com/environment.php"

    if len(date.split("-")) == 3:
        params = {
            "date": date,
            "tab": "rank",
            "order": "DESC",
            "type": "DAY",
        }
        r = requests.get(url, params=params)
        return pd.read_html(r.text)[1].iloc[1:, :]
    elif len(date.split("-")) == 2:
        params = {
            "month": date,
            "tab": "rank",
            "order": "DESC",
            "type": "MONTH",
        }
        r = requests.get(url, params=params)
        return pd.read_html(r.text)[2].iloc[1:, :]
    elif len(date.split("-")) == 1 and date != "实时":
        params = {
            "year": date,
            "tab": "rank",
            "order": "DESC",
            "type": "YEAR",
        }
        r = requests.get(url, params=params)
        return pd.read_html(r.text)[3].iloc[1:, :]
    if date == "实时":
        params = {
            "tab": "rank",
            "order": "DESC",
            "type": "MONTH",
        }
        r = requests.get(url, params=params)
        return pd.read_html(r.text)[0].iloc[1:, :]


def createFile(city, start_data, end_data, period) -> str:
    folderPath = 'D:\\air_quality_results'
    isFolderExist = os.path.exists(folderPath)
    if not isFolderExist:
        os.makedirs(folderPath)

    filePath = folderPath + '\\' + city + '_' + start_data + '_' + end_data + '_' + period + '.xlsx'
    return filePath


def write_to_excel(path, aq_df):
    writer = pd.ExcelWriter(path)
    aq_df.to_excel(writer, 'air_quality')

    # 设置自适应宽度
    column_widths = (
        aq_df.columns.to_series().apply(lambda x: len(x.encode('gbk'))).values)
    max_widths = (aq_df.astype(str).applymap(
        lambda x: len(x.encode('gbk'))).agg(max).values)
    widths = np.max([column_widths, max_widths], axis=0)
    worksheet = writer.sheets['air_quality']
    for i, width in enumerate(widths, 1):
        worksheet.column_dimensions[get_column_letter(i + 1)].width = width + 4

    writer.save()


def getOpt():
    city = '北京'
    period = 'day'
    start_date = '20220501'
    end_date = '202205010'
    column = [
        'aqi', 'pm2_5', 'pm10', 'co', 'no2', 'o3', 'so2', 'complexindex',
        'rank', 'primary_pollutant', 'temp', 'humi', 'windlevel',
        'winddirection', 'weather'
    ]

    opts, args = getopt.getopt(sys.argv[1:], 'c:s:e:p:l:')
    for o, a in opts:
        if o in ('-c'):
            city = a
            continue
        if o in ('-s'):
            start_date = a
            continue
        if o in ('-e'):
            end_date = a
            continue
        if o in ('-p'):
            period = a
            continue
        if o in ('-l'):
            column = a.replace(' ', '').split(',')
            continue

    # 获取数据
    print(city, start_date, end_date, period, column)
    aq_df = air_quality_hist(city=city,
                             start_date=start_date,
                             end_date=end_date,
                             period=period,
                             column=column)
    print('获取数据完成：')
    print(aq_df)

    # 创建文件
    path = createFile(city, start_date, end_date, period)
    print('创建文件完成: ', path)

    # 写入excel
    write_to_excel(path, aq_df)

    # 打开文件
    os.startfile(path)


if __name__ == "__main__":
    getOpt()
