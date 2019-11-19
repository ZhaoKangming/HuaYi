# -*- coding:utf-8 -*-
'''
Function：处理华医学术卡的数据表，生成数据周报
Author: ZhaoKangming
E-mail: zhaokm0@gmail.com
Version: V0.1
'''
import urllib.request
import io
import requests
import shutil
import os
from sys import intern
import win32com.client as win32
import datetime
from openpyxl import *
from openpyxl.styles import Font, Border, Side, Alignment
import sys
import time


def download_data_xls() -> list:
    '''
    【功能】从数据统计后台中下载相应的华医学术卡数据原始记录表
    '''
    today_date: str = str(datetime.date.today()).replace("-", "")
    card_data_dict: dict = {'kind': 'card_yaoqi'}

    #-------------------------- 模拟登陆 --------------------------
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')  # 改变标准输出的默认编码
    data = {'login_name': 'yaoqi', 'login_pwd': 1}  # 登录时需要POST的数据
    login_url: str = 'http://192.168.1.240:8122/Login/Login'  # 登录时表单提交到的地址
    session = requests.Session()
    resp = session.post(login_url, data)

    #-------------------------- 请求数据 --------------------------
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
                'Accept-Encoding': 'gzip, deflate',
                'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
                'Cache-Control': 'max-age=0',
                'Connection': 'keep-alive',
                'Content-Type': 'application/x-www-form-urlencoded',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}
    download_url: str = 'http://192.168.1.240:8122/Common/ToExcel'  # 下载POST地址

    download_state_list: list = []
    card_download_resp = session.post(download_url, card_data_dict, headers = headers)
    card_file_path: str = os.path.join(os.path.dirname(os.path.realpath(__file__)), f'data\\华医学术卡数据-{today_date}.xls')
    with open(card_file_path, "wb") as card_downloaded_file:
        card_downloaded_file.write(card_download_resp.content)
    download_state_list.append(card_download_resp.status_code)
    download_state_list.append(card_file_path)

    return download_state_list


def xls_to_xlsx(xls_path: str) -> str:
    '''
    【功能】将 xls 文件转化为 xlsx 文件
    :param xls_path: xls文件的路径
    '''
    # 文件格式转化
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(xls_path)
    xlsx_path: str = xls_path + 'x'
    # FileFormat = 51 is for .xlsx extension, FileFormat = 56 is for .xls extension
    wb.SaveAs(xlsx_path, FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    # 源文件处理与新文件输出
    os.remove(xls_path)
    return xlsx_path


# 删除多余的列
del_cols_list = [15, 12, 10, 9, 6, 4, 3, 1]

# 插入新列

# 设置列名

# 数据清洗
省直辖县级行政单位、省直辖县级行政区划

-请选择-


