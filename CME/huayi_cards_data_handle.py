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
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
import sys
import time
import re


''' 
提前的准备工作
1. 三个表复制数据列并清空数据，按照相同格式增加日期标注
2. 在《企业投放统计》表格中新增卡类型
3. 
'''

# 全局变量的定义及赋值
workspace_path: str = os.path.dirname(os.path.realpath(__file__))
today_date: str = str(datetime.date.today()).replace("-", "").replace('2019','19')


def download_data_xls() -> list:
    '''
    【功能】从数据统计后台中下载相应的华医学术卡数据原始记录表
    '''
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
    card_file_path: str = os.path.join(workspace_path, 'data', f'华医学术卡原始数据-{today_date}.xls')
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
    wb.SaveAs(xlsx_path, FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    # 源文件处理与新文件输出
    os.remove(xls_path)
    return xlsx_path


def only_chinese(content: str) -> str:
    '''
    【功能】将传入的文本仅保留中文
    '''
    # 处理前进行相关的处理，包括转换成Unicode等
    pattern = re.compile('[^\u4e00-\u9fa50-9]')  # 中文的编码范围是：\u4e00到\u9fa5
    zh_str: str = "".join(pattern.split(content))
    return zh_str



def statistic_data():
    '''
    【功能】进行华医学术卡数据统计分析
    '''


    # ------------------- 获取模板与数据文件并备份 -------------------
    # 载入表格
    data_xlsx_path: str = os.path.join(workspace_path, 'data', f'华医学术卡原始数据-{today_date}.xlsx')
    data_wb = load_workbook(data_xlsx_path)
    template_xlsx_path: str = os.path.join(workspace_path, '华医网学术卡数据周报.xlsx')
    template_wb = load_workbook(template_xlsx_path)

    # 备份并重命名《学习记录》原始数据表
    data_sht = data_wb.copy_worksheet(data_wb['sheet0'])
    data_wb['sheet0'].title = 'backup'
    data_sht.title = 'data'
    data_last_row: int = data_sht.max_row
    sum_numb: int = data_last_row - 1       # 总绑卡数

    # ------------------- 数据清洗 -------------------
    # 删除多余的列
    spare_cols_list: list = [15, 12, 10, 9, 6, 4, 3, 1]
    for spare_col_numb in spare_cols_list:
        data_sht.delete_cols(spare_col_numb)

    # 插入新列并获取绑卡的月份，小时
    data_sht.insert_cols(3, 2)
    for i in range(2, data_last_row + 1):
        data_sht.cell(i, 3).value = str(data_sht.cell(i, 2).value)[5:7]  # 月份
        data_sht.cell(i, 4).value = str(data_sht.cell(i, 2).value)[11:13] # 小时

    # 设置列名
    col_name_list: list = ['卡类型','绑定时间','月份','小时','省份','城市','单位级别','职称','专业']
    for i in range(len(col_name_list)):
        data_sht.cell(1, i+1).value = col_name_list[i]

    # 数据清洗
    content_col_dict: dict = {'省份':'5' , '城市':'6' , '单位级别':'7' , '职称':'8' , '专业':'9'}
    outlier_list: list = ['', '其它', '-请选择-', 'NULL', 'null']
    city_outlier_list: list = ['省直属单位', '省直辖县级行政单位', '省直辖县级行政区划']
    #TODO:检查dict是否有重复的列值，有问题则报错
    for i in range(2, data_last_row + 1):
        # 统一遍历清洗
        for k,v in content_col_dict.items():
            if not v == '':
                if not data_sht.cell(i, int(v)).value and data_sht.cell(i, int(v)).value != 0:
                    data_sht.cell(i, int(v)).value = '其他'
                else:
                    data_sht.cell(i, int(v)).value = only_chinese(data_sht.cell(i, int(v)).value) # 仅保留中文文本
                    if data_sht.cell(i, int(v)).value in outlier_list:
                        data_sht.cell(i, int(v)).value = '其他'
                    if k == '城市' and data_sht.cell(i, int(v)).value in city_outlier_list:
                        data_sht.cell(i, int(v)).value += data_sht.cell(i, int(content_col_dict['省份'])).value
                    if k == '省份' and data_sht.cell(i, int(v)).value == '黑龙江森林工业总局卫生局':
                        data_sht.cell(i, int(v)).value = '黑龙江森工'

    # ------------------- 数据统计 -------------------
    pro_card_dict: dict = {}
    pro_card: str = ''
    pro_dict: dict = {"广东省" : 0, "四川省" : 0, "海南省" : 0, "山西省" : 0, "江苏省" : 0, "山东省" : 0, "北京市" : 0, "江西省" : 0, "重庆市" : 0, "湖南省" : 0, "河南省" : 0, "福建省" : 0, "贵州省" : 0, "黑龙江省" : 0, "安徽省" : 0, "新疆维吾尔自治区" : 0, "吉林省" : 0, "河北省" : 0, "广西壮族自治区" : 0, "辽宁省" : 0, "浙江省" : 0, "陕西省" : 0, "宁夏回族自治区" : 0, "湖北省" : 0, "甘肃省" : 0, "青海省" : 0, "西藏自治区" : 0, "天津市" : 0, "上海市" : 0, "云南省" : 0, "内蒙古自治区" : 0, "香港特别行政区" : 0, "澳门特别行政区" : 0, "台湾省" : 0, "其他" : 0} 
    pro: str = ''
    city_dict: dict = {}
    hour_dict: dict = {"00": 0, "01": 0, "02": 0, "03": 0, "04": 0, "05": 0, "06": 0, "07": 0, "08": 0, "09": 0, "10": 0, "11": 0, "12": 0, "13": 0, "14": 0, "15": 0, "16": 0, "17": 0, "18": 0, "19": 0, "20": 0, "21": 0, "22": 0, "23": 0}
    month_dict: dict = {"02" : 0, "03" : 0, "04" : 0, "05" : 0, "06" : 0, "07" : 0, "08" : 0, "09" : 0, "10" : 0, "11" : 0, "12" : 0}
    hosp_dict: dict = {"三甲" : 0, "三乙" : 0, "二甲" : 0, "二乙" : 0, "一甲" : 0, "一乙" : 0, "其他" : 0}
    

    for i in range(2, data_last_row + 1):
        pro_card = data_sht.cell(i, 1).value + '#' + data_sht.cell(i, 5).value
        # 省份-卡类型字典统计
        pro_card_dict.setdefault(pro_card, 0)
        pro_card_dict[pro_card] += 1

        # 省份统计
        if '新疆' in data_sht.cell(i, 5).value:
            pro = '新疆维吾尔族自治区'           # 发卡区域合并：新疆生产建设兵团
        elif '黑龙江' in data_sht.cell(i, 5).value:
            pro = '黑龙江省'                    # 发卡区域合并：黑龙江森林工业总局卫生局、黑龙江农垦、黑龙江森工
        else:
            pro = data_sht.cell(i, 5).value
        pro_dict[pro] += 1

        # 城市统计
        city_dict.setdefault(data_sht.cell(i, 6).value, 0)
        city_dict[data_sht.cell(i, 6).value] += 1

        hour_dict[data_sht.cell(i, 4).value] += 1      # 小时统计        
        month_dict[data_sht.cell(i, 3).value] += 1      # 月份统计
        hosp_dict[data_sht.cell(i, 7).value] += 1      # 医院级别统计


    # ------------------- 数据写入 -------------------

    # 《Chart_Data》数据写入

    # 《省份分布表》数据写入
    pro_last_row: int = template_wb['省份分布'].max_row
    # ----- 写入备注
    notes_dict: dict = {}
    for i in range(2, pro_last_row):
        if template_wb['省份分布'].cell(i, 1).value in notes_dict.keys():
            template_wb['省份分布'].cell(i, 2).value = notes_dict[template_wb['省份分布'].cell(i, 1).value]

    # 《卡类状况表》数据写入

    # 《企业投放统计表》数据写入
    cpy_lastrow: int = template_wb['企业投放统计']



    # 表格的保存
    data_wb.save(data_xlsx_path)
    report_wb_path: str = os.path.join(workspace_path, 'history', f'华医网学术卡数据周报-{today_date}.xlsx')
    template_wb.save(report_wb_path)

    print("【STEP-10】数据写入文件保存\n\t\t [OK] --> 已经将数据写入到表格中！\n")


# ------------------------------ 主体调用部分 ------------------------------
def huayi_card_report():
    data_result: list = download_data_xls()
    if data_result[0] == 200:
        print(f'【STEP-1】爬虫下载原始数据\n\t\t[OK] --> 已经成功下载数据文件!\n')
        xlsx_path: str = xls_to_xlsx(data_result[1])
        print(f'【STEP-2】文件格式转换\n\t\t[OK] --> 已经将文件转化为xlsx格式!\n')

    else:
        print(f'【STEP-1】爬虫下载原始数据\n\t\t[ERROR] --> 未能从服务器中成功爬取数据:状态码为 {data_result[0]}\n')


huayi_card_report()
