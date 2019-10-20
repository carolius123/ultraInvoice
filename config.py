#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : config.py
# @Author: Sui Huafeng
# @Date  : 2019/7/25
# @Desc  : 全局变量和公共函数
#

import json
import logging
import os
import sys
import time
from datetime import date, timedelta
from shutil import rmtree

import boto3
import requests


def nextMonth(yy_mm):
    return yy_mm[:2] + '%02d' % (int(yy_mm[2:]) + 1) if yy_mm[2:] < '12' else '%02d01' % (int(yy_mm[:2]) + 1)


# 初始化日志设置
def initLogger(level, logPath):
    log = logging.getLogger('ultraInvoice')
    LogLevel = {'NOTSET': 0, 'DEBUG': 10, 'INFO': 20, 'WARNING': 30, 'ERROR': 40, 'CRITICAL': 50}
    logLevel = LogLevel.get(level.upper(), '20')
    log.setLevel(logLevel)
    formatter = logging.Formatter('%(asctime)s\t%(levelname)s\t%(filename)s %(lineno)d\t%(message)s')
    logFile = os.path.join(logPath, 'ultraInvoice.log')
    fh = logging.FileHandler(logFile, encoding='GBK')
    fh.setFormatter(formatter)
    log.addHandler(fh)
    fh = logging.StreamHandler()
    fh.setFormatter(formatter)
    log.addHandler(fh)
    log.debug("stared!")
    return log


# 从S3下载Credit余额，报表模板等配置文件
def downloadMetaData():
    s3 = boto3.resource('s3')
    bucket = s3.Bucket(Bucket)
    for obj in bucket.objects.filter(Prefix='Meta/'):
        if obj.key[-1] == '/':
            continue
        file_name = os.path.join(MetaPath, obj.key.split('/')[-1])
        log.info('Downloading %s' % obj.key)
        bucket.download_file(obj.key, file_name)


Bucket = 'billing-up-972221870813'
Argv = {p.split('=')[0]: p.split('=')[-1] for p in sys.argv[1:]}
BillMonth = Argv.get('month', (date(date.today().year, date.today().month, 1) - timedelta(days=10)).strftime('%y%m'))

WorkPath = Argv.get('path', '/tmp')
MetaPath = os.path.join(WorkPath, 'meta')  # 保存阶梯价格、Credit等配置和消费日志数据
TmpPath = os.path.join(WorkPath, 'tmp')  # 保存中间结果
if os.path.exists(TmpPath):
    rmtree(TmpPath)
    time.sleep(1)  # 防止立刻建立目录出错
[os.mkdir(folder) for folder in (WorkPath, MetaPath, TmpPath) if not os.path.exists(folder)]
log = initLogger(Argv.get('log', 'INFO'), TmpPath)

downloadMetaData()

with open(os.path.join(MetaPath, 'Payers.json'), 'r', encoding='utf8') as fp:
    payers = json.load(fp)
with open(os.path.join(MetaPath, 'customers.json'), 'r', encoding='utf8') as fp:
    Customers = json.load(fp)

CustomerNames = {}
for customer_name, customer in Customers.items():
    for awsAccount in customer['Accounts']:
        CustomerNames[awsAccount] = customer_name

"""
libs
"""


# 多级字典中，增加一行的key和默认值
def cascadeDictDefault(dictionary, key_defaults):
    upper = dictionary
    for key, default in key_defaults:
        if key not in dictionary.keys():
            dictionary[key] = default
        upper = dictionary
        dictionary = dictionary[key]
    return dictionary


# 从中行官网爬取当月1日16点左右的美元现汇卖出价
def getUsdExchangeRate(yy_mm):
    try:
        first_work_day = date(year=int('20' + yy_mm[:2]), month=int(yy_mm[2:]), day=1)
        days = first_work_day.weekday() - 4
        if days > 0:
            first_work_day += timedelta(days=3 - days)

        url_parameters = {'erectDate': first_work_day.strftime('%Y-%m-%d 15:00:00'),
                          'nothing': first_work_day.strftime('%Y-%m-%d 16:00:00'),
                          'pjname': 1316, 'page': 1}
        url = 'http://srh.bankofchina.com/search/whpj/search.jsp?%s' \
              % ('&'.join(['%s=%s' % (k, v) for k, v in url_parameters.items()]))
        html = requests.get(url).text.split('\r\n')
        header = ['<select name="pjname" id="pjname">', '<th>货币名称</th>']
        idx = 0
        for lines, line in enumerate(html):  # 找到不连续多行的特征序列
            if header[idx] in line:
                idx += 1
                if idx == len(header):
                    break
        else:
            return 0

        for spacing, line in enumerate(html[lines:]):  # 找出表格所处列数，也就是跟相隔行数
            if '<th>现汇卖出价</th>' in line:
                break
        else:
            return 0

        spacing1 = 0
        lines += spacing
        for line in html[lines:]:  # 按相同列数定位数据行位置
            if '<td>美元</td>' not in line and spacing1 == 0:
                continue
            spacing1 += 1
            if spacing1 > spacing:
                break
        else:
            return 0
    except Exception as err:
        log.exception(err)
        return 0
    exchange_rate = float(line.split('<td>')[-1].split('</td>')[0]) / 100
    log.info('Get USD exchange rate from BOC website：%.3f' % exchange_rate)
    return exchange_rate
