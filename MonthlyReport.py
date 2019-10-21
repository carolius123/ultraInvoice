#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : MonthlyReport.py
# @Author: Sui Huafeng
# @Date  : 2019/10/16
# @Desc  : 以代付账户的月度报告MR为数据源，生成客户收据和账单
#  1. 按AWS账户拆分每个payer account的MR
#  2. 修正Credit，只扣减本账号自己的(Payer账号代充共享的扣掉，不管子账号余额)
# 按AWS账户和客户拆分PayerAccount-aws-billing-csv-yyyy-mm.csv
#
import json
import os
import boto3
import csv
import config as cfg
from XlsBill import XlsBill
from Mail import Mail


class Mr(object):
    SumCols = {'UsageQuantity', 'CostBeforeTax', 'Credits', 'TaxAmount', 'TotalCost'}
    CsvWriters = {}
    Receiver = os.environ.get('receiver', 'maling1@ultrapower.com.cn,jiayf@ultrapower.com.cn').split(',')

    # 按aws账户拆账
    def split(self, payerAccount):
        try:
            fileFullname = downloadAwsMR(payerAccount)
        except Exception as err:
            cfg.log.warning('[%s]downloadAwsMR：%s' % (payerAccount, str(err)))
            return

        accountTotal = getAccountTotal(fileFullname)
        self.__writeItems(fileFullname, accountTotal, payerAccount)

        for linkedAccountId, rowTotal in accountTotal.items():
            customerName = cfg.CustomerNames.get(linkedAccountId, None)
            if not customerName:  # 账号不在配置文件中
                cfg.log.warning('no customer for %s, please update customer.json.' % linkedAccountId)
                continue
            self.__writeSupport(rowTotal)
            self.__writeDiscount(rowTotal, linkedAccountId)
            fileName = mrFileName(rowTotal)
            self.CsvWriters[fileName][0].writerow(rowTotal)
        self.__close()

    # 拆分行写入csv文件
    def __writeItems(self, file_fullname, accountTotal, payerAccount):
        removeCredit = cfg.payers.get(payerAccount, {}).get('removeCredit', False)
        with open(file_fullname, 'r', encoding='utf-8', newline='') as fp:
            for row in csv.DictReader(fp):
                if row['RecordType'] != 'LinkedLineItem':  # 只拆分资源消费记录
                    continue
                if removeCredit:  # 去掉Credit抵扣，按泰岳折扣算
                    row['Credits'] = '0'
                    row['TotalCost'] = row['CostBeforeTax']
                for col in self.SumCols:
                    accountTotal[row['LinkedAccountId']][col] += float(row[col])
                fileName = mrFileName(row)
                self.__open(fileName, row) if fileName not in self.CsvWriters else None
                self.CsvWriters[fileName][0].writerow(row)

    # 添加泰岳折扣行，并更新rowTotal
    def __writeDiscount(self, rowTotal, linkedAccountId):
        linkedAccount = cfg.Customers[cfg.CustomerNames[linkedAccountId]]['Accounts'][linkedAccountId]
        discountRate = linkedAccount.get('Discount', 0)
        if not discountRate:
            return
        rowDiscount = rowTotal.copy()
        rowDiscount['RecordType'] = 'LinkedLineItem'
        rowDiscount['RecordID'] = ''
        rowDiscount['ItemDescription'] = '%g%% Discount offered by Ultrapower' % discountRate
        rowDiscount['ProductName'] = rowDiscount['UsageType'] = 'Ultrapower discount'
        rowDiscount['ProductCode'] = 'UltraDiscount'
        for col in self.SumCols:
            rowDiscount[col] = '0'

        discount = - rowTotal['TotalCost'] * discountRate / 100
        rowTotal['TotalCost'] += discount  # 更新汇总行
        rowDiscount['UsageQuantity'] = str(rowTotal['TotalCost'])
        rowDiscount['TotalCost'] = str(discount)

        fileName = mrFileName(rowDiscount)
        self.CsvWriters[fileName][0].writerow(rowDiscount)

    # 添加aws Support, 并更新rowTotal
    def __writeSupport(self, rowTotal):
        rowSupport = rowTotal.copy()
        rowSupport['RecordType'] = 'LinkedLineItem'
        rowSupport['ProductCode'] = 'OCBPremiumSupport'
        rowSupport['ProductName'] = 'AWS Premium Support'
        rowSupport['ItemDescription'] = 'AWS Support'
        for col in self.SumCols:
            rowSupport[col] = '0'
        # TBD
        return

    # 新建输出文件
    def __open(self, file_name, row):
        fileFullname = os.path.join(cfg.TmpPath, file_name)
        exists = os.path.exists(fileFullname)
        fp = open(fileFullname, 'a', encoding='utf-8', newline='')
        self.CsvWriters[file_name] = (csv.DictWriter(fp, fieldnames=row.keys()), fp)
        if not exists:
            self.CsvWriters[file_name][0].writeheader()

    def __close(self):
        for _, fp in self.CsvWriters.values():
            fp.close()
        self.CsvWriters = {}

    # 汇聚成客户的总账单
    def merge(self, customerName):
        consolidateRows = {}
        accounts = set()  # 实际产生费用的账户
        fileName = mrFileName('', customerName)
        fileFullname = os.path.join(cfg.TmpPath, fileName)
        row = None
        with open(fileFullname, 'r', encoding='utf-8', newline='') as fp:
            for row in csv.DictReader(fp):
                if row['RecordType'] != 'LinkedLineItem':
                    continue
                accounts.add(row['LinkedAccountId'])
                key = (row['ProductCode'], row['UsageType'])
                for col in self.SumCols: row[col] = float(row[col])
                if key not in consolidateRows:
                    consolidateRows[key] = row
                else:
                    for col in self.SumCols: consolidateRows[key][col] += row[col]

        if len(accounts) < 2:
            return  # 只有一个有费用账号
        with open(fileFullname, 'a', encoding='utf-8', newline='') as fp:
            csvWriter = csv.DictWriter(fp, fieldnames=row.keys())
            statementTotalRow = getStatementTotal(row.copy())
            for consolidateRow in consolidateRows.values():
                consolidateRow['RecordType'] = 'PayerLineItem'
                consolidateRow['LinkedAccountId'] = consolidateRow['LinkedAccountName'] = ''
                for col in self.SumCols:
                    statementTotalRow[col] += consolidateRow[col]
                    consolidateRow[col] = str(consolidateRow[col])
                csvWriter.writerow(consolidateRow)

            for col in self.SumCols:
                statementTotalRow[col] = str(statementTotalRow[col])
            csvWriter.writerow(statementTotalRow)

    # 汇总成靠近AWS系统内形式的分层的的月账单
    def bill(self, customerName):
        monthBills = {}
        fileName = mrFileName('', customerName)
        fileFullname = os.path.join(cfg.TmpPath, fileName)
        with open(fileFullname, 'r', encoding='utf-8', newline='') as fp:
            [add(monthBills, row) for row in csv.DictReader(fp) if row['RecordType'][-8:] == 'LineItem']

        monthBill = monthBills[''] if '' in monthBills else list(monthBills.values())[0]
        invoice = [(product, round(v['TotalCost'], 2))
                   for product, v in monthBill.items() if product != 'TotalCost' and abs(v['TotalCost']) > 0.005]
        invoice.sort(key=lambda v: abs(v[1]), reverse=True)
        if len(invoice) > 15:  # 最多保留15行
            invoice[14] = ('Others', sum(v for d, v in invoice[14:]))
            invoice = invoice[:15]
        invoice.sort(key=lambda v: v[1], reverse=True)

        with open(os.path.join(cfg.TmpPath, '%s.json' % customerName), 'w', encoding='utf-8') as fp:
            json.dump((monthBills, invoice), fp, indent=2, sort_keys=True)


# 下载AWS MR文件
def downloadAwsMR(payerAccount):
    bucket = boto3.resource('s3').Bucket(cfg.Bucket)
    fileName = '%s-aws-billing-csv-20%s-%s.csv' % (payerAccount, cfg.BillMonth[:2], cfg.BillMonth[2:])
    fileFullname = os.path.join(cfg.TmpPath, fileName)
    cfg.log.info('Downloading [%s]%s(MR)' % (cfg.Bucket, fileName))
    bucket.download_file(fileName, fileFullname)
    return fileFullname


def mrFileName(row, customerName=None):
    customerName = customerName if customerName else cfg.CustomerNames.get(row['LinkedAccountId'], 'Missing')
    fileName = '%s-aws-billing-csv-20%s-%s.csv' % (customerName, cfg.BillMonth[:2], cfg.BillMonth[2:])
    return fileName


def getAccountTotal(file_fullname):
    accountTotal = {}
    with open(file_fullname, 'r', encoding='utf-8', newline='') as fp:
        for row in csv.DictReader(fp):
            if row['RecordType'] != 'AccountTotal':
                continue
            for col in Mr.SumCols:
                row[col] = 0
            accountTotal[row['LinkedAccountId']] = row
    return accountTotal


def getStatementTotal(row):
    row['RecordType'] = row['RecordID'] = 'StatementTotal'
    row['LinkedAccountId'] = row['LinkedAccountName'] = ''
    row['InvoiceID'] = row['InvoiceDate'] = ''
    row['ProductCode'] = row['ProductName'] = row['UsageType'] = ''
    row['Operation'] = row['RateId'] = ''
    row['UsageStartDate'] = row['UsageEndDate'] = row['UsageQuantity'] = ''
    period = '%s - %s' % (row['BillingPeriodStartDate'], row['BillingPeriodEndDate'])
    row['ItemDescription'] = 'Total statement amount for period ' + period
    for col in Mr.SumCols:
        row[col] = 0
    return row


def add(bill, row):
    for col in (row['LinkedAccountId'], row['ProductName'], row['UsageType'], row['ItemDescription']):
        cfg.cascadeDictDefault(bill, [(col, {}), ('TotalCost', 0)])
        bill = bill[col]
        bill['TotalCost'] += float(row['TotalCost'])
        if col == row['ItemDescription']:
            cfg.cascadeDictDefault(bill, [('UsageQuantity', 0)])
            bill['UsageQuantity'] += float(row['UsageQuantity'])


def sendBill():
    with Mail() as mail:
        for customerName in cfg.Customers.keys():
            file_names = ['%s-%s.xls' % (customerName, cfg.BillMonth), mrFileName('', customerName)]
            subject = 'AWS账单(20%s-%s)%s' % (cfg.BillMonth[:2], cfg.BillMonth[2:], customerName)
            text = '尊敬的 %s 客户:<br>附件是贵司(20%s-%s)月AWS账单。请查收<br>' \
                   % (customerName, cfg.BillMonth[:2], cfg.BillMonth[2:])
            mail.send(Mr.Receiver, file_names, sub=subject, text=text)


def run():
    mr = Mr()
    for payerAccount in cfg.payers:
        mr.split(payerAccount)

    exchangeRate = cfg.getUsdExchangeRate(cfg.nextMonth(cfg.BillMonth))
    for customerName, cust_obj in cfg.Customers.items():
        fileName = mrFileName('', customerName)
        fileFullname = os.path.join(cfg.TmpPath, fileName)
        if not os.path.exists(fileFullname):
            cfg.log.warning('%s not exists, please confirm.' % fileFullname)
            continue
        if len(cust_obj['Accounts']) > 1:  # 多账号才汇总
            mr.merge(customerName)
        mr.bill(customerName)  # 生成分层账单数据和收据数据
        currency = cust_obj.get('Currency', 'USD').upper()
        XlsBill(exchangeRate, currency).run(customerName)  # 生成xls账单
    sendBill()  # 邮件发送账单


# AWSLamda调用的入口函数
def lambda_handler(event, context):
    run()
    return {
        'statusCode': 200,
        'body': json.dumps('UltraInvoice Finished!')
    }


if __name__ == '__main__':
    run()
