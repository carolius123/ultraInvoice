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

if __name__ == '__main__':
    # os.environ['receiver'] = 'sui.hf@163.com'
    # os.environ['month'] = '1910'
    os.environ['log'] = 'DEBUG'

import config as cfg
from XlsBill import XlsBill
from Mail import Mail

Receiver = os.environ.get('receiver', 'maling1@ultrapower.com.cn,jiayf@ultrapower.com.cn').split(',')
SumCols = {'UsageQuantity', 'CostBeforeTax', 'Credits', 'TaxAmount', 'TotalCost'}


class CsvFile(object):
    CsvWriters = {}

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
            self.__writeSupport(rowTotal, linkedAccountId)
            self.__writeDiscount(rowTotal, linkedAccountId)
            fileName = mrFileName(rowTotal)
            self.CsvWriters[fileName][0].writerow(rowTotal)
        self.__close()

    # 拆分行写入客户csv文件
    def __writeItems(self, file_fullname, accountTotal, payerAccount):
        removeCredit = cfg.payers.get(payerAccount, {}).get('removeCredit', False)
        with open(file_fullname, 'r', encoding='utf-8', newline='') as fp:
            fp.readline()  # 抛弃首行提示信息 Don't see your tags in the report....
            for row in csv.DictReader(fp):
                if row['RecordType'] != 'LinkedLineItem':  # 只拆分资源消费记录
                    continue
                if removeCredit:  # 去掉Credit抵扣，按泰岳折扣算
                    row['Credits'] = '0'
                    row['TotalCost'] = row['CostBeforeTax']
                for col in SumCols:
                    accountTotal[row['LinkedAccountId']][col] += float(row[col])
                fileName = mrFileName(row)
                self.__open(fileName, row) if fileName not in self.CsvWriters else None
                self.CsvWriters[fileName][0].writerow(row)

    # 添加泰岳折扣行，并更新rowTotal(优先使用Customer上的Discount属性，其次使用Account上的)
    def __writeDiscount(self, rowTotal, linkedAccountId):
        cust_obj = cfg.Customers[cfg.CustomerNames[linkedAccountId]]
        linkedAccount = cust_obj['Accounts'][linkedAccountId]
        discountRate = cust_obj.get('Discount', linkedAccount.get('Discount', 0))
        if not discountRate:
            return

        row = rowTotal.copy()
        row['RecordType'] = 'LinkedLineItem'
        row['RecordID'] = ''
        row['ItemDescription'] = '%g%% Discount offered by Ultrapower' % discountRate
        row['ProductName'] = 'Ultrapower Discount'
        row['ProductCode'] = row['UsageType'] = 'UltraDiscount'
        row['UsageType'] = 'NORG-Discount'
        for col in SumCols:
            row[col] = '0'

        discount = - rowTotal['TotalCost'] * discountRate / 100
        rowTotal['TotalCost'] += discount  # 更新汇总行
        row['UsageQuantity'] = str(rowTotal['TotalCost'])
        row['TotalCost'] = str(discount)

        fileName = mrFileName(row)
        self.CsvWriters[fileName][0].writerow(row)
        return

    # 添加aws Support, 并更新rowTotal
    def __writeSupport(self, rowTotal, linkedAccountId):
        cust_obj = cfg.Customers[cfg.CustomerNames[linkedAccountId]]
        linkedAccount = cust_obj['Accounts'][linkedAccountId]
        supportRate = cust_obj.get('SupportRate', linkedAccount.get('SupportRate', 0))
        supportDiscount = cust_obj.get('SupportDiscount', linkedAccount.get('SupportDiscount', 0))
        if not (supportRate or supportDiscount):
            return

        row = rowTotal.copy()
        row['RecordType'] = 'LinkedLineItem'
        row['RecordID'] = ''
        row['ItemDescription'] = 'AWS Support(%g%%)' % supportRate
        row['ProductName'] = 'AWS Premium Support(%g%%)' % supportRate
        row['ProductCode'] = row['UsageType'] = 'OCBPremiumSupport'
        row['UsageType'] = 'NORG-Support'
        for col in SumCols:
            row[col] = '0'

        support = rowTotal['TotalCost'] * supportRate / 100
        rowTotal['TotalCost'] += support  # 更新汇总行
        row['UsageQuantity'] = str(rowTotal['TotalCost'])
        row['TotalCost'] = str(support)

        fileName = mrFileName(row)
        self.CsvWriters[fileName][0].writerow(row)
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


class JsonFile(object):
    RegionCode2Name = {'ap-northeast-1': 'Asia Pacific (Tokyo)', 'ap-northeast-2': 'Asia Pacific (Seoul)',
                       'ap-east-1': 'Asia Pacific (Hong Kong)', 'ap-southeast-1': 'Asia Pacific (Singapore)',
                       'ap-southeast-2': 'Asia Pacific (Sydney)', 'ap-south-1': 'Asia Pacific (Mumbai)',
                       'ca-central-1': 'Canada (Central)', 'eu-west-1': 'EU (Ireland)',
                       'eu-central-1': 'EU (Germany)', 'eu-north-1': 'EU (Stockholm)',
                       'eu-west-2': 'EU (London)', 'eu-west-3': 'EU (Paris)',
                       'me-south-1': 'Middle East (Bahrain)', 'sa-east-1': 'South America (Sao Paulo)',
                       'us-gov-east-1': 'AWS GovCloud (US-East)', 'us-gov-west-1': 'AWS GovCloud (US-West)',
                       'us-east-1': 'US East (Northern Virginia)', 'us-east-2': 'US East (Ohio)',
                       'us-west-1': 'US West (Northern California)', 'us-west-2': 'US West (Oregon)',
                       'ap-northeast-3': 'Asia Pacific (Osaka-Local)'
                       }
    RegionAbbr2Name = {'APN1': 'Asia Pacific (Tokyo)', 'APN2': 'Asia Pacific (Seoul)',
                       'APE1': 'Asia Pacific (Hong Kong)', 'APS1': 'Asia Pacific (Singapore)',
                       'APS2': 'Asia Pacific (Sydney)', 'APS3': 'Asia Pacific (Mumbai)',
                       'CAN1': 'Canada (Central)', 'EU': 'EU (Ireland)',
                       'EUC1': 'EU (Germany)', 'EUN1': 'EU (Stockholm)',
                       'EUW2': 'EU (London)', 'EUW3': 'EU (Paris)',
                       'MES1': 'Middle East (Bahrain)', 'SAE1': 'South America (Sao Paulo)',
                       'UGW1': 'AWS GovCloud (US-West)', 'USE1': 'US East (Northern Virginia)',
                       'USE2': 'US East (Ohio)', 'USW1': 'US West (Northern California)',
                       'USW2': 'US West (Oregon)', 'APN3': 'Asia Pacific (Osaka-Local)',
                       'UGE1': 'AWS GovCloud (US-East)'
                       }
    CloudFrontRegionAbbr2Name = {'AP': 'Asia Pacific', 'AU': 'Australia', 'CA': 'Canada', 'EU': 'Europe',
                                 'IN': 'India', 'JP': 'Japan', 'ME': 'Middle East', 'SA': 'South Africa',
                                 'ZA': 'South America', 'US': 'United States'
                                 }

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
                key = (row['ProductCode'], row['UsageType'], row['ItemDescription'])
                for col in SumCols:
                    row[col] = float(row[col])

                if key not in consolidateRows:
                    consolidateRows[key] = row
                else:
                    for col in SumCols:
                        consolidateRows[key][col] += row[col]

        if len(accounts) < 2:
            return  # 只有一个有费用账号
        with open(fileFullname, 'a', encoding='utf-8', newline='') as fp:
            csvWriter = csv.DictWriter(fp, fieldnames=row.keys())
            statementTotalRow = getStatementTotal(row.copy())
            for consolidateRow in consolidateRows.values():
                consolidateRow['RecordType'] = 'PayerLineItem'
                consolidateRow['LinkedAccountId'] = consolidateRow['LinkedAccountName'] = ' All'
                for col in SumCols:
                    statementTotalRow[col] += consolidateRow[col]
                    consolidateRow[col] = str(consolidateRow[col])
                csvWriter.writerow(consolidateRow)

            for col in SumCols:
                statementTotalRow[col] = str(statementTotalRow[col])
            csvWriter.writerow(statementTotalRow)

    # 汇总成靠近AWS系统内形式的分层的的月账单
    def invoice(self, customerName):
        monthBills = {}
        fileName = mrFileName('', customerName)
        fileFullname = os.path.join(cfg.TmpPath, fileName)
        with open(fileFullname, 'r', encoding='utf-8', newline='') as fp:
            [self.add(monthBills, row) for row in csv.DictReader(fp) if row['RecordType'][-8:] == 'LineItem']

        monthBill = monthBills[' All'] if ' All' in monthBills else list(monthBills.values())[0]
        invoice = [(product, round(v['TotalCost'], 2))
                   for product, v in monthBill.items() if product != 'TotalCost' and abs(v['TotalCost']) > 0.005]
        invoice.sort(key=lambda v: abs(v[1]), reverse=True)
        if len(invoice) > 15:  # 最多保留15行
            others = sum(v for d, v in invoice[14:])
            invoice[14] = ('Others', 0)
            invoice = invoice[:15]
        else:
            others = None
        invoice.sort(key=lambda v: v[1], reverse=True)
        if others:  # 保证其他在所有正值行之后
            invoice[invoice.index(('Others', 0))] = ('Others', others)

        with open(os.path.join(cfg.TmpPath, '%s.json' % customerName), 'w', encoding='utf-8') as fp:
            json.dump((monthBills, invoice), fp, indent=2, sort_keys=True)

    def add(self, bill, row):
        productName = row['ProductName'].split('AWS ', 1)[-1].split('Amazon ')[-1]
        regionName = self.__getRegion(row)
        l4Name = self.__l4Label(row)
        for layer in (row['LinkedAccountId'], productName, regionName, l4Name, row['ItemDescription']):
            if not layer:  # 不需要l4Name这一级
                continue
            cfg.cascadeDictDefault(bill, [(layer, {}), ('TotalCost', 0)])
            bill = bill[layer]
            bill['TotalCost'] += float(row['TotalCost'])
        cfg.cascadeDictDefault(bill, [('UsageQuantity', 0)])
        bill['UsageQuantity'] += float(row['UsageQuantity'])

    # 从账单行中尽量抽取region信息，向AWS系统中账单层次靠拢
    def __getRegion(self, row):
        if float(row['TotalCost']) < 0:
            return ' No Region'
        availabilityZone = row.get('AvailabilityZone', '')
        if availabilityZone:
            regionCode = availabilityZone if availabilityZone[-1] <= '9' else availabilityZone[:-1]
            regionName = self.RegionCode2Name.get(regionCode, '')
            if regionName:
                return regionName

        productCode = row['ProductCode']
        splitedUsageType = row['UsageType'].split('-')
        if productCode == 'AWSGlobalAccelerator':  # AP/NA/Global
            regionName = splitedUsageType[0]
        elif productCode == 'awskms':  # usageType中带着regionCode
            regionCode = 'Global' if len(splitedUsageType) < 3 else '-'.join(splitedUsageType[:3])
            regionName = self.RegionCode2Name.get(regionCode, 'US East (Northern Virginia)')
        elif productCode == 'AmazonCloudFront':
            regionName = self.CloudFrontRegionAbbr2Name.get(splitedUsageType[0], 'Global')
        else:
            regionName = self.RegionAbbr2Name.get(splitedUsageType[0], 'US East (Northern Virginia)')
        return regionName

    RdsL4 = {' storage ': ' Storage and I/O', ' I/O requests ': ' Storage and I/O',
             'Aurora': ' for Aurora', 'MySQL': ' for MySQL Community Edition',
             'backup storage': ' Backup Storage', 'rovisioned ': ' Provisioned Storage'
             }

    def __l4Label(self, row):
        product_name = row['ProductName']
        operation = row['Operation']
        usage_type = row['UsageType']
        line_item_description = row['ItemDescription']

        if product_name == 'Amazon Relational Database Service':
            for k, v in self.RdsL4.items():
                if k in line_item_description:
                    type_ = product_name + v
                    break
            else:
                type_ = product_name

            if 'hour (' not in line_item_description and 'for ' in type_:
                type_ += ' Reserved Instances'

        elif product_name == 'Amazon Elastic Compute Cloud':
            type_ = operation if 'LoadBalancing' in operation or 'NatGateway' in operation \
                else 'EBS' if 'EBS' in usage_type \
                else 'Elastic IP Addresses' if 'ElasticIP' in usage_type \
                else 'Amazon Elastic Compute Cloud '
            if type_ == 'Amazon Elastic Compute Cloud ':
                type_ += 'Running Linux/UNIX' if 'Linux' in line_item_description \
                    else 'Running Windows' if 'Windows' in line_item_description \
                    else ''
                type_ += ' Spot Instances' if 'SpotUsage' in usage_type \
                    else ' Reserved Instances' if 'On Demand' not in line_item_description \
                    else ''
        elif product_name in ('Amazon ElastiCache', 'Amazon Redshift'):
            type_ = product_name + ' ' + operation
            if ' instance' in line_item_description:
                type_ += ' Reserved Instances'
        elif product_name == 'Amazon Elastic File System' and operation == 'Storage':
            type_ = product_name + ' Standard Storage Class'
        elif product_name == 'AWS Shield':
            type_ = product_name + ' ' + usage_type if usage_type == 'Shield-Monthly-Fee' else 'Bandwidth'
        else:
            type_ = None

        if float(row['TotalCost']) < 0:
            type_ = None

        return type_


# 下载AWS MR文件
def downloadAwsMR(payerAccount):
    bucket = boto3.resource('s3').Bucket(cfg.Bucket)
    fileName = '%s-aws-cost-allocation-20%s-%s.csv' % (payerAccount, cfg.BillMonth[:2], cfg.BillMonth[2:])
    fileFullname = os.path.join(cfg.TmpPath, fileName)
    cfg.log.info('Downloading [%s]%s(MR)' % (cfg.Bucket, fileName))
    bucket.download_file(fileName, fileFullname)
    return fileFullname


def mrFileName(row, customerName=None):
    customerName = customerName if customerName else cfg.CustomerNames.get(row['LinkedAccountId'], 'Missing')
    fileName = '%s-%s.csv' % (customerName, cfg.BillMonth)
    return fileName


# 重新计算各账户的汇总数据
def getAccountTotal(file_fullname):
    accountTotal = {}
    with open(file_fullname, 'r', encoding='utf-8', newline='') as fp:
        fp.readline()  # 抛弃首行提示信息 Don't see your tags in the report....
        for row in csv.DictReader(fp):
            if row['RecordType'] != 'AccountTotal':
                continue
            for col in SumCols:
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
    for col in SumCols:
        row[col] = 0
    return row


def sendBill():
    with Mail() as mail:
        for customerName, cust_obj in cfg.Customers.items():
            file_names = ['%s-%s.xls' % (customerName, cfg.BillMonth)]
            send_csv = cust_obj.get('Csv', 0)
            if send_csv:
                file_names.append(mrFileName('', customerName))
            subject = 'AWS账单(20%s-%s)%s' % (cfg.BillMonth[:2], cfg.BillMonth[2:], customerName)
            text = '尊敬的 %s 客户:<br>附件是贵司(20%s-%s)月AWS账单。请查收<br>' \
                   % (customerName, cfg.BillMonth[:2], cfg.BillMonth[2:])
            mail.send(Receiver, file_names, sub=subject, text=text)


def run():
    csvFile = CsvFile()
    for payerAccount in cfg.payers:
        csvFile.split(payerAccount)

    exchangeRate = cfg.getUsdExchangeRate(cfg.nextMonth(cfg.BillMonth))
    invoieNo = 0
    jsonFile = JsonFile()
    for customerName, cust_obj in cfg.Customers.items():
        fileName = mrFileName('', customerName)
        fileFullname = os.path.join(cfg.TmpPath, fileName)
        if not os.path.exists(fileFullname):
            cfg.log.warning('%s not exists, please confirm.' % fileFullname)
            continue
        if len(cust_obj['Accounts']) > 1:  # 多账号才汇总
            jsonFile.merge(customerName)
        jsonFile.invoice(customerName)  # 生成分层账单数据和收据数据
        currency = cust_obj.get('Currency', 'USD').upper()
        invoieNo += 1
        XlsBill(exchangeRate, currency).run(customerName, invoieNo)  # 生成xls账单
        x = 1
    # sendBill()  # 邮件发送账单


# AWSLamda调用的入口函数
def lambda_handler(event, context):
    run()
    return {
        'statusCode': 200,
        'body': json.dumps('UltraInvoice Finished!')
    }


# EC2上运行的入口
if __name__ == '__main__':
    run()
