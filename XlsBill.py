#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : XlsBill.py
# @Author: Sui Huafeng
# @Date  : 2019/8/28
# @Desc  : 为每个客户生成xls格式的月账单，为收美元的客户生成Invoice
#
import json
import os
from datetime import date
import xlrd
import xlwt
from xlutils import copy
import config as cfg


class XlsBill(object):
    LastDay = {'01': '31', '03': '31', '05': '31', '07': '31', '08': '31', '10': '31', '12': '31',
               '02': '28', '04': '30', '06': '30', '09': '30', '11': '30'}

    def __init__(self, exchange_rate, currency='USD'):
        self.currency = currency
        self.styles = dict(Invoice={}, Bill={})
        self.template = xlrd.open_workbook(os.path.join(cfg.MetaPath, 'Invoice%s.xls' % currency), formatting_info=True)
        self.headers = {self.template.sheet_names()[i]: self.template.sheet_by_index(i).row_values(0)
                        for i in range(1, 2)}
        self.book = copy.copy(self.template)
        self.__setStyle()
        self.exchangeRate = exchange_rate
        self.xlsRow = 1

    def __setStyle(self):
        sheet, styles = self.book.get_sheet('Invoice'), self.styles['Invoice']
        styles.update({col: getStyle(sheet, 11, col) for col in range(1, 4)})
        styles['No'] = getStyle(sheet, 4, 3)
        styles['Date'] = getStyle(sheet, 5, 3)
        styles['Name'] = getStyle(sheet, 6, 2)
        styles['Addr'] = getStyle(sheet, 7, 2)
        styles['Tel'] = getStyle(sheet, 8, 2)
        styles['billMonth'] = getStyle(sheet, 8, 3)
        styles['Amount'] = getStyle(sheet, 26, 3)
        styles['TotalRmb'] = getStyle(sheet, 27, 1)
        styles['AmountRmb'] = getStyle(sheet, 27, 3)

        sheet, styles = self.book.get_sheet('Bill'), self.styles['Bill']
        for layer in range(6):
            styles[layer] = {header: getStyle(sheet, layer, col) for col, header in enumerate(self.headers['Bill'])}

    def run(self, customerName, invoiceNo):
        with open(os.path.join(cfg.TmpPath, '%s.json' % customerName), 'r', encoding='utf-8') as fp:
            monthBills, invoice = json.load(fp)

        self.book = copy.copy(self.template)
        self.fillInvoice(customerName, invoice, invoiceNo)
        self.fillBill(monthBills)
        file_name = os.path.join(cfg.TmpPath, '%s-%s.xls' % (customerName, cfg.BillMonth))
        self.book.save(file_name)  # 保存文件
        cfg.log.info('%s has been saved.' % file_name)

    def fillBill(self, monthBills):
        sheet = self.book.get_sheet('Bill')
        colWidth = [sheet.col(i).width for i in range(10)]
        for account, l1Bill in monthBills.items():
            sheet = self.__addSheet(sheet, colWidth, account)
            self.xlsRow = 0  # 重设xls当前行
            self.__writeLine(sheet, account, l1Bill, 1)
            for product, l2Bill in l1Bill.items():
                self.__writeLine(sheet, product, l2Bill, 2)
                for region, l3Bill in l2Bill.items():
                    self.__writeLine(sheet, region, l3Bill, 3)
                    for l4, l4Bill in l3Bill.items():
                        layer = 5 if 'UsageQuantity' in l4Bill.keys() else 4
                        self.__writeLine(sheet, l4, l4Bill, layer)
                        if 'UsageQuantity' in l4Bill.keys():
                            continue
                        for description, l5Bill in l4Bill.items():
                            self.__writeLine(sheet, description, l5Bill, 5)
            sheet = None

    def __writeLine(self, sheet, key, bill, layer):
        self.xlsRow += 1
        description = 'Account:%s' % key if layer == 1 else key
        description = '  ' * layer + description
        usageQuantity = bill.get('UsageQuantity', '')
        line = (layer, description, usageQuantity, '', bill['TotalCost'])
        bill.pop('TotalCost')
        for header, data in zip(self.headers['Bill'], line):
            writeCell(sheet, self.styles['Bill'][layer][header], data, row=self.xlsRow)

    def __addSheet(self, sheet, colWidth, account):
        if not sheet:
            sheet = self.book.add_sheet(account, cell_overwrite_ok=True)
            for idx, header in enumerate(self.headers['Bill']):
                sheet.col(idx).width = colWidth[idx]
                writeCell(sheet, self.styles['Bill'][0][header], header)
        sheet.pans_frozen = True
        sheet.horz_split_pos = 1

        return sheet

    def fillInvoice(self, customerName, invoice, invoiceNo):
        cust_obj = cfg.Customers[customerName]
        sheet, styles = self.book.get_sheet('Invoice'), self.styles['Invoice']
        sheet.show_headers = False
        sheet.footer_str = b''
        sheet.header_str = b''
        writeCell(sheet, styles['No'], 'P/I 20%s%04d' % (cfg.nextMonth(cfg.BillMonth), invoiceNo))
        writeCell(sheet, styles['Date'], date.today().strftime('%Y-%m-%d'))
        writeCell(sheet, styles['Name'], cust_obj.get('Name', customerName))
        writeCell(sheet, styles['Addr'], cust_obj.get('Addr', ''))
        writeCell(sheet, styles['Tel'], cust_obj.get('Tel', ''))

        first_join_date = cust_obj.get('FirstJoinedTime', '')
        yy, mm = cfg.BillMonth[:2], cfg.BillMonth[2:]
        start_dd = first_join_date[-2:] if first_join_date[:2] == yy and first_join_date[3:5] == mm else '01'
        end_dd = '29' if mm == '02' and int(yy) / 4 == 0 else self.LastDay[mm]
        writeCell(sheet, styles['billMonth'], '20%s/%s/%s - %s/%s' % (yy, mm, start_dd, mm, end_dd))

        for lines, (product, fee) in enumerate(invoice):
            writeCell(sheet, styles[2], product, row=lines + 11)
            writeCell(sheet, styles[3], fee, row=lines + 11)
        writeCell(sheet, styles['Amount'], xlwt.Formula("SUM(D12:D26)"))

        if self.currency == 'RMB2USD':
            writeCell(sheet, styles['TotalRmb'], '应付人民币总额(汇率：%.4f, 税率：17%%)' % self.exchangeRate)
            writeCell(sheet, styles['AmountRmb'], xlwt.Formula("D27*(1+B29)*%.4f" % self.exchangeRate))


def writeCell(sheet, template_cell, value, row=None, col=None):
    row = template_cell['row'] if row is None else row
    col = template_cell['col'] if col is None else col
    sheet.write(row, col, value)
    cell_style = getStyle(sheet, row, col)['style'] if template_cell['style'] else None
    if cell_style:
        cell_style.xf_idx = template_cell['style'].xf_idx


def getStyle(sheet, row, col):
    row_ = sheet._Worksheet__rows.get(row)
    style = row_._Row__cells.get(col) if row_ else None
    return {'row': row, 'col': col, 'style': style}

