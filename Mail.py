#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : Mail.py
# @Author: Sui Huafeng
# @Date  : 2019/10/17
# @Desc  : 
#

import os
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import config as cfg


class Mail(object):
    Sender = 'yyw1@ultrapower.com.cn'  # 发件人邮箱账号
    Passwd = 'ultra#0917'
    MailServer = 'mail.ultrapower.com.cn'
    Port = '587'
    Bcc = ['suihf@ultrapower.com.cn']

    def __init__(self, subject='', text=''):
        self.sub = subject
        self.text = text
        self.server = smtplib.SMTP(self.MailServer, self.Port)  # 发件人邮箱中的SMTP服务器，端口是25
        self.server.login(self.Sender.split('@')[0], self.Passwd)  # 发件人邮箱账号、邮箱密码

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.server.quit()

    def send(self, receive, attachments, sub='', text='', ):
        try:
            msg = MIMEMultipart('related')
            msg['From'] = formataddr(["泰岳云业务", self.Sender])  # 发件人邮箱昵称、发件人邮箱账号
            msg['To'] = receive[0]  # 收件人邮箱昵称、收件人邮箱账号
            if len(receive) > 1:
                msg['Cc'] = ','.join(receive[1:])  # 收件人邮箱昵称、收件人邮箱账号
            msg['Subject'] = sub
            text += """<br><br><br><small>
            (如您收到的附件后缀为.dat, 是您的邮件客户端配置引起的，请保存并修改为.xls)</small><br><br>
            云业务事业部<br>
            神州泰岳软件股份有限公司<br>
            电话：+86(10) 58847555<br>
            传真：+86(10) 58847576<br>
            网址：www.ultrapower.com.cn<br>
            地址：北京市朝阳区北苑路甲13号院北辰泰岳大厦22层  100107<br>
            （地铁5号线北苑路北站，A2出口，向北100米十字路口西南角<br>
      ----------- 居利思义 • 身劳心安 • 人强我强 • 共同发展 -----------<br>
<small>保密提示：本邮件仅供上述收件人使用，并可能包含受法律保护之秘密信息。如果您并非上述所列之收件人，请立即删除本邮件，并勿阅读、传播、复制或以任何方式扩散本邮件。谢谢！</small>
            """
            msg.attach(MIMEText(text, 'html', 'utf-8'))

            for attachment in attachments:
                att = MIMEBase('application', 'octet-stream')
                att.set_payload(open(os.path.join(cfg.TmpPath, attachment), 'rb').read())
                att.add_header('Content-Disposition', 'attachment', filename=Header(attachment, 'utf-8').encode())
                encoders.encode_base64(att)
                msg.attach(att)

            self.server.sendmail(self.Sender, receive + self.Bcc, msg.as_string())
        except Exception as err:
            cfg.log.warning('Mail failed to sent:' + str(err))
            return
        cfg.log.info('[%s %s] sent to %s' % (sub, attachments[0], ','.join(x for x in receive)))
