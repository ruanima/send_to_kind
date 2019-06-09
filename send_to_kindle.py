#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
    自动将epub转换成mobi推送到kindle
'''

import poplib
import email
import re
import base64
import time
import os
import subprocess
import smtplib
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr, formatdate
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase


kindlegen = 'bin/kindle'

check_interval = 30     # 邮件收取间隔时间
# 中转邮箱，需要将其添加到kindle的收信人列表(有的邮箱推送会失败，原因不明)
from_addr = 'your_email@tom.com'   
password = os.environ['PASSWORD']    # 中转邮箱密码，从环境变量中获取
pop3_server = 'pop.tom.com'
smtp_server = 'smtp.tom.com'
smtp_server_port = 25
encryption = ''  # TLS, SSL, ''
to_addr = 'your_email@kindle.cn'   # kindle推送收件邮箱


def parse_attach_name(raw_name):
    '''
    args:
        raw_name: eg. =?gb18030?B?t8nN+bDNwOi1xMSpsOC7+i50eHQ=?=
    '''
    print(raw_name)
    t = re.findall(r'(?<=\=\?)[\w\-]*(?=\?[bB])', raw_name)
    if t:
        encoding = t[0]
        return base64.b64decode(raw_name[len(encoding) + 5:]).decode(encoding, errors='replace')
    else:
        return raw_name


def download_attach(mail):
    t = []
    for part in mail.walk():
        if part.get_content_type() != 'application/octet-stream':
            continue

        file_name = parse_attach_name(part.get_filename())
        file_name = 'tmp/{}'.format(file_name)
        payload = base64.b64decode(part.get_payload())
        file = open(file_name, 'wb')
        file.write(payload)
        file.close()
        t.append(file_name)
    return t


def convert_ebook(file_list):
    t = []
    for i in file_list:
        base, ext = os.path.splitext(i)
        if not ext or ext.lower() not in ('.epub', '.mobi', 'azw3'):
            t.append(i)
            continue
        try:
            subprocess.check_call(['bin/kindlegen', '-locale', 'zh', '-c1', i])
        except subprocess.CalledProcessError:
            pass
        if os.path.exists(base + '.mobi'):
            t.append(base + '.mobi')
    return t


def push_to_kindle(file_list):
    if not file_list:
        return

    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['Reply-To'] = from_addr
    msg['To'] = to_addr
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = 'Sent to Kindle'
    msg.attach(MIMEText('send ebooks: \n {}'.format('\n'.join(file_list)), 'plain', 'utf-8'))
    print(file_list)

    for index, i in enumerate(file_list):
        with open(i, 'rb') as f:
            f_name = os.path.split(i)[-1]
            mime = MIMEBase('application', 'octet-stream', name=Header(f_name, 'utf-8').encode())
            mime.add_header('Content-Disposition', 'attachment',
                            filename=Header(f_name, 'utf-8').encode())
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)

    if encryption == 'SSL':
        server = smtplib.SMTP_SSL(smtp_server, smtp_server_port)
    elif encryption == 'TLS':
        server = smtplib.SMTP(smtp_server, smtp_server_port)
        server.starttls()
    else:
        server = smtplib.SMTP(smtp_server, smtp_server_port)
    server.ehlo()
    # server.set_debuglevel(1)
    server.login(from_addr, password)
    server.sendmail(from_addr, [to_addr], msg.as_string())
    server.quit()


def main():
    if not os.path.exists('tmp'):
        os.mkdir('tmp')

    old_index_set = set()
    while True:
        server = poplib.POP3(pop3_server)
        server.user(from_addr)
        server.pass_(password)
        print(server.list())
        mail_indexs = [i for i in server.list()[1] if i not in old_index_set]
        old_index_set.update(mail_indexs)
        print(old_index_set, flush=True)
        messages = [server.retr(int(n.split()[0])) for n in mail_indexs]

        emails = [email.message_from_string(b'\n'.join(message[1]).decode('utf8'))
                for message in messages]
        for index, mail in enumerate(emails):
            raw_books = download_attach(mail)
            if not raw_books: continue
            file_list = convert_ebook(raw_books)
            push_to_kindle(file_list)
            mail_index = int(mail_indexs[index].split()[0])
            server.dele(mail_index)
            old_index_set.remove(mail_indexs[index])
        server.quit()
        time.sleep(check_interval)


if __name__ == '__main__':
    main()
