# -*- coding:utf-8 -*-
# @Time : 2018/11/23 22:40
# @Author : naihai

import imaplib
import email
import const
import time
import sms


def login():
    try:
        conn = imaplib.IMAP4_SSL(host=const.EMAIL_IMAP_HOST, port=const.EMAIL_IMAP_SSL_PORT)
        conn.login(const.EMAIL_USER, const.EMAIL_PASSWORD)
    except IOError:
        print("登录失败")
        return None
    print("登录成功")
    return conn


def parse_last_mail(conn, lists, time_interval):
    conn.select(mailbox='INBOX')

    # 查找今天的邮件
    today_search = time.strftime("SENTON %d-%b-%Y", time.localtime())
    _, search_result = conn.search(None, today_search)

    # 邮件id集合
    mail_ids = search_result[0].decode('utf-8').split()
    mail_ids.reverse()

    if len(mail_ids) == 0:
        print("今天暂时未收到邮件")
        return

    # 获取最新的一封邮件
    _, data = conn.fetch(mail_ids[0], '(RFC822)')
    msg = email.message_from_bytes(data[0][1])

    # 解析接收时间
    receive_de = email.header.decode_header(msg.get("Received"))[0][0].split("\r\n\t")[-1]
    receive_time = int(time.mktime(time.strptime(receive_de, "%a, %d %b %Y %H:%M:%S +0800 (CST)")))
    if receive_time + time_interval > int(time.time()):
        print("新邮件达到")
    else:
        print("无新邮件")
        return

    # 解析主题
    subject_de = email.header.decode_header(msg.get("Subject"))
    subject = subject_de[0][0].decode(subject_de[0][1])

    # 解析发送人
    sender = None
    sender_de = email.header.decode_header(msg.get("From"))
    if len(sender_de) == 1:
        sender = sender_de[0][0]
    elif len(sender_de) == 2:
        sender = sender_de[2][0].decode("utf-8").split("<")[1].split(">")[0]

    if sender in lists:
        print("有重要邮件达到:" + sender + subject)
        sms.send("【" + sender + "】", "【" + subject + "】", "xxx")


if __name__ == '__main__':
    connection = login()
    important_emails = ["xxx@xx.com", "xx@qq.com"]
    time_inter = 300
    while 1:
        parse_last_mail(connection, important_emails, time_inter)
        time.sleep(time_inter)
