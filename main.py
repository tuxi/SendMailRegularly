# -*- coding: utf-8 -*-
# @Time    : 2018/8/30 下午11:37
# @Author  : alpface
# @Email   : xiaoyuan1314@me.com
# @File    : main.py
# @Software: PyCharm

from smtplib import SMTP_SSL
from email.header import Header
from email.mime.text import MIMEText
from mailcontent import getMailContent, mailTitle
import datetime #定时发送，以及日期
import os
import xlrd # 虽然没有使用，但是也要添加，不然会报错 ImportError: Install xlrd >= 0.9.0 for Excel support
import random

EMAIL_HOST = 'smtp.mxhichina.com'
EMAIL_PORT = 465
EMAIL_HOST_USER = 'yangxy@zorrogps.com' #os.environ.get('DJANGO_EMAIL_USER')

# 使用文件存放密码的好处是: 可防止每次提交代码时不小心把密码也提交了，这里已在.gitignore中忽略了passwordfile
passwordfile = "emailpassword.txt"
if not os.path.exists(passwordfile):
    # 调用系统命令行来创建文件
    os.system(r"touch {}".format(passwordfile))
with open(passwordfile, 'r+') as pf:
    EMAIL_HOST_PASSWORD = pf.readline()

EMAIL_TO = ["392237716@qq.com", "xy@swift.top", "gcloud@dingtalk.com"]
'''
kongp@zorrogps.com zhangy@zorrogps.com
emergy@erlinyou.com
'''
EMAIL_CC = ["sey@live.cn"]
DEFAULT_FROM_EMAIL = EMAIL_HOST_USER
SERVER_EMAIL = EMAIL_HOST_USER

starttime =   [12,  55,  0]
endtime =   [16, 55,  0]


mail_info = {
    "from": EMAIL_HOST_USER,
    "to": ','.join(EMAIL_TO),
     "Cc": ','.join(EMAIL_CC),
    "hostname": EMAIL_HOST,
    "username": EMAIL_HOST_USER,
    "password": EMAIL_HOST_PASSWORD,
    "mail_subject": mailTitle(),
    "mail_encoding": "utf-8"
}

def sendExcelEmail(csvcontent):

    try:
        # 这里使用SMTP_SSL就是默认使用465端口
        smtp = SMTP_SSL(mail_info["hostname"])
        smtp.set_debuglevel(1)

        smtp.ehlo(mail_info["hostname"])
        smtp.login(mail_info["username"], mail_info["password"])

        msg = MIMEText(csvcontent, "html", mail_info["mail_encoding"])
        msg["Subject"] = Header(mail_info["mail_subject"], mail_info["mail_encoding"])
        msg["from"] = mail_info["from"]
        msg["to"] = mail_info["to"]
        msg["Cc"] = mail_info["Cc"]

        smtp.sendmail(mail_info["from"], EMAIL_TO+EMAIL_CC, msg.as_string())

        smtp.quit()
    except Exception as e:
        raise e



def random_datetime(start_date, end_date):
    delta = end_date - start_date
    if delta.days < 0:
        raise Exception('時間範圍錯誤: 結束時間不能小於起始時間')
    seconds = delta.total_seconds()
    inc = random.randrange(seconds)
    return start_date + datetime.timedelta(seconds=inc)

def getRunDateTime():
    '''
    在设定的两个时间范围内，随机获取一个时间，此时间作为发送邮件的时间
    :note:当起始时间小于当前时间并且结束时间还未到达时，起始时间为当前时间+20秒后的时间
    :return:
    '''
    currentDate_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    currentDate = datetime.datetime.strptime(currentDate_str, "%Y-%m-%d %H:%M:%S")
    start_date = datetime.datetime(year=currentDate.year, month=currentDate.month, day=currentDate.day, hour=starttime[0], minute=starttime[1], second=starttime[2])
    end_date = datetime.datetime(year=currentDate.year, month=currentDate.month, day=currentDate.day, hour=endtime[0], minute=endtime[1], second=endtime[2])
    # 如果start_date小于当前时间，并且结束时间还未到达时，则start_date=currentDate
    # 格式化這兩個date，比較其字符串大小
    start_date_str = start_date.strftime('%Y-%m-%d %H:%M:%S')
    end_date_str = end_date.strftime('%Y-%m-%d %H:%M:%S')
    if start_date_str < currentDate_str and end_date_str > currentDate_str:
        start_date = currentDate + datetime.timedelta(seconds=20)
    dt = random_datetime(start_date, end_date)
    seconds = (dt - currentDate).seconds
    if seconds <= 0:
        return currentDate
    return dt


def isCanRunNow(runDate):
    '''
    判断当前时间是否可以发送邮件
    :return:
    '''
    currentDate = datetime.datetime.now()
    # 计算两个时间间隔的秒数
    second = (runDate-currentDate).seconds
    # 正常情况下时间到了runDate就要发送邮件，可有时候会有误差，这里允许误差5秒
    if second <= 5:
        return True
    else:
        return False


if __name__=='__main__':

    runDateTime = getRunDateTime()
    print('发送时间{}'.format(str(runDateTime)))
    if len(starttime) != len(endtime):
        raise Exception('# Error: the run time format is not correct!')
    else:
        while True:
            if isCanRunNow(runDateTime):
                try:
                    sendExcelEmail(getMailContent())
                except Exception as e:
                    print(e.__str__())
                break

        print('任务结束')
