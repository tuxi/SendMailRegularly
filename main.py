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

EMAIL_TO = ["wangshuai@swift.top", "coderhong@126.com", "392237716@qq.com"]
'''
kongp@zorrogps.com zhangy@zorrogps.com
emergy@erlinyou.com
'''
EMAIL_CC = ["sey@live.cn"]
DEFAULT_FROM_EMAIL = EMAIL_HOST_USER
SERVER_EMAIL = EMAIL_HOST_USER

kHr = 0      # index for hour
kMin = 1   # index for minute
kSec = 2    # index for second
kPeriod1 = 0  #时间段1，这里定义了两个代码执行的时间段
starttime =   [[10,  00,  0]]     # 一个时间段的起始时间，hour, minute 和 second
endtime =   [[20, 10,  0]]    # 一个时间段的终止时间
sleeptime = 5    # 扫描间隔时间，s


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
        print(e.__str__())


def T1LaterThanT2(time1,time2):   # 根据给定时分秒的两个时间，比较其先后关系
    # t1 < t2, false, t1 >= t2, true
    if len(time1) != 3 or len(time2) != 3:
         raise Exception('# Error: time format error!')
    T1 = time1[kHr]*3600 + time1[kMin]*60 + time1[kSec] # s
    T2 = time2[kHr]*3600 + time2[kMin]*60 + time2[kSec] # s
    if T1 < T2:
        return False
    else:
        return True

def random_datetime(start_datetime, end_datetime):
    delta = end_datetime - start_datetime
    inc = random.randrange(delta.total_seconds())
    return start_datetime + datetime.timedelta(seconds=inc)

def getRunDateTime():
    mytime = datetime.datetime.now()
    start_datetime = datetime.datetime(year=mytime.year, month=mytime.month, day=mytime.day, hour=19, minute=35, second=0)
    # datetime.datetime(2016, 8, 17, 10, 0, 0)
    end_datetime = datetime.datetime(year=mytime.year, month=mytime.month, day=mytime.day, hour=19, minute=38, second=2)
    #datetime.datetime(2016, 8, 17, 18, 0, 0)
    dt = random_datetime(start_datetime, end_datetime)
    print(dt)
    return dt

# def isCanRunNow():
#     '''
#     判断当前时间是否可以发送邮件
#     :return:
#     '''
#     mytime = datetime.datetime.now()
#     currtime = [mytime.hour, mytime.minute, mytime.second]
#     if (T1LaterThanT2(currtime, starttime[kPeriod1]) and (not T1LaterThanT2(currtime, endtime[kPeriod1]))):
#         return True
#     else:
#         return False

def isCanRunNow():
    '''
    判断当前时间是否可以发送邮件
    :return:
    '''
    mytime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    runDateTime = getRunDateTime()
    if str(mytime) == str(runDateTime):
        return True
    else:
        return False


if __name__=='__main__':

    if len(starttime) != len(endtime):
        raise Exception('# Error: the run time format is not correct!')
    else:
        while True:
            if isCanRunNow():
                sendExcelEmail(getMailContent())
                break

        print('任务结束')
