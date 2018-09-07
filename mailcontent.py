# -*- coding: utf-8 -*-
# @Time    : 2018/8/30 下午11:48
# @Author  : alpface
# @Email   : xiaoyuan1314@me.com
# @File    : mailcontent.py
# @Software: PyCharm

import pandas
import os
import codecs
import datetime

excelpath = 'test_excel.xls' # 需要发送邮件的excel

# 将邮件内容转换为html时，contentpath作为log日志记录
log_path = 'content.html'
if os.path.exists(log_path):
    os.system(r'touch {}'.format(log_path))

month_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec']

def mailTitle():
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day
    month = month_list[month-1]
    mail_title = "xiaoyuan's {month} {day} daily report".format(month=month, day=day)
    return mail_title


def custonHtmlContent(title='邮件标题'):
    '''
    自定义html，生成自定义的表格样式
    :param df:
    :param title:
    :return:
    '''
    # 读取excelpath表中的Sheet1表单
    df = pandas.read_excel(excelpath, sheetname='Sheet1')#index_col='Date' 可设置索引列为Date列，但是没有列的名称了郁闷
   # df.set_index(['Date', 'Task', 'Memo'], inplace=True, drop=False) # 设置索引列
    # 给Nan填充为空字符串 要不然邮件中表格如果为空就会显示nan，
    # note 执行fillna后，日期莫名就从2018-10-11 变成 2018-10-11 00:00:00
    df = df.fillna('')


    # 去掉日期中的时间只保留年月日
    # df['Date'] = pandas.to_datetime(df['Date']).dt.normalize()
    # 去掉日期中的时间只保留年月日
    # df['Date'] = pandas.to_datetime(df['Date'], format='%Y-%m-%d')#%Y是4位年（2018），%y是2位年(18)


    columns = list(df.columns)
    index = list(df.index)

    f = open(log_path, 'w')
    f.write('<div style="line-height:1.7;color:#000000;font-size:14px;font-family:Arial">\n\n')
    f.write("<style>\n")
    # 设置边框颜色solid #2F4F4F
    f.write("table,table tr th, table tr td { border:1px solid #2F4F4F; }\n")
    f.write("table { width: auto; min-height: 25px; line-height: 25px; \
   text-align: center; border-collapse: collapse; padding:2px;} \n")
    f.write("</style>\n\n")
    if len(title):
        # 写入邮件标题
        f.write("<table>\n")
        f.write("<caption>" + title + "</caption>\n")
        f.write("<tbody>\n")

    colstr = "<tr height=17 style='height:12.75pt'>\n"
    colstr += "<td></td>\n"
    for i in range(0, len(columns)):
        colstr += '<td>' + columns[i] + '</td>\n'
    colstr += '</tr>'
    f.write(colstr + '\n')
    f.write('\n')

    for i in range(0, len(index)):
        row = str(index[i])
        rowstr = "<tr height=17 style='height:12.75pt'>\n" + "<td>" + row + "</td>\n"
        for j in range(0, len(columns)):
            column = columns[j]
            if column == 'Date':

                # 处理Date列的日期格式，将2018-10-11 00:00:00 修改为 2018-10-11
                content2 = str(df.iat[i, j])
                content2 = fromatDateString(content2)

            else:
                content2 = str(df.iat[i, j])

            rowstr += "<td>" + content2 + "</td>\n"
        rowstr += "</tr>\n"
        f.write(rowstr)

    f.write('</tbody>\n</table>\n</div>')
    f.close()

    content = None
    with open(log_path, 'r+') as f:
        content = f.read()

    return content

def fromatDateString(dateStr):
    newStr = ''
    if len(dateStr):
        d = stringToDate(dateStr)
        newStr = dateToString(d)
    return newStr

#String to Date(datetime)
def stringToDate(string):
    #example '2013-07-22 09:44:15+00:00'
    dt = datetime.datetime.strptime(string, "%Y-%m-%d %H:%M:%S")
    #print dt
    return dt
#Date(datetime) to String
def dateToString(date):
    ds = date.strftime('%Y-%m-%d')
    return ds

def getHtmlContent(excelPath):
    '''
    获取默认的html
    :param excelPath: excel的路径
    :return:
    '''
    xd = pandas.ExcelFile(excelPath)
    df = xd.parse()
    # 给Nan填充为空字符串 要不然邮件中表格如果为空就会显示nan
    df = df.fillna(' ')  #
    with codecs.open(log_path, 'w', 'utf-8') as html_file:
        html_file.write(df.to_html(header=True, index=False))

    content = None
    with open(log_path, 'r+') as html_file:
        content = html_file.read()

    return content

def getMailContent():

    mail_content = custonHtmlContent(title=mailTitle())
    # 使用pd生成的html，文本过长则会显示不全
    # mail_content = getHtmlContent(excelpath)
    return mail_content

