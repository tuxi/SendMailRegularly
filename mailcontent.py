# -*- coding: utf-8 -*-
# @Time    : 2018/8/30 下午11:48
# @Author  : alpface
# @Email   : xiaoyuan1314@me.com
# @File    : mailcontent.py
# @Software: PyCharm

import pandas
import os
import codecs

excelpath = 'test_excel.xls' # 需要发送邮件的excel
mail_title = "xiaoyuan's Aug 31 daily report"

# 将邮件内容转换为html时，contentpath作为log日志记录
log_path = 'content.html'
if os.path.exists(log_path):
    os.system(r'touch {}'.format(log_path))

def custonHtmlContent(title='邮件标题'):
    '''
    自定义html，生成自定义的表格样式
    :param df:
    :param title:
    :return:
    '''
    # 读取excelpath表中的Sheet1表单
    df = pandas.read_excel(excelpath, sheetname='Sheet1')

    # 给Nan填充为空字符串 要不然邮件中表格如果为空就会显示nan
    df = df.fillna(' ')

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

    mail_content = custonHtmlContent(title=mail_title)
    # 使用pd生成的html，文本过长则会显示不全
    # mail_content = getHtmlContent(excelpath)
    return mail_content

