# -*- coding: utf-8 -*-
# @Time    : 2018/8/30 下午11:48
# @Author  : alpface
# @Email   : xiaoyuan1314@me.com
# @File    : mailcontent.py
# @Software: PyCharm

import pandas
import os

excelpath = 'test1.csv' # note: 需要將excel表格导出为csv格式，不然pandas解析会报错


contentpath = 'contentpath.txt'
if os.path.exists(contentpath):
    os.system(r'touch {}'.format(contentpath))

def genHtmlContent(df, title):
    '''

    :param df:
    :param title:
    :return:
    '''
    columns = list(df.columns)
    index = list(df.index)

    f = open(contentpath, 'w')
    f.write('<div style="line-height:1.7;color:#000000;font-size:14px;font-family:Arial">\n\n')
    f.write("<style>\n")
    f.write("table,table tr th, table tr td { border:1px solid #0094ff; }\n")
    f.write("table { width: auto; min-height: 25px; line-height: 25px; \
   text-align: center; border-collapse: collapse; padding:2px;} \n")
    f.write("</style>\n\n")
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
        rowstr = "<tr height=17 style='height:12.75pt'>\n" + "<td>" + str(index[i]) + "</td>\n"
        for j in range(0, len(columns)):
            rowstr += "<td>" + str(df.iat[i, j]) + "</td>\n"
        rowstr += "</tr>\n"
        f.write(rowstr)

    f.write('</tbody>\n</table>\n</div>')
    f.close()

    content = None
    with open(contentpath, 'r+') as f:
        content = f.read()

    return content

def getMailContent():
    df1 = pandas.read_csv(excelpath, encoding='utf-8', index_col=0)

    title1 = '测试标题'

    mail_content = genHtmlContent(df1, title1)
    return mail_content
