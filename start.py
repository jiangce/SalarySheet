# -*- coding: utf-8 -*-

from lib.excel_parser import ExcelParser
from lib.email_util import send_email
from lib.email_content import EmailContent
import os

PATH = os.path.dirname(__file__)


def get_input(prompt, method=None, default=None):
    result = ''
    if default is None:
        prompt = '%s<EXIT退出>：' % prompt
    else:
        prompt = '%s【默认 - %s】<EXIT退出>：' % (prompt, default)
    while result == '':
        result = input(prompt).strip()
        if result.lower() == 'exit':
            exit(0)
        result = result or default or ''
        if method:
            try:
                result = method(result)
            except:
                result = ''
    return result


def parse_script():
    print('#' * 50)
    print('\t您好，高Boss')
    print('\t欢迎使用工资条发放程序')
    print('#' * 50)
    files = list(enumerate([file_name for file_name in os.listdir(PATH)
                            if os.path.splitext(file_name)[-1].lower() in ['.xlsx', '.xls']], 1))
    if not files:
        print('对不起，在当前目录下找不到任何Excel文件，请确认后重新开始！')
        return
    for f in files:
        print('\t%s - %s' % f)
    file_name = None
    while not file_name:
        file_name = {index: file_name for index, file_name in files}.get(get_input('请选择Excel文件', int))
    parser = ExcelParser(os.path.join(PATH, file_name))
    sheets = list(enumerate(parser.sheet_names, 1))
    for s in sheets:
        print('\t%s - %s' % s)
    sheet_name = None
    while not sheet_name:
        sheet_name = {index: sheet for index, sheet in sheets}.get(get_input('请选择Sheet', int, 1))
    origin_title = get_input('有效表格左上角标题是', str, '姓名')
    people_indices, title_indices, rows = parser.parse(sheet_name, origin_title)
    email_user = None
    email_pass = None
    EmailContent.load_template()
    for name, name_index in people_indices.items():
        print('\n\t%s' % name)
        values = [(parser.config.title[i], rows[name_index][title_indices[i]].value)
                  for i in range(len(parser.config.title))]
        for value in values:
            print('\t\t%s - %s' % value)
        if get_input('是否发邮件给%s' % parser.config.people[name], str, 'n').lower() in ['y', 'yes']:
            if not email_user:
                email_user = get_input('输入发送邮箱用户', str, 'gaoweijie') + '@togeek.cn'
            if not email_pass:
                email_pass = get_input('输入发送邮箱密码', str)
            ec = EmailContent(name, values)
            if send_email([parser.config.people[name]], email_user, email_pass,
                          'smtp.ym.163.com', 25, '工资条', ec.html):
                print('发送成功！')
            else:
                print('发送失败!')


if __name__ == '__main__':
    parse_script()
