# -*- coding: utf-8 -*-

import smtplib
import socket
from email.mime.text import MIMEText


def send_email(to_email, from_email, email_password, email_host, email_port, email_title, email_contents):
    """
    to_email:目标地址,
    from_email：发送地址,
    email_password：邮箱密码,
    email_host：SMPT服务器地址,
    email_port：SMPT费武器端口,
    email_title:发送邮件标题,
    email_contents:发送邮件内容
    """
    try:
        smtpserver = smtplib.SMTP(email_host, email_port)
        code = smtpserver.ehlo()[0]
        usesesmtp = 1
        if not (200 <= code <= 299):
            usesesmtp = 0
            code, resp = smtpserver.helo()
            if not (200 <= code <= 299):
                raise smtplib.SMTPHeloError(code, resp)
        if usesesmtp and smtpserver.has_extn('auth'):
            try:
                smtpserver.login(from_email, email_password)
            except smtplib.SMTPException as e:
                print('登录认证失败！失败代码：%s' % e)
                return False
        else:
            print("EMail服务器不支持认证连接，请使用普通连接")
        msg = MIMEText(email_contents, _subtype='html', _charset='utf8')
        msg['Subject'] = email_title
        msg['From'] = from_email
        msg['To'] = ';'.join(to_email)
        smtpserver.sendmail(from_email, to_email, msg.as_string())
        return True
    except(socket.gaierror, socket.error, socket.herror, smtplib.SMTPException) as e:
        print('Send 发送失败！失败码：%s' % e)
        return False


if __name__ == '__main__':
    from lib.email_content import template

    content = template % {'name': '姜策'}

    send_result = send_email(to_email=['jiangce@togeek.cn'], from_email='jiangce@togeek.cn',
                             email_password='jl123456', email_host='smtp.ym.163.com', email_port=25,
                             email_title='test', email_contents=content)
    if send_result:
        print('send success!')
    else:
        print('send fail!')
