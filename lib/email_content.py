# -*- coding: utf-8 -*-

import os


class EmailContent:
    template = ''

    def __init__(self, name, values):
        self.name = name
        self.values = values

    @classmethod
    def load_template(cls):
        template_name = os.path.join(os.path.dirname(__file__), 'template.html')
        try:
            with open(template_name, encoding='utf-8') as f:
                cls.template = f.read()
        except UnicodeDecodeError:
            with open(template_name, encoding='gbk') as f:
                cls.template = f.read()

    @property
    def html(self):
        lines = ['<TABLE border=1><TBODY>']
        tr1 = ['<TR>']
        tr2 = ['<TR>']
        for key, value in self.values:
            tr1.append('<TD>%s</TD>' % key)
            tr2.append('<TD>%s</TD>' % value)
        tr1.append('</TR>')
        tr2.append('</TR>')
        lines.append('\n'.join(tr1))
        lines.append('\n'.join(tr2))
        lines.append('</TBODY></TABLE>')
        return self.template % {'name': self.name, 'content': '\n'.join(lines)}
