# -*- coding: utf-8 -*-

import xlrd
import os
from lib.singleton import Singleton


class Config(metaclass=Singleton):
    def __init__(self):
        self.people = None
        self.title = None
        self.read_config()

    def read_config(self):
        path = os.path.dirname(__file__)
        config_file = os.path.join(path, 'config.xlsx')
        wb = xlrd.open_workbook(config_file)
        self.read_people(wb.sheet_by_name('人员'))
        self.read_title(wb.sheet_by_name('标题'))

    def read_people(self, sheet):
        self.people = {}
        for row in sheet.get_rows():
            name, email = row[0:2]
            if name.value and email.value:
                self.people[name.value] = email.value

    def read_title(self, sheet):
        self.title = [cell.value for cell in list(sheet.get_rows())[0] if cell.value]


if __name__ == '__main__':
    config = Config()
    print(config.people)
    print(config.title)
