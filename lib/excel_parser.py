# -*- coding: utf-8 -*-

import xlrd
from lib.config import Config


class ExcelParser:
    def __init__(self, file_name):
        self.workbook = xlrd.open_workbook(file_name)
        self.config = Config()

    @property
    def sheet_names(self):
        return self.workbook.sheet_names()

    def parse(self, sheet_name, origin_title='姓名'):
        sheet = self.workbook.sheet_by_name(sheet_name)
        rows = list(sheet.get_rows())
        origin = self.find_origin(rows, origin_title)
        title_indices = self.find_title_index(rows, origin)
        people_indices = self.find_people_index(rows, origin)
        return people_indices, title_indices, rows

    def find_title_index(self, rows, origin):
        titles = [row.value for row in rows[origin[0]]]
        return [titles.index(title) for title in self.config.title]

    def find_people_index(self, rows, origin):
        peoples = [row[origin[1]].value for row in rows]
        return {name: peoples.index(name) for name in self.config.people.keys()}

    @staticmethod
    def find_origin(rows, origin_title):
        i = 0
        for row in rows:
            j = 0
            for cell in row:
                if cell.value == origin_title:
                    return i, j
                j += 1
            i += 1


if __name__ == '__main__':
    parser = ExcelParser(r'C:\Users\策\Desktop\工资条\测试.xlsx')
    print(parser.parse('7月'))
