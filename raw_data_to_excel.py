# -*- coding: utf-8 -*-
"""
Created on Mon May 11 11:38:54 2020

@author: BreezeCat
"""

import xlwings as xw
import json

RAW_DATA_FILE = 'raw_data/95_0511_test.json'
Label_list = ['V', 'W', 'r1', 'Vmax', 'gx', 'gy', 'gth', 'Px2', 'Py2', 'Vx2', 'Vy2', 'r2']


def Fill_data(data_file, SHEET):
    file = open(data_file, 'r')
    data_line = file.readline()
    i = 1
    while(data_line):
        i = i + 1
        data = json.loads(data_line)
        j = 0
        for item in Label_list:
             j = j + 1
             SHEET.cells(i,j).value = data[item]
        data_line = file.readline()
    return
         
    

def Label_first_row(SHEET):
    i = 1
    for item in Label_list:
        SHEET.cells(1,i).value = item
        i = i + 1
    return


if __name__ == '__main__':
    workbook = xw.Book()
    sheet = workbook.sheets['工作表1']
    Label_first_row(sheet)
    Fill_data(RAW_DATA_FILE, sheet)
    