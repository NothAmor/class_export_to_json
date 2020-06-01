#! -*- coding=utf-8 -*-

import xlrd
from datetime import date, datetime
import json

# 表格文件名, 将表格放入跟此脚本同目录下
file = 'class.xls'

def read_excel(line):
    line -= 1

    wb = xlrd.open_workbook(filename=file)

    # 选择 "授课任务" sheet
    st = wb.sheet_by_name("授课任务")

    class_list = []

    for i in range(len(st.row_values(line))):

        # 这里>3是因为我们的课表Excel文件前面有三列啥用没有的东西, 所以不管, 从3开始读取
        if i > 3:
            # 将这一行中所有非空的单元格存入class_list二维数组中, 第二个存入的i是后面要算这是第几周的第几节课
            if st.row_values(line)[i] != "":
                class_list.append([st.row_values(line)[i], i])

    json_data = "["


    for i in range(len(class_list)):
        # 首先先减去前面三个什么用都没有的单元格
        class_time = class_list[i][1] - 3

        weekday = 0

        # 然后一直循环到数值小于5, 一天占五行, 只要计出循环了几次和循环后的数值就能知道这是周几的第几节课
        while class_time >= 5:
            weekday += 1
            class_time -= 5
        
        weekday += 1

        # 存储json数据
        json_data += '{"id": ' + str(i) + ',"name": "' + str(class_list[i][0]) + '","teacherName": "","place": "网络授课","startJc": ' + str(class_time) + ',"endJc": ' + str(class_time) + ',"weekday": ' + str(weekday) + ',"week": "15","flag": 0},'

    # 写入文件
    export_file = open("class_list.json", encoding="utf-8", mode = "w")
    export_file.write(json_data + "]")
    export_file.close()

class_line = input("请输入你的班级在Excel课表中的行数:")
read_excel(int(class_line))