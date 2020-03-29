#!/usr/bin/python
# -*- coding:utf-8 -*-

import datetime, os
import logging
from multiprocessing import freeze_support
import WebInfo


# 1. 读取配置文件
def read_conf(conf_path):
    import json
    with open(conf_path, encoding='utf-8-sig') as f:
        data = json.load(f)
    # for key, value in data.items():
    #     print("配置文件：", key, ":", value)
    return data


# 2. 读取待处理文件-excel
# 2.1 选择文件, 并复制到目录文件
def select_file(conf):
    import shutil, re
    files = os.listdir(conf["log"])
    files_set = []
    data = {}
    print("文件列表：")
    for file in files:
        if os.path.splitext(file)[1] == ".txt":
            files_set.append(file)
            print('[', len(files_set), ']:', file)
            if "main" in file:
                continue
            with open(conf["log"] + "/" + file, encoding='utf-8') as f:
                line = f.readline()
                while line:
                    # print(line)
                    line.rstrip("\n")
                    if "DNS - INFO - [" in line:
                        newl = re.findall(r'[[](.*)[]]', line)
                        strs = newl[0].split(',')
                        data[int(strs[0])] = strs
                    elif "SELENIUM - INFO - [" in line:
                        newl = re.findall(r'[[](.*)[]]', line)
                        strs = newl[0].split(',')
                        data[int(strs[0])] = strs
                    line = f.readline()
    for key, value in data.items():
        print(key, len(value), value)
    # return dst_file


# 2.2 读取文件内容
def get_excel_data(fp, fn, width, isprocess):
    from openpyxl import Workbook
    from openpyxl import load_workbook

    wb = load_workbook(fp + "/" + fn)
    ws = wb.worksheets[0]
    data_arry = {}
    for r in range(ws.max_row):
        if isprocess == 1:
            if (ws.cell(row=r+1, column=width+1).value is None) and (ws.cell(row=r+1, column=width+2).value is None):
                data_arry[r + 1] = []
                for i in range(1, width+1):
                    data_arry[r + 1].append(ws.cell(row=r+1, column=i).value)
        else:
            data_arry[r + 1] = []
            for i in range(1, width + 1):
                data_arry[r + 1].append(ws.cell(row=r + 1, column=i).value)
    return data_arry


def main():
    # begin
    conf_fp = "../workdata/conf/config.json"
    conf_data = read_conf(conf_fp)
    select_file(conf_data)

    # os.system("pause")


if __name__ == "__main__":
    freeze_support()
    main()
