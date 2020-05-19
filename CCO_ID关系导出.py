#!/usr/bin/env python
# -*- conding: utf-8 -*-

import pyodbc
from openpyxl import Workbook
import os
import pyautogui

if __name__ == "__main__":

    # 数据库驱动
    str = 'DRIVER=Microsoft Access Driver (*.mdb, *.accdb);DBQ=C:/HiStudio-工装/3105集中器通信模块出厂检测/data/ndm/NoStaEquip/NoStaCCOCheckDB.mdb'
    # conn = pypyodbc.win_connect_mdb(str)
    # 提示输入开始测试日期时间
    line_split='###'*20
    while True:
        start_date = input("请输入开始测试日期时间(例如：2008-08-08 20:08):\n")
        if len(start_date) != 16:
            print('\n')
            print(line_split)
            print('\n')
            print("你输入的日期时间格式不对，请严格参照示例输入！")
            print('\n')
            print(line_split)
            print('\n')
            continue
        elif start_date[10] != ' ':
            print('\n')
            print(line_split)
            print('\n')
            print("你输入的日期时间格式不对，请严格参照示例输入！")
            print('\n')
            print(line_split)
            print('\n')
            continue
        else:
            break
    # 请输入派工单号
    print('\n')
    wo = input("请输入派工单号及产品形态,(例如:X20190617Y-1-单相):\n")
    wo1 = wo + '.xlsx'
    # 工作簿保存路径
    path_name = os.path.join(r"C:\Users\Lenovo\Desktop\ID清单，请手下留情，勿删！！！", wo1)
    # 新建工作簿
    wb = Workbook(path_name)
    ws = wb.create_sheet(wo, 0)
    ws.append(['芯片ID', '模块ID'])
    # 创建数据库连接对象
    conn = pyodbc.connect(str)
    # 创建游标对象
    cur = conn.cursor()
    # 查询语句
    query_statement = """SELECT 芯片ID值,模块ID值 from (SELECT distinct 芯片ID值,模块ID值,日期 FROM NoStaTableCCOCheck where 芯片ID值<>'' and 日期>=? order by 日期 asc);"""
    # 已测模块数量
    # qty_statement = """select count(distinct 芯片ID值) FROM NoStaTableCCOCheck where 芯片ID值<>'' and 日期>=?"""
    # 执行查询（传递开始测试日期时间参数）
    cur.execute(query_statement, start_date)
    # 查询结果
    result = cur.fetchall()
    result1 = []
    for i in result:
        if i not in result1:
            result1.append(i)
    for row in result1:
        ws.append(list(row))
    wb.save(path_name)
    pyautogui.alert("本批测试 %s 个模块，请注意检查是否有漏测！\n首个ID为：%s" %
                    (len(result1), result1[0][0]))
    conn.commit()
    cur.close()
    conn.close()
