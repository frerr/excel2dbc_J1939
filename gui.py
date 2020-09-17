#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import time
import re
import xlrd
import os

import srcDF2dbc

LOG_LINE_NUM = 0

fin = []
sheet_table = []
select_sheet_table = []

sub_info = '''
Version: 0.1.0
Author: Wu Bing
Contact: bing.wu@inceptio.ai
'''
#初始化
def init1():
    try:
        os.mkdir("dbc")
        os.mkdir("log")
    except:
        print("init dir...")
    print(sub_info)

#日志动态打印
def write_log_to_Text(logmsg):
    current_time = get_current_time()
    logmsg_in = str(current_time) +": " + str(logmsg) + "\n"      #换行
    log_text.insert(tk.END,logmsg_in)

#功能函数
def excel2dbc_button():
    global sheet_table
    write_log_to_Text("Start excel2dbc...")
    write_log_to_Text(sheet_table)

    out = srcDF2dbc.excel2dbc(fin[0], select_sheet_table[0])

    write_log_to_Text("build "+out +"completed")

#功能函数
def excel_import_button():
    global sheet_table
    file_path = filedialog.askopenfilename()
    write_log_to_Text("import: "+str(file_path))
    list1 = re.split(r'[\,/]',file_path)
    if len(fin) != 0:
        fin.pop()
    fin.append(list1[-1])

    data = xlrd.open_workbook(fin[0])
    if len(sheet_table) != 0:
        sheet_table.pop()
    sheet_table = data.sheet_names()

    select_sheet_comboxlist['values'] = sheet_table

#获取当前时间
def get_current_time():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
    return current_time

#用户选择sheet
def combox_sheet_name(event):
    global select_sheet_table 
    if len(select_sheet_table) != 0:
        select_sheet_table.pop()
    select_sheet_table.append(select_sheet_comboxlist.get()) 


#设置窗口
try:
    win = tk.Tk()
    win.title("Excel2dbc_tool")                       #窗口名
    win.geometry('680x180')                         
except:
    print("windows init Error")

#标签
try:
    tk.Label(win, text="log").grid(row=10, column=0)
except:
    print("label Error")
#日志框
try:
    log_text = tk.Text(win, width=100, height=80)
    log_text.grid(row=12, column=0, columnspan=20)
except:
    print("text Error")
#按钮
try:
    build_button = tk.Button(win, text="build", bg="lightblue", width=10, command=excel2dbc_button)
    build_button.grid(row=1, column=8)
    import_button = tk.Button(win, text="import", bg="lightblue", width=10, command=excel_import_button)
    import_button.grid(row=1, column=0)
except:
    print("button Error")    
#下拉框
try:
    number = tk.StringVar()
    select_sheet_comboxlist = ttk.Combobox(win, width=12, textvariable=number)
    select_sheet_comboxlist['values'] = ()     # 设置下拉列表的值

    select_sheet_comboxlist.grid(row=1, column=4)      # 设置其在界面中出现的位置  column代表列   row代表行
    select_sheet_comboxlist.bind("<<ComboboxSelected>>",combox_sheet_name)  #绑定事件(下拉列表框被选中时，绑定combox_sheet_name()函数)
    
except:
    print("comboxlist Error")

init1()
win.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示
