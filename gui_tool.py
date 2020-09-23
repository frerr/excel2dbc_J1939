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
##############################global config######################################
LOG_LINE_NUM = 0
fin = []
sheet_table = []
select_sheet_table = []

sub_info = '''
Version: 0.1.1
Author: Wu Bing
Contact: bing.wu@inceptio.ai
'''

##############################GUI class###########################################
class APP:
    def __init__(self, win):
        self.win = win
        try:
            self.init_gui(self.win)
            self.init_dir()
        except EnvironmentError:
            print("init gui error")
        else:
            self.write_log_to_Text("GUI init ... OK")

######################################初始化所需目录################################
    def init_dir(self):
        global sub_info
        try:
            if not os.path.exists("dbc"):
                os.mkdir("dbc")
            if not os.path.exists("log"):
                os.mkdir("log")                 
        except:
            raise EnvironmentError("init dir Error!")
        print(self,sub_info)

#################################初始化窗口#########################################
    def init_gui(self,win):
        try:
        #设置窗口
            self.win.title("Excel2dbc_tool")#窗口名
            self.win.geometry('580x180')#大小                        
        #标签
            tk.Label(self.win, text="log").grid(row=10, column=0)
        #日志框
            self.log_text = tk.Text(self.win, width=100, height=80)
            self.log_text.grid(row=12, column=0, columnspan=20)
        #按钮
            self.build_button = tk.Button(self.win, text="build", bg="lightblue", width=10, command=self.excel2dbc_button)
            self.build_button.grid(row=1, column=8)
            self.import_button = tk.Button(self.win, text="import", bg="lightblue", width=10, command=self.excel_import_button)
            self.import_button.grid(row=1, column=0)
        #下拉框
            number = tk.StringVar()
            self.select_sheet_comboxlist = ttk.Combobox(self.win, width=12, textvariable=number)
            self.select_sheet_comboxlist['values'] = ()     # 设置下拉列表的值

            self.select_sheet_comboxlist.grid(row=1, column=4)      # 设置其在界面中出现的位置  column代表列   row代表行
            self.select_sheet_comboxlist.bind("<<ComboboxSelected>>",self.combox_sheet_name)  #绑定事件(下拉列表框被选中时，绑定combox_sheet_name()函数)

        except:
            raise EnvironmentError("env Error")

#################################功能函数1#########################################
    def excel2dbc_button(self):
        global sheet_table
        self.write_log_to_Text("Start excel2dbc...")
        self.write_log_to_Text(sheet_table)
        try:
            out = srcDF2dbc.excel2dbc(fin[0], select_sheet_table[0])
        except:
            print("Some Errors....Please check log file...")

        self.write_log_to_Text("build "+out +"completed")

#################################功能函数2#########################################
    def excel_import_button(self):
        global sheet_table
        file_path = filedialog.askopenfilename()
        self.write_log_to_Text("import: "+str(file_path))
        list1 = re.split(r'[\,/]',file_path)
        if len(fin) != 0:
            fin.pop()
        fin.append(list1[-1])

        data = xlrd.open_workbook(fin[0])
        if len(sheet_table) != 0:
            sheet_table.pop()
        sheet_table = data.sheet_names()

        self.select_sheet_comboxlist['values'] = sheet_table

####################3#######返回用户选择的sheet val#################################
    def combox_sheet_name(self,event):
        global select_sheet_table 
        if len(select_sheet_table) != 0:
            select_sheet_table.pop()
        select_sheet_table.append(self.select_sheet_comboxlist.get()) 

###############################获取当前时间########################################
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time

#################################日志动态打印######################################
    def write_log_to_Text(self,logmsg):
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +": " + str(logmsg) + "\n"      #换行
        self.log_text.insert(tk.END,logmsg_in)    

if __name__ == '__main__':
    win = tk.Tk()
    wintk = APP(win)          #生成GUI实例
    win.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示
