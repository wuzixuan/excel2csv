#!/usr/bin/env python
# -*- coding: utf-8 -*-

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import hashlib
import time
import PandasUtils as pu

LOG_LINE_NUM = 0

class MY_GUI():
    def __init__(self,root):
        self.root = root


    # 设置窗口
    def set_init_window(self):
        self.root.title("excel转换工具")  # 窗口名
        self.root.geometry('681x681+10+10')

        #按钮
        #打开excel
        self.open_excel_button = Button(self.root, text="打开excel", bg="lightblue", width=10,command=self.openfile)  # 调用内部方法  加()为直接调用
        self.open_excel_button.grid(row=1, column=1)
        #转换
        self.convers_to_csv_button = Button(self.root, text="转换", bg="lightblue", width=10,command=self.convers_to_csv)  # 调用内部方法  加()为直接调用
        self.convers_to_csv_button.grid(row=2, column=1)



        # 列表
        self.sheets_list_label = Label(self.root, text="选择需要转换的sheet：")
        self.sheets_list_label.grid(row=2, column=3)
        self.sheets_list = Listbox(self.root, width=66, height=9)
        self.sheets_list.grid(row=3, column=3)
        #文本
        self.log_data_Text_label = Label(self.root, text="提示框：")
        self.log_data_Text_label.grid(row=4, column=3)
        self.log_data_Text = Text(self.root, width=66, height=9)  # 日志框
        self.log_data_Text.grid(row=13, column=3, columnspan=10)



        #功能函数
    def openfile(self):
        file_path = filedialog.askopenfilename()
        if str(file_path).endswith(".xlsx"):
            self.sheets_list.delete(0,END)
            self.write_log_to_Text("打开excel")
            sheets = pu.get_sheet_name_list(file_path)
            for sheet in sheets:
                self.sheets_list.insert(END, sheet)

        else:
            self.write_log_to_Text("这不是excel")


    def convers_to_csv(self):
        print(1)

    #日志动态打印
    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END, logmsg_in)




    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time



def gui_start():
    root = Tk()  # 实例化出一个父窗口
    ZMJ_PORTAL = MY_GUI(root)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()
    root.mainloop()  # 父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示



if __name__ == '__main__':
    gui_start()