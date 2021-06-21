# -*- coding: utf-8 -*-
# @Date    : 2021-06-12 17:41:21
# @Author  : Nora Yao(norayao0817@gmail.com)
# @Link    :
# @Version : Python3.7

# import sys
# import json
# import requests
# import random
# import time
# # google_trans_new 依赖于 json, requests, random, re
# from google_trans_new import google_translator

import os
import re
import tkinter as tk
from tkinter import filedialog, Label
from openpyxl import Workbook


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.getFile_btn = tk.Button(self)
        self.filePath_entry = tk.Entry(self, width=30)
        self.master = master
        self.pack()
        # self.translator = google_translator()
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.create_widgets()

    def create_widgets(self):
        # 显示文件路径
        self.filePath_entry.grid(row=1, column=1, padx=0, pady=10)
        self.filePath_entry.delete(0, "end")
        self.filePath_entry.insert(0, "请选择文件")

        # 获取文件
        self.getFile_btn['width'] = 15
        self.getFile_btn['height'] = 1
        self.getFile_btn["text"] = "打开"
        self.getFile_btn.grid(row=1, column=2, padx=5, pady=10)
        self.getFile_btn["command"] = self.select_dir

        # 显示结果

    # 打开文件并显示路径
    def select_dir(self):
        default_dir = r"文件路径"
        self.pathValue = tk.filedialog.askdirectory(title=u'选择文件', initialdir=(os.path.expanduser(default_dir)))
        self.filePath_entry.delete(0, "end")
        self.filePath_entry.insert(0, self.pathValue)

        fileList = self.achieve_filelist(self.pathValue)
        pathPrefix = (self.pathValue).replace(os.path.basename(self.pathValue), '')

        catalog = self.pathValue + '.xlsx'
        self.worksheet['A1'] = '类目'
        self.worksheet['B1'] = '文件名'
        self.worksheet['C1'] = '参考翻译'
        self.worksheet['D1'] = '文件类型'
        self.worksheet['E1'] = '路径'
        self.workbook.save(catalog)

        n = 2
        for root, paths, files in fileList:
            for file in files:
                folderName = root.replace(pathPrefix, '')
                fileSplit = file.split('.')
                fileName = file.split('.')[0]
                fileSuffix = ''
                if len(fileSplit) > 1:
                    fileSuffix = fileSplit[1]
                filePath = os.path.join(root, file)

                self.worksheet['A' + str(n)] = folderName
                self.worksheet['B' + str(n)] = fileName
                self.worksheet['C' + str(n)] = '' # self.refactor(self.filename_translation(self.trim(fileName)))
                self.worksheet['D' + str(n)] = fileSuffix
                self.worksheet['E' + str(n)].value = filePath
                self.worksheet['E' + str(n)].hyperlink = "file://" + filePath
                n += 1
                self.workbook.save(catalog)
                folderLabelText = folderName + "     目录已添加"
                self.folderLabel = tk.Label(self, text=folderLabelText, wraplength=350, justify='center')
                self.folderLabel.grid(row=2, column=1, padx=5, pady=10)
            # 停止运行5秒防止谷歌拒绝服务
            # time.sleep(5)


        label_text = self.pathValue + "     目录已生成"
        self.result_lable = tk.Label(self, text=label_text, wraplength=350, justify='center')
        self.result_lable.grid(row=3, column=1, padx=5, pady=10)

    def achieve_filelist(self, path):
        fileList = os.walk(path)
        return fileList

    # def trim(self, text):
    #     result = re.sub(u"([^\u4e00-\u9fa5\u0030-\u0039\u0041-\u005a\u0061-\u007a])", ",", text)
    #     return result
    #
    # def refactor(self, text):
    #     # 此处的'，'为中文字符，Google translate自动转换了符号
    #     result = text.replace('，', ' ')
    #     return result
    #
    # def filename_translation(self, text):
    #     translate_text = self.translator.translate(self.trim(text), lang_tgt='zh')
    #     return self.refactor(translate_text)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("文件名翻译")
    root.geometry("600x400+600+300")

    app = Application(master=root)
    app.mainloop()
