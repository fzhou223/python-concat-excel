# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import os
import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog

def main():
    try:

        selected_path = select_path()

        filesArrary = list_all_files(selected_path)
        dataframe = concat_file_to_dataframe(filesArrary)

        output_path = os.path.join(selected_path,'output')
        mkdir(output_path)

        output_file = os.path.join(output_path,'output.xlsx')
        output(dataframe, output_file)

        tkinter.messagebox.showinfo('成功','合併成功，請在 output 中查看')
    except:
        tkinter.messagebox.showinfo('失敗','合併失敗')

# 選擇檔案位於的資料夾
def select_path():
    root = tk.Tk()
    root.withdraw()
    Fpath = tkinter.filedialog.askdirectory(title='please select the path /　請選擇需要合併檔案的位置')
    return Fpath

# 創建資料夾
def mkdir(path):
    # 清理 path 傳入的字符
    path = path.strip()
    path = path.rstrip('\\')

    # 判斷路徑（True：存在；False：不存在）
    isExists = os.path.exists(path)

    # 判斷結果
    if not isExists:
        os.makedirs(path)
        print('新增目録：',path)
        return True
    else:
        print('目録已經存在')
        return False

# 列出文件
def list_all_files(rootdir):
    _files = []
    _list = []

    # 判斷後綴名
    for i in os.listdir(rootdir):
        if os.path.splitext(i)[1] == '.xlsx':
            _list.append(i)
        else:
            print('跳過檔案：',i)

    for i in range(0,len(_list)):
        path = os.path.join(rootdir,_list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path))
        if os.path.isfile(path):
            _files.append(path)

    return _files

# 合并 excel 至 dataframe
def concat_file_to_dataframe(filesArrary):
    _frames = []

    for file in filesArrary:
        _df = pd.read_excel(file)
        _frames.append(_df)
        print('發現檔案：',file)

    _result = pd.concat(_frames, sort= False)
    print('合並後行數：',len(_result))
    _result_drop_duplicates = _result.drop_duplicates()
    print('去重後行數：',len(_result_drop_duplicates))
    return _result_drop_duplicates

# 輸出 dataframe 至 excel
def output(dataframe,outputDir):
    dataframe.to_excel(outputDir, sheet_name = 'data', index = None)

if __name__ == '__main__':
    main()