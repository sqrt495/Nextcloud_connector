import os
from time import sleep
import stat
import openpyxl
from time import sleep


def divide_replacer(path, f, method):
    #print(method)
    sleep(1)
    try:
        os.mkdir(path + "\\" + method + "\\")
        os.replace(path+"\\"+f, path + "\\" + method + "\\" + f)
    except:
        os.replace(path+"\\"+f, path + "\\" + method + "\\" + f)


def divide_check_headers(path, f):
    wb = openpyxl.load_workbook(path+"\\"+f, data_only=True)
    for ws in wb:
        try:
            for n, i in enumerate(ws):
                if n <= 20:
                    for j in i:
                        if j.value == 'Диспансерное наблюдение':
                            return "old"
                        elif j.value == 'Подстатус':
                            return "new"
        except:
            pass
    return "else"