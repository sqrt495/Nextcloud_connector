import os
from time import sleep
import stat
import win32com.client
#print(win32com.__gen_path__)
def close_excel_by_force(excel):
    import win32process
    import win32gui
    import win32api
    import win32con

    # Get the window's process id's
    hwnd = excel.Hwnd
    t, p = win32process.GetWindowThreadProcessId(hwnd)
    # Ask window nicely to close
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    # Allow some time for app to close
    sleep(10)
    # If the application didn't close, force close
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
        if handle:
            win32api.TerminateProcess(handle, 0)
            win32api.CloseHandle(handle)
    except:
        pass

path = os.getcwd()
mass = []
ffolder = os.listdir()
for f in ffolder:
    file = path+"\\"+f
    if (".xlsx" not in f) and ((".xls" in f) or (".ods" in f)):
        mass.append(file)
        res = "." + f.split('.')[-1]
        filename = f.replace(res, "")
        print(filename)

        excel = win32com.client.Dispatch('Excel.Application')
        #excel = win32com.client.dynamic.Dispatch('Excel.Application')
        #excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file)
        excel.DisplayAlerts = False

        wb.SaveAs(path+"\\"+"MOD"+filename+".xlsx", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension #FileFormat = 56 is for .xls extension
        wb.Close()
        del wb
        excel.Application.Quit()
        excel.Quit()
        close_excel_by_force(excel)
        del excel



del ffolder, path
sleep(2)
for i in mass:
    os.chmod(i, stat.S_IWRITE)
    os.remove(i)