#-*- coding: utf-8 -*-
import time
import os, threading
import os, sys
from watchdog.observers import Observer
from watchdog.events import *
from watchdog.utils.dirsnapshot import DirectorySnapshot, DirectorySnapshotDiff
from watchdog.events import LoggingEventHandler
from colorama import init, Fore, Back #字體顏色
init(autoreset=True)#字體顏色
import plyer
from plyer import notification
import requests
#調整視窗大小
from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT
import ctypes
import threading
import tkinter as tk
import pystray
from PIL import Image
from pystray import MenuItem, Menu


def quit_window(icon: pystray.Icon):
    icon.stop()
    win.destroy()


def show_window():
    win.deiconify()


def on_exit():
    win.withdraw()


def delete_all():
    text.delete(1.0, "end")      # 刪除全部內容




menu = (MenuItem('顯示', show_window, default=True), Menu.SEPARATOR, MenuItem('退出', quit_window))
image = Image.open("light.ico")
icon = pystray.Icon("icon", image, "Kolink傳真通知", menu)
win = tk.Tk()
win.title('Kolink傳真通知紀錄')
win.geometry("650x130")
win.resizable(width=False, height=False)
text = tk.Text(win)  # 顯示文字
text.config(padx = 1 , pady = 2)
text.config(bg = '#dcdcdc' , fg = '#191970')
scrollbar = tk.Scrollbar(win)          # 建立滾動條
scrollbar.pack(side='right', fill='y')
btn1=tk.Button(win, text="清空畫面", command=delete_all)
btn1.pack(side=tk.BOTTOM)
text.tag_configure("left", justify='left') #輸出文字對齊用


WindowsSTDOUT = windll.kernel32.GetStdHandle(-11)
dimensions = SMALL_RECT(-10, -10, 100, 20) # (left, top, right, bottom)
# Width = (Right - Left) + 1; Height = (Bottom - Top) + 1
windll.kernel32.SetConsoleWindowInfo(WindowsSTDOUT, True, byref(dimensions))



#授權時間
def now():
    return time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
s = '2023-06-14 23:59:59'
if now() > s:
    #notification.notify(title = '歡迎使用Kolink傳真通知', message = 'by F0614 Martin Chung at Coolink CNC Dept. in DEC 2022' ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 5 )
    text.insert(tk.INSERT, 'Welcome to use FAX Notification For Kolink'+ str('\n'))
    text.insert(tk.INSERT, 'by Martin Chung at Coolink CNC Dept. in DEC 2022'+ str('\n'))
else:
    text.insert(tk.INSERT, 'Welcome to use FAX Notification For Kolink'+ str('\n'))

if os.path.isfile('C:/Users/Public/Documents/傳真通知紀錄.csv'): #檢查檔案在不在--公司用
    #print('找到更新紀錄檔')
    text.insert(tk.INSERT, ' '+str('\n'))
    text.insert(tk.INSERT, '找到更新紀錄檔...'+str('\n'))
    text.pack()
else:
    #print('更新記錄檔不存在')
    #print('將會自動建立發行圖面更新紀錄檔')
    text.insert(tk.INSERT, ' '+str('\n'))
    text.insert(tk.INSERT, '更新記錄檔不存在'+ str('\n'))
    text.insert(tk.INSERT, '將會自動建立發行圖面更新紀錄檔...'+ str('\n'))
    text.pack()
    
    with open('C:/Users/Public/Documents/傳真通知紀錄.csv', 'w', encoding = 'cp950') as f:
        a = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        f.write(str(a)+ '\n')

notification.notify(title = 'Kolink傳真通知', message = '啟動偵測中' ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )#--公司用

class FileEventHandler(FileSystemEventHandler):
    def __init__(self, aim_path):
        FileSystemEventHandler.__init__(self)
        self.aim_path = aim_path
        self.timer = None
        self.snapshot = DirectorySnapshot(self.aim_path)
        self.timer = threading.Timer(0.5, self.checkSnapshot)
        self.timer.start()

    
    def on_any_event(self, event):
        #if self.timer:
            #self.timer.cancel()
        self.timer = threading.Timer(0.5, self.checkSnapshot)
        self.timer.start()
    
    def checkSnapshot(self):
           
        snapshot = DirectorySnapshot(self.aim_path)
        diff = DirectorySnapshotDiff(self.snapshot, snapshot)
        self.snapshot = snapshot
        self.timer = threading.Timer(0.5, self.snapshot)
        if diff.files_created == []:
            pass
        else:
            for dc in diff.files_created:
                log = []
                ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))," 通知: ", dc
                # 文字標示所在視窗  
                text.insert(tk.INSERT, str((ans[0][5:]) + ans[1] + ans[2][21:] + '\n'))
                text.tag_add('left', 1.0, "end")  #輸出文字對齊用
                text.yview("end") #滾動到最下面
                text.pack()
                #print(ans[0], Fore.YELLOW + ans[1], Fore.CYAN + ans[2])
                notification.notify(title = 'Kolink傳真通知',  message = str((ans[2][21:])) ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )
                #data = {'message':ans}     # 設定要發送的訊息
                #data = requests.post(url, headers=headers, data=data)   # 使用 POST 方法
                with open('C:/Users/Public/Documents/傳真通知紀錄.csv', 'r', encoding = 'cp950') as f:
                    for txtlog in f:
                        log.append(str(txtlog))
                with open('C:/Users/Public/Documents/傳真通知紀錄.csv', 'w', encoding = 'cp950') as f:
                    for l in log:
                        f.write(str(l))
                    f.write(str(ans[0]) + str(ans[1]) + str((ans[2][15:])) + '\n')
            



class DirMonitor(object):
    """文件夾監視類"""
    
    def __init__(self, aim_path):
        """構造函數"""
        
        self.aim_path= aim_path
        self.observer = Observer()
    
    def start(self):
        """啟動"""
        
        event_handler = FileEventHandler(self.aim_path)
        self.observer.schedule(event_handler, self.aim_path, True)
        self.observer.start()
    
    def stop(self):
        """停止"""
        
        self.observer.stop()
    

 
if __name__ == "__main__":
    monitor = DirMonitor(r'//192.168.0.17/共用資料夾/二樓傳真收件夾 2681-5569')
    monitor.start()
    monitor = DirMonitor(r'//192.168.0.17/共用資料夾/一樓傳真收件夾 2681-3620')
    monitor.start()
    monitor = DirMonitor(r'//192.168.0.17/共用資料夾/夾層傳真收件夾 2687-3701')
    monitor.start()

    win.protocol('WM_DELETE_WINDOW', on_exit)
    threading.Thread(target=icon.run, daemon=True).start()
    scrollbar.config(command=text.yview)

    win.mainloop()

    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        monitor.stop()
        monitor.join()

