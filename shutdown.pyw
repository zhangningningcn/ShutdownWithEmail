# -*- coding: utf-8 -*-
"""
此脚本通过邮件远程关闭电脑
shutdown -s 直接后台运行
shutdown.conf 配置文件
maillist.txt 邮件UID列表

"""
import getpass, poplib
import re
import base64
import time
import os
import sys
from tkinter import *
#import tkinter
import threading
import ctypes
import ctypes.wintypes
import win32con

pop3addr = "test"
mailaddr = None
mailpasswd = None
cmdmailaddr = None
cmdstr = None
checkinterval = None


def readconf():
    #return "ok"
    gvars = ("pop3addr","mailaddr","mailpasswd","cmdmailaddr","cmdstr","checkinterval")
    global pop3addr,mailaddr,mailpasswd,cmdmailaddr,cmdstr,checkinterval
    try:
        f = open("shutdown.conf")
    except:
        f = open("shutdown.conf","w")
        writestr = """
pop3addr=pop3.139.com
#我用的是139邮箱，其他邮箱理论上也支持
mailaddr=邮箱地址
mailpasswd=邮箱密码
cmdmailaddr=发送命令邮箱
cmdstr=命令字符串
# 命令字符串'#'后面的部分忽略。
# 查时间询间隔 600秒
checkinterval=600
# 注意：等号两端不能有空格，行尾也不能有多余空格。
        """
        f.write(writestr)
        f.close()
        return "读取配置参数失败，\n请配置shutdown.conf"
    while True:
        line = f.readline()
        if not line:
            break
        line = line.strip()
        #line = line.replace(" ","")
        #line = line.replace('"','')
        d = line.split("=")
        if len(d) == 2 and (d[0] in gvars):
            if '#' in d[1]:
                d[1] = d[1][:d[1].index("#")]
            cmdline = 'global '+ d[0] + ';' + d[0] + ' = "' + d[1] +'"'
            exec(cmdline)
    f.close()



    if not '@' in mailaddr:
       mailaddr = None
    if not '@' in cmdmailaddr:
       cmdmailaddr = None

    if (mailaddr == None) or (pop3addr == None) or (mailpasswd == None):
        return "读取配置参数失败，\n请配置shutdown.conf"
    if cmdmailaddr == None or cmdstr == None or checkinterval == None:
        return "读取配置参数失败，\n请配置shutdown.conf"
    try:
        checkinterval = int(checkinterval)
    except:
        return "checkinterval 需要配置成整数"
    return "ok";

def getmail():
    print("run " + time.strftime('%H:%M:%S',time.localtime()))
    #return 0
    maillist = []
    try:
        #这个文件保存邮件UID，用于判断邮件之前是否读取过
        f = open("maillist.txt")
        line = f.readline()
        line = line.strip()
        maillist = line.split(",")
        f.close()
    except:
        pass
    try:
        M = poplib.POP3(pop3addr)
        M.user(mailaddr)
        M.pass_(mailpasswd)
    except:
        return 1
    uidl = M.uidl()[1]
    uidl = {i.decode("ascii").split(" ")[1]:i.decode("ascii").split(" ")[0] for i in uidl}
    tempuidl = [i for i in uidl]
    subuidl = set(tempuidl) - set(maillist)
    mlist = M.list()[1]
    mlist = {i.decode("ascii").split(" ")[0]:i.decode("ascii").split(" ")[1] for i in mlist}
    print(mlist)

    pattern = re.compile(r"Received.*from\s*<(.*)>")#.from\s*<()>
    state = 1
    maildatestr = ""
    for uid in subuidl:
        i = uidl[uid]
        #大于100k的邮件不处理
        if int(mlist[i]) > 100000:
            continue
        i = int(i)
        base64msg = b''
        shutdown = False
        for j in M.retr(i)[1]:
            # print(j)
            if state < 3:
                m = pattern.match(j.decode(encoding="ascii",errors="ignore"))
                if m:
                    if state == 1:
                        if m.group(1) == cmdmailaddr:
                            state = 3
            else:
                if state == 3:
                    if not j:
                        state = 4
                elif state == 4:
                    if j:
                        base64msg += j
                    else:
                        state = 1
                        strmsg = base64.b64decode(base64msg).decode(encoding="gb2312",errors="ignore")
                        # '#'后面的字符串忽略
                        if '#' in strmsg:
                            strmsg = strmsg[:strmsg.index("#")]
                        if strmsg == cmdstr:
                            shutdown = True
                            break
        if shutdown:
            os.system("shutdown -s -t 300")
            break
    M.quit()
    if subuidl:
        try:
            pre = str([i for i in uidl])
            pre = pre.strip("[]")
            pre = pre.replace("'","")
            pre = pre.replace(" ","")
            f = open("maillist.txt","w")
            f.write(pre)
            f.close()
        except:
            return 2
    return 0

class BackConnTimer(threading.Thread):
    def run(self):
        self.stop = False
        self.lastgetmail = self.tmnow
        self.confstr = "最后一次查询时间 " + time.strftime('%H:%M:%S',self.tmnow)
        delay = time.mktime(time.localtime()) - time.mktime(self.tmnow)
        delay = checkinterval - delay
        if delay < 1:
           delay = 1
        time.sleep(delay)
        while not self.stop:
            #getmail()
            self.lastgetmail = time.localtime()
            res = getmail()
            if res == 0:
               tmnow = time.localtime()
               self.confstr = "最后一次查询时间 " + time.strftime('%H:%M:%S',tmnow)
            elif res == 1:
               self.confstr = '读取邮件失败'
            elif res == 2:
                self.confstr = '写文件失败'
            time.sleep(checkinterval)
    def stopthread(self):
        self.stop = True
    def getlastgetmail(self):
        return self.lastgetmail
    def setastgetmail(self,tmnow):
        self.tmnow = tmnow
    def getconfstr(self):
        return self.confstr

def runbackup():
    global root,b_runbackguound,user32,keyinfo
    if not user32.RegisterHotKey(None, 99, win32con.MOD_CONTROL, win32con.VK_F7):
        keyinfo["text"] = "无法注册快捷键Ctrl+F7"
    else:
        keyinfo["text"] = "Ctrl+F7调回前台"
        #messagebox.showerror("无法后台运行","注册热键失败")
        return
    root.destroy()
    b_runbackguound = True

def UI_Update():
    global root,tmnow,disinfo
    res = getmail()
    if res == 0:
       tmnow = time.localtime()
       confstr = "最后一次查询时间 " + time.strftime('%H:%M:%S',tmnow)
    elif res == 1:
       confstr = '读取邮件失败'
    elif res == 2:
        confstr = '写文件失败'
    disinfo["text"] = confstr
    root.after(checkinterval*1000,UI_Update)

if __name__ == "__main__":
    global root,b_runbackguound,guidisplay,user32,tmnow,disinfo,keyinfo

    user32 = ctypes.windll.user32
    confstr = readconf()
    if confstr == "ok":
        #confstr = "关机程序已运行"
        res = getmail()
        if res == 0:
           tmnow = time.localtime()
           confstr = "最后一次查询时间 " + time.strftime('%H:%M:%S',tmnow)
           delay1 = checkinterval*1000
        elif res == 1:
           confstr = '读取邮件失败'
        elif res == 2:
            confstr = '写文件失败'
    else:
        delay1 = -1
    guidisplay = True
    
    start_s = False
    if len(sys.argv) == 2:
       if sys.argv[1] == "-s":
          start_s = True
          
    while guidisplay:
        b_runbackguound = False
        guidisplay = False
        keyerror = False

        if start_s:
            b_runbackguound = True
            start_s = False
            
        if b_runbackguound:
            if not user32.RegisterHotKey(None, 99, win32con.MOD_CONTROL, win32con.VK_F7):
                b_runbackguound = False
                keyerror = True
                
        if not b_runbackguound:
            root = Tk()
            if delay1 >= 0:
               if delay1 < 500:
                  delay1 = 500
               root.after(delay1,UI_Update)
            root.title("shutdown")
            root.geometry('200x100')                 #是x 不是*
            root.resizable(width=False, height=True) #宽不可变, 高可变,默认为True
            disinfo = Message(root, text=confstr,width = 150, font=('Arial', 10))
            disinfo.pack()
            if delay1 < 0:
                Button(root, text="后台运行", state = "disabled").pack()
            else:
                Button(root, text="后台运行", command = runbackup).pack()
            keyinfo = Message(root, text="Ctrl+F7调回前台",width = 150, font=('Arial', 10))
            keyinfo.pack()
            if keyerror:
                keyinfo["text"] = "无法注册快捷键Ctrl+F7"
            root.mainloop()

            
        if b_runbackguound:
            backrun = BackConnTimer()
            backrun.setastgetmail(tmnow)
            backrun.start()
            try:
                msg = ctypes.wintypes.MSG()
                #print msg
                while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                    if msg.message == win32con.WM_HOTKEY:
                        if msg.wParam == 99:
                            guidisplay = True
                            delay1 = time.mktime(time.localtime()) - time.mktime(backrun.getlastgetmail())
                            delay1 = (checkinterval - int(delay1))*1000
                            confstr = backrun.getconfstr()
                            backrun.stopthread()
                            break
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageA(ctypes.byref(msg))
            finally:
                user32.UnregisterHotKey(None, 99)
