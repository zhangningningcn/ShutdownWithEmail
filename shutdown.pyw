# -*- coding: utf-8 -*-
"""
此脚本通过邮件远程关闭电脑


"""
import getpass, poplib
import re
import base64
import time
import os



pop3addr = "pop3.139.com" #我用的是139邮箱，其他邮箱理论上也支持
mailaddr = "邮箱地址"
mailpasswd = "邮箱密码"
cmdmailaddr = "发送命令邮箱"
cmdstr = "命令字符串"
# 命令字符串'#'后面的部分忽略。
checkinterval = 600 #查时间询间隔 600秒 

def getmail():
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
        print("login error")
        return
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
    shutdown = False
    for uid in subuidl:
        i = uidl[uid]
        #大于100k的邮件不处理
        if int(mlist[i]) > 100000:
            continue
        i = int(i)
        base64msg = b''
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
            pass
if __name__ == "__main__":
    while True:
        getmail()
        #break
        time.sleep(checkinterval)
