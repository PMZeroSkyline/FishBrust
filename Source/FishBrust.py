#免责声明 不得使用此软件发送盗版软件，不得使用本软件从事任何非法行为，本软件为免费软件，本软件作者不承担任何责任

import smtplib
import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import webbrowser
import os
import re
import json
import ctypes
import base64
from icon import img

window = tk.Tk()
window.title('咸鱼突刺！附件拆分发送器(╯°□°）╯︵ 📧')
window.geometry('400x520')
window.configure(bg='#3c3f41')
window.resizable(width=False, height=False)

tmp = open("tmp.ico","wb+")
tmp.write(base64.b64decode(img))
tmp.close()
window.iconbitmap("tmp.ico")
os.remove("tmp.ico")

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MY_GUI")

menubar = tk.Menu(window)
setting_menu = tk.Menu(menubar, tearoff=0, bg="#3c3f41", fg='#afb1b3')

menubar.add_cascade(label='预设', menu=setting_menu)

def newSetting():
    e1.delete(0, 'end')
    e2.delete(0, 'end')
    e3.delete(0, 'end')
    showinfo('预设已还原', '~ o(*￣▽￣*)ブ鱼干已销毁')

setting_menu.add_command(label='还原预设', command=newSetting)

def loadSetting():
    filepath = fd.askopenfilename(defaultextension=".json",filetypes = [('JSON', '.json')])
    maildata = json.load(open(filepath))
    e1.delete(0, 'end')
    e2.delete(0, 'end')
    e3.delete(0, 'end')
    e1.insert(0, maildata['from'])
    e2.insert(0, maildata['to'])
    e3.insert(0, maildata['key'])
    showinfo('预设已加载', '~ o(*￣▽￣*)ブ鱼干已就绪')

setting_menu.add_command(label='载入预设', command=loadSetting)

def saveSetting():
    filepath = fd.asksaveasfilename(defaultextension=".json",filetypes = [('JSON', '.json')])
    maildata = {
        'from': e1.get(),
        'to': e2.get(),
        'key': e3.get()
    }
    jsondata = json.dumps(maildata)
    with open(filepath, 'w') as file:
        file.write(jsondata)
        file.close()
    showinfo('预设已保存', '~ o(*￣▽￣*)ブ鱼干已保存')

setting_menu.add_command(label='保存预设', command=saveSetting)

help_menu = tk.Menu(menubar, tearoff=0, bg="#3c3f41", fg='#afb1b3')
menubar.add_cascade(label='帮助', menu=help_menu)

def UseTips():
    showinfo("使用说明", "建议把大文件，用解压软件，压缩成多个小压缩包，加一下解压密码，添加并发送")

help_menu.add_command(label='使用说明', command=UseTips)

def smtpTips():
    showinfo("填写SMTP服务授权码", "以QQ邮箱为例 请在 邮箱设置->账户->开启服务：开启SMTP服务即可获得授权码，为了您的邮箱安全 建议您使用完毕之后关闭SMTP服务")
    webbrowser.open('https://service.mail.qq.com/cgi-bin/help?subtype=1&&id=28&&no=1001256')

help_menu.add_command(label='获取SMTP服务授权码', command=smtpTips)

window.config(menu=menubar)

l = tk.Label(text = "发件人：", height=2, bg="#3c3f41", fg='#afb1b3')
l.pack()

e1 = tk.Entry(window, show = None, width=50, bg="#2b2b2b", fg='#3592c4')
e1.pack()

l = tk.Label(text = "收件人：", height=2, bg="#3c3f41", fg='#afb1b3')
l.pack()

e2 = tk.Entry(window, show = None, width=50, bg="#2b2b2b", fg='#3592c4')
e2.pack()

l = tk.Label(text = "SMTP服务授权码：", height=2, bg="#3c3f41", fg='#afb1b3')
l.pack()

e3 = tk.Entry(window, show = None, width=50, bg="#2b2b2b", fg='#3592c4')
e3.pack()

l = tk.Label(text = "附件列表：", height=2, bg="#3c3f41", fg='#afb1b3')
l.pack()

lb = tk.Listbox(window, width=50,height=10, bg="#2b2b2b", fg='#499c54')
lb.pack()

def fun_b1():
    filepaths = fd.askopenfilenames(title = '选择所有要发送的文件')
    for filepath in filepaths:
        lb.insert('end', filepath)
b1 = tk.Button(window, text='添加附件', width=10,height=1, command=fun_b1, bg="#3c3f41", fg='#afb1b3')
b1.place(x=50, y=450)

def fun_b2():
    lb.delete(0,'end')
b2 = tk.Button(window, text='清空附件', width=10,height=1, command=fun_b2, bg="#3c3f41", fg='#afb1b3')
b2.place(x=150, y=450)

def fun_send(filepath):
    em = MIMEMultipart()
    em['Subject'] = '咸鱼突刺'
    em['from'] = e1.get()
    em['to'] = e2.get()
    content = MIMEText("来了, 来了________" + os.path.basename(filepath) + "________冲进来了!")
    em.attach(content)
    app = MIMEApplication(open(filepath, mode="rb").read())
    app.add_header('content-disposition', 'attachment', filename=os.path.basename(filepath))
    em.attach(app)
    smtp = smtplib.SMTP()
    findMail = e1.get().find('@')
    mailType = e1.get()[findMail + 1 : len(e1.get())]
    connetHost = 'smtp.' + mailType
    smtp.connect(connetHost)
    smtp.login(e1.get(), e3.get())
    smtp.send_message(em)
    smtp.close()

def fun_b3():
    p = re.compile('^[A-Za-z0-9\u4e00-\u9fa5]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$')
    if len(e1.get()) == 0:
        showinfo("请填写发件人邮箱", "在发件人处填写你的邮箱")
        return
    if p.match(e1.get()) == None:
        showinfo("邮箱格式不正确", "发件人邮箱格式不正确")
        return
    if len(e2.get()) == 0:
        showinfo("请填写收件人邮箱", "在收件人处填写你的邮箱")
        return
    if p.match(e2.get()) == None:
        showinfo("邮箱格式不正确", "收件人邮箱格式不正确")
        return
    if len(e3.get()) == 0:
        smtpTips()
        return
    filepaths = lb.get(0,'end')
    if len(filepaths) == 0:
        showinfo("请添加附件", "请点击 “添加附件” 按钮 选择传输的文件")
        return
    try:
        showinfo("开冲！开冲！即将开始传输", "关闭此对话框后文件自动上传 文件上传需要消耗一段时间，我没做进度条，发送完成会有消息提示，请勿以为程序卡死")
        for filepath in filepaths:
            fun_send(filepath)
        showinfo("发送完成", "文件发送完成")
        fun_b2()
    except:
        showinfo("文件发送失败", "请检查SMTP服务授权码")

b3 = tk.Button(window, text='发送邮件', width=10,height=1, command=fun_b3, bg="#3c3f41", fg='#afb1b3')
b3.place(x=250, y=450)

l = tk.Label(text = "免责声明:不得使用本软件从事任何非法行为,本软件作者不承担任何责任", height=2, font=('Arial', 9), bg="#3c3f41", fg='#afb1b3')
l.pack()

#window.iconbitmap('fish.ico')
window.mainloop()