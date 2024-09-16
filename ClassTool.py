import random as rd
import pandas as pd
import tkinter as tk
df = pd.read_excel('NameList.xlsx')
NameList = df['姓名'].tolist()
flag=True

#抽取
def random():
    RandList=''
    global NameList
    i=0
    while i<int(enter.get()):
        try:
            RandNum=rd.choice(NameList)
        except:
            NameList = df['姓名'].tolist()
            RandNum=rd.choice(NameList)
        RandList+=RandNum+" "
        NameList.remove(RandNum)
        i+=1

    #显示
    global mes
    mes.pack_forget()
    mes=tk.Message(windows,font=('微软雅黑',20),text=RandList,bg="white",width=windows.winfo_width()*0.9)
    mes.pack(anchor="center")

def close():
    global flag,sub
    flag=True
    sub.destroy()

#替换
def replace():
    global flag,sub
    if flag:
        flag=False
        sub=tk.Toplevel()
        sub.configure(bg="white")
        sub.title("替换")
        sub.iconbitmap()
        sub.protocol("WM_DELETE_WINDOW",close)
        CB=tk.Checkbutton(sub)
        CB.pack()
        
#创建窗口
windows=tk.Tk()
#窗口名称
windows.title("Class Tool")
#窗口图标
windows.iconbitmap()
#窗口比例
windows.geometry(str(int(windows.winfo_screenwidth())//2)+'x'+str(int(windows.winfo_screenheight())//2))
#窗口颜色
windows.configure(bg="white")

LB=tk.LabelFrame(windows,bg='white',borderwidth=0)

replace=tk.Button(LB,text="替换",command=replace,font=('微软雅黑'))
replace.pack(side='left',anchor='w',padx=100)

bm=tk.Button(LB,text="抽取",command=random,width=10,font=('微软雅黑'))
bm.pack(side=tk.LEFT,anchor=tk.S)

enter=tk.Entry(LB,bg="#FEFEFE",borderwidth=2,width=5,font=('微软雅黑'))
enter.pack(side=tk.LEFT,anchor=tk.S)

n=tk.Label(LB,text="<--抽取人数",font=('微软雅黑'),bg='white')
n.pack(side=tk.LEFT,anchor=tk.S)

LB.pack(side='bottom',anchor='w',pady=10)


#init
mes = tk.Message(windows,text='',bg="white")
mes.place(x=100, y=200)

windows.mainloop()
