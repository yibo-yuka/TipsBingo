import tkinter as tk
from tkinter import messagebox
import sqlite3 as sq
import os
import csv
import datetime
import requests as rq
from bs4 import BeautifulSoup as bs
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties as fonprop
font1 = fonprop(fname=r"C:\Users\yibo\Desktop\105787228曲采妮\tips_number\Iansui-Regular.ttf",size = 10)
# region 介面設定
#紀錄發票號碼
win = tk.Tk()
win.geometry("400x300")
win.title("祝發票每張都中")
win.configure(background = "#faf")
myhead = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}


input_area = tk.Canvas(win,width = 200,height = 220,bg = "#00ff88")
other_win = tk.Canvas(win,width = 400,height = 80,bg = "#ffff99")
Font = ("新細明體",10)
en_width = 15
label_x = 10
x_pos = 75
y_pos = [15+i*30 for i in range(6)]

thisyear = datetime.date.today().year
year = []
#月份區間
thismon = datetime.date.today().month
tod = datetime.date.today().day
monTH = ["11~12月","01~02月","03~04月","05~06月","07~08月","09~10月"]
ind = (int(thismon)-1)//2
year.append(thisyear-1911)
if (int(thismon)%2 != 0)&(int(tod)<25):
    ind = ind-1
if ind == 0 or ind-1 == 0:
    year.append(thisyear-1911-1)
month = [monTH[ind],monTH[ind-1]]#如果是1/25前,ind = -1也可以選到09~10月

#年分
yy = tk.StringVar()
yy.set(thisyear-1911)
deyy = yy._default
Y = tk.OptionMenu(input_area,yy,*year)
Y.config(width = 7,height = 1)
Y.place(x = x_pos,y = y_pos[0])
Year = tk.Label(input_area,text = "年份",font = Font)
Year.place(x = label_x,y = y_pos[0]+5)
val = tk.StringVar()

val.set(month[0])
de = val._default
M = tk.OptionMenu(input_area,val,*month)
M.config(width = 7,height = 1)
M.place(x = x_pos,y = y_pos[1])
Month = tk.Label(input_area,text = "月份區間",font = Font)
Month.place(x = label_x,y = y_pos[1]+5)
#號碼
nums = tk.Entry(input_area,width = en_width)
nums.place(x = x_pos,y = y_pos[2]+5)
Nums = tk.Label(input_area,text = "發票號碼",font = Font)
Nums.place(x = label_x,y = y_pos[2]+5)
#花費
costs = tk.Entry(input_area,width = en_width)
costs.place(x = x_pos,y = y_pos[3]+5)
Costs = tk.Label(input_area,text = "花費",font = Font)
Costs.place(x = label_x,y = y_pos[3]+5)
#備註(消費類型)
pp = tk.StringVar()
pp.set('Other')
ppde = pp._default
ppty = ["Eat","Cloth","Trival","Home","Book","Workout","DaliyUse","Comic","Movie","Other"]
ps = tk.OptionMenu(input_area,pp,*ppty)
ps.config(width = 7,height = 1)
ps.place(x = x_pos,y = y_pos[4]+5)
Ps = tk.Label(input_area,text = "消費類型",font = Font)
Ps.place(x = label_x,y = y_pos[4]+10)
#加入資料
BingoDetail = []
total = 0
tax = 0
all_you_cost = 0
yearcol = []
monthcol = []
numcol = []
costcol = []
pscol = []
def store_detail():
    global val,month,BingoDetail,total,tax,all_you_cost,yearcol,monthcol,numcol,costcol,pscol
    no_content = False
    info = []
    n = nums.get()
    c = costs.get()
    p = pp.get()
    v = val.get()
    yv = yy.get() 
    if (n == ''):
        info.append("請輸入發票號碼!!")
        no_content = True
    if (len(n) != 8):
        info.append("請輸入發票8碼數字")
        no_content = True
    if (c == ''):
        info.append("請輸入花費金額!!")
        no_content = True
    
    if no_content:
        text = ''
        for i in info:
            text = text+f"{i}\n"
        messagebox.showerror('showerror', text)
    if not no_content:
        yearcol.append(yv)#年份
        monthcol.append(v)#月份區間
        numcol.append(n)#發票號碼
        costcol.append(int(c))#花費
        pscol.append(p)#備註
        
        if v == month[0]:
            url = 'https://invoice.etax.nat.gov.tw/index.html'
            
        else:
            url = "https://invoice.etax.nat.gov.tw/lastNumber.html"
            
        info = rq.get(url,headers = myhead)
        info.encoding = "utf-8"
        Info = bs(info.text,"html5lib")
        
        #月份
        mon = Info.find("a","etw-on").text.strip()
        print(mon)
        #頭獎
        num = Info.find("div","container-fluid etw-bgbox mb-4")
        Num = num.find_all("span","font-weight-bold etw-color-red")
        preFixs = num.find_all("span","font-weight-bold")
        prefixs = []
        for i in preFixs:
            prefixs.append(i.text.strip())
        #['63603594', '73155944', '94985', '899', '57283', '420', '62825', '278', '63603594', '73155944', '94985', '899', '57283', '420', '62825', '278']
        del_index = []
        for i in range(len(prefixs)):
            if len(prefixs[i]) == 5:
                del_index.append(i)
                continue
            if len(prefixs[i]) == 3:
                prefixs[i] = prefixs[i-1]+prefixs[i]
                print(prefixs[i])
        minus_index = 0
        for i in del_index:
            i-=minus_index
            del prefixs[i]
            minus_index+=1#每少掉一個，要往前遞補的位子數
            
        prefixs = prefixs[:5]
        dic = {"type":["特別獎","特獎","頭獎","頭獎","頭獎"],"number":prefixs}
        df = pd.DataFrame(dic)
        print(df)
        
        writer = pd.ExcelWriter("tips_number\中獎號碼.xlsx",mode = "a",engine = "openpyxl")
        wr = pd.ExcelFile("tips_number\中獎號碼.xlsx")
        s_name = wr.sheet_names
        if mon[:10] not in s_name:
        #wb = writer.book
            df.to_excel(writer,mon[:10],index = False)
        writer.close()
        
        
        #check bingo
        stand = dic["number"]
        if n == stand[0]:
            print("中特別獎!!")
            pr = "中特別獎!!"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+(10**7)
                tax = tax+(10**7)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n == stand[1]:
            print("中特獎!!")
            pr = "中特獎!!"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+2*(10**6)
                tax = tax+2*(10**6)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n == stand[2] or n == stand[3] or n == stand[4]:
            print("頭獎")
            pr = "頭獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+2*(10**5)
                tax = tax+2*(10**5)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n[1:] == stand[2][1:] or n[1:] == stand[3][1:] or n[1:] == stand[4][1:]:
            print("二獎")
            pr = "二獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+4*(10**4)
                tax = tax+4*(10**4)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n[2:] == stand[2][2:] or n[2:] == stand[3][2:] or n[2:] == stand[4][2:]:
            print("三獎")
            pr = "三獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+(10**4)
                tax = tax+(10**4)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n[3:] == stand[2][3:] or n[3:] == stand[3][3:] or n[3:] == stand[4][3:]:
            print("四獎")
            pr = "四獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+4*(10**3)
                tax = tax+4*(10**3)*20*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n[4:] == stand[2][4:] or n[4:] == stand[3][4:] or n[4:] == stand[4][4:]:
            print("五獎")
            pr = "五獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+(10**3)
                tax = tax+(10**3)*0.4*0.01
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        elif n[5:] == stand[2][5:] or n[5:] == stand[3][5:] or n[5:] == stand[4][5:]:
            print("六獎")
            pr = "六獎"
            if f"{v}-{n}-{c}元" not in BingoDetail:
                total = total+2*(10**2)
                all_you_cost = all_you_cost+int(c)
                bd = f"{v}-{n}-{c}元"
                list_bingo.insert(tk.END, bd)
                BingoDetail.append(f"{v}-{n}-{c}元")
        else:
            print("no")
            all_you_cost = all_you_cost+int(c)
            pr = "沒中..."
        
        px = tk.StringVar()
        #p.set("")
        px.set(f"{pr},\n目前獎金{total}元,\n稅金{tax}元,\n共{total-tax}元")
        praze = tk.Label(other_win,textvariable = px,font =("微軟正黑體",10),bg = "#ff8")
        praze.place(x = 30,y = 5)
        all_costs = tk.StringVar()
        #p.set("")
        all_costs.set(f"目前花費\n{all_you_cost}\n元")
        All_costs = tk.Label(other_win,textvariable = all_costs,font =("微軟正黑體",12),bg = "#ff8")
        All_costs.place(x = 200,y = 5)
        costs.delete(0,"end")
        nums.delete(0,"end")
        
        
        
def analy_data():
    global yearcol,monthcol,numcol,costcol,pscol,thisyear,ind
    
    datA = {"Year":yearcol,"Month":monthcol,"Numbers":numcol,"Costs":costcol,"Type":pscol}
    df = pd.DataFrame(datA)
    if len(df["Month"].unique().tolist())>1:
        diff_m = df.groupby("Month")
        lastm = diff_m.get_group(month[ind-1])
        thism = diff_m.get_group(month[ind])
        #d1 = pd.pivot_table(lastm,index = "消費類型",values = "消費金額",aggfunc = "sum")
        #d2 = pd.pivot_table(thism,index = "消費類型",values = "消費金額",aggfunc = "sum")
        fig,axe = plt.subplots(nrows = 1,ncols = 2)
        
        f1 = lastm.groupby("Type").sum().plot(ax = axe[0],kind = "pie",y = "Costs",legend=False,title = "consume analyze in "+str(month[ind-1])[:5],autopct = "%.1f%%",labeldistance = 0.6,pctdistance = 0.3)
        #plt.title(str(month[ind-1])+"consume analyze",fontproperties = font1)
        #plt.ylabel("Costs",fontproperties = font1)
        #plt.show()
        #f1.set(ylabel = "消費金額")
        f2 = thism.groupby("Type").sum().plot(ax = axe[1],kind = "pie",y = "Costs",legend=False,title = "consume analyze in "+str(month[ind])[:5],autopct = "%.1f%%",labeldistance = 0.6,pctdistance = 0.3)
        #plt.title(str(month[ind])+"consume analyze",fontproperties = font1)
        #f2.set(ylabel = "消費金額")
        #plt.ylabel("Costs",fontproperties = font1)
        plt.show()
    else:
        #d = pd.pivot_table(df,index = "消費類型",values = "消費金額",aggfunc = "sum")
        d = df.groupby("Type").sum()
        mm = df["Month"].unique().tolist()[0]
        f = d.plot(kind = "pie",y = "Costs",autopct = "%.1f%%",labeldistance = 0.6,pctdistance = 0.3)
        plt.title(str(mm)+"consume analyze",fontproperties = font1)
        
        plt.ylabel("Costs",fontproperties = font1)
        plt.show()
#Get_num = tk.Button(other_win,text = "獲得中獎號碼",font = ("新細明體",10),width = 25)
#Get_num.place(x = 10,y = 25)
bingo = tk.Button(input_area,text = "對獎",font = ("新細明體",15),width = 10,command = store_detail)
bingo.place(x = 40,y = 180)

input_area.place(x = 0,y = 0)
other_win.place(x = 0,y = 220)
#analyze = tk.Button(win,text = "消費分析")
#analyze.place(x = 222,y = 20)
#saveall_analy = tk.Button(win,text = "儲存所有\n分析圖")
#saveall_analy.place(x = 222,y = 60)
frame = tk.Frame(win, width=15)        # 加入頁框元件，設定寬度
frame.place(x = 205,y = 60)

scrollbar = tk.Scrollbar(frame)         # 在頁框中加入捲軸元件
scrollbar.pack(side='right', fill='y')  # 設定捲軸的位置以及填滿方式


det = tk.StringVar()
det.set(BingoDetail)
List_Bingo = tk.Label(win,text = "中獎發票明細",font = ("微軟正黑體",15,"bold"),bg = "#ffff00")
List_Bingo.place(x = 235,y = 20)
list_bingo = tk.Listbox(frame,width = 25,height = 8)
list_bingo.pack(side='left', fill='y')

scrollbar.config(command = list_bingo.yview)

def destroy():
    win.destroy()

comsume_analy = tk.Button(other_win,text = "消費分析",command = analy_data)
comsume_analy.place(x = 322,y = 10)
close_win = tk.Button(other_win,text = "關閉視窗",command = destroy)
close_win.place(x = 322,y = 45)

# endregion

win.mainloop()
print("window closed.")