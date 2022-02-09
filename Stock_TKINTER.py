import tkinter as tk
import sqlite3
from PIL import ImageTk, Image
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import matplotlib.pyplot as plt

################模板設定############################
win=tk.Tk()
canvas=tk.Canvas(win,width=700,height=100,bd=0,highlightthickness=0)
imgpath='money.gif'
img=Image.open(imgpath)
photo=ImageTk.PhotoImage(img)
win.geometry('+450+100')
win.title('存股Searcher')
canvas.create_image(350,49,image=photo)
canvas.pack()
win.iconbitmap('Pauloruberto-Custom-Round-Yosemite-Python.ico')


################區塊設定############################
frame1=tk.Frame(win)
frame1.pack(fill='x')
frame2=tk.Frame(win)
frame2.pack(fill='x')
frame3=tk.Frame(win)
frame3.pack(fill='x')
frame4=tk.Frame(win)
frame4.pack(fill='x')
frame5=tk.Frame(win,bg='pink')
frame5.pack(fill='x')
frame6=tk.Frame(win)
frame6.pack(fill='x')
frame7=tk.Frame(win)
frame7.pack(fill='x')


################主程式區段############################
def showmain():

    #######搜尋函式######
    def stock_search():
        global data
        if conn==None:
            label.config(text='資料庫未開啟!',fg='red')
            return
        try:
            with sqlite3.connect(file_name) as con:             
                df=pd.read_sql_query('select * from "deposit"',con=con)
                stock=pd.read_sql_query(f'select * from "{stock_num_ent.get()}"',con=con)
            label.config(text='搜尋完成!',fg='yellow')
        except Exception as e:
            print(e)
            label.config(text='資料庫(表)讀取失敗!!',fg='red')
        if stock_num_ent.get()=='':
            label.config(text='請輸入股號!!',fg='red')
            return
        
        #--------股名-------#
        try:
            df[df['編號']==f'{stock_num_ent.get()}']['股名'].values[-1]
        except Exception as e:
            print(e)
            stock_name.config(text='無此股票',font=('華康正顏楷體W5',24),fg='black')
            label.config(text='請確認股票編號!!',fg='red')
            equity_result.config(text='',font=('Arial',15),fg='black')
            dividendyear_result.config(text='',font=('Arial',15),fg='black')
            closeing_result.config(text='',font=('Arial',15),fg='black')
            listVar1.set([])
            listVar2.set([])
            total_result.config(text='',font=('Arial',18),fg='black')
            per.config(text='',font=('Arial',18),fg='black')
            money_result.config(text='',font=('Arial',18),fg='black')
            
        stock_name.config(text=df[df['編號']==f'{stock_num_ent.get()}']['股名'].values[-1],font=('華康正顏楷體W5',24),fg='blue')
        
        #-------股本--------#
        url=f'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID={stock_num_ent.get()}'
        try:
            chrom=get_chrom(url)
            soup2=BeautifulSoup(chrom.page_source,'lxml')
            equity=soup2.find('table',class_="b1 p4_4 r10").find_all('tr')[4].find_all('td')[1].text.strip()
            equity_result.config(text=equity,font=('Arial',15),fg='red')
        except Exception as e:
            print(e)
            equity_result.config(text='此股無股本',font=('Arial',15),fg='red')
            
        #------配息年數------#
        try:
            year=soup2.find('table',id="FINANCE_DIVIDEND").find_all('nobr')[-2].text.strip()
            test=year.split('續')[1].split('年')[1].split('發')
            if test[0]=='配':
                ans=year.split('續')[1].split('年')[0]
                dividendyear_result.config(text=f'連續配息{ans}年',font=('Arial',14),fg='red')
            else:
                dividendyear_result.config(text='無連續配息',font=('Arial',14),fg='red')
            chrom.quit()
        except Exception as e:
            print(e)
            label.config(text='查詢失敗!!',fg='red')   
            
        #----前一日收盤價----#
        try:
            sql=f'select * from "{stock_num_ent.get()}"'
            datas=list(cursor.execute(sql))
            conn.commit()
            close=round(datas[-1][-2],2)
            closeing_result.config(text=close,font=('Arial',15),fg='red')
        except Exception as e:
            print(e)
            label.config(text='資料表內無該股收盤資料!!',fg='red')
        
        #------董監持股------#
        try:
            url=f'https://goodinfo.tw/tw/StockDirectorSharehold.asp?STOCK_ID={stock_num_ent.get()}'
            chrom=get_chrom(url)
            soup3=BeautifulSoup(chrom.page_source,'lxml')
            test=soup3.find('table',class_="b1 p4_0 r0_10 row_bg_2n row_mouse_over").find_all('tr',align="center")[1:]
            board_datas=[]
            for i in range(3):
                year=test[i].find('a').text.strip()
                percent=test[i].find_all('td')[-5].text.strip()
                board_datas.append([f'{year},董監持股:{percent}%'])
            listVar1.set(board_datas)
            chrom.quit()
        except Exception as e:
            print(e)
            label.config(text='查詢失敗!!',fg='red')
        
        #-----最高及最低-----#
        try:
            datas=[]
            for value in df['除息日'].values:
                if '即將除息' in value:
                    df=df.replace(value,value[0:-4])
                if '今日除息' in value:
                    df=df.replace(value,value[0:-4])
            df.index=pd.to_datetime(df['除息日'],yearfirst=True)
            stock.index=pd.to_datetime(stock['date'])
            today_year=datetime.today().strftime('%Y-%m-%d')[0:4]   ##今年
            year_=eval(today_year)-3
            max_=round(stock.loc[f'{year_}':]['Close'].max())
            min_=round(stock.loc[f'{year_}':]['Close'].min())
            avg=round(df[df['編號']==f'{stock_num_ent.get()}'].loc['2021']['現金殖利率'].sum(),2)
            datas.append([f'近3年最高價:{max_}'])
            datas.append([f'近3年最低價:{min_}'])
            datas.append([f'殖利率:{avg}%'])
            listVar2.set(datas)
        except Exception as e:
            print(e)
            label.config(text='請檢查資料庫(表)!!',fg='red')
        data=[]
        
        
    #######試算函式######
    def count():
        if conn==None:
            label.config(text='資料庫未開啟!',fg='red')
            return
        if stock_num_ent.get()=='':
            label.config(text='請輸入股號!!',fg='red')
            return
        if cost_result.get()=='':
            label.config(text='請輸入投資金額!!',fg='red')
            return
        if day_result.get()=='':
            label.config(text='請輸入每月幾號投資!!',fg='red')
            return
        if buyyear_result.get()=='':
            label.config(text='請輸入連續買進年數!!',fg='red')
            return
        try:
            with sqlite3.connect('stock.db') as con:             
                df=pd.read_sql_query('select * from "deposit"',con=con)         
                stock=pd.read_sql_query(f'select * from "{stock_num_ent.get()}"',con=con)
        except Exception as e:
            print(e)
            label.config(text='資料庫(表)讀取失敗!!',fg='red')
        try:
            for value in df['除息日'].values:
                if '即將除息' in value:
                    df=df.replace(value,value[0:-4])
                if '今日除息' in value:
                    df=df.replace(value,value[0:-4])
            df.index=pd.to_datetime(df['除息日'],yearfirst=True)
            stock.index=pd.to_datetime(stock['date'])
    
            today_year=datetime.today().strftime('%Y-%m-%d')[0:4]   ##今年
            today_month=int(datetime.today().strftime('%Y-%m-%d')[5:7])  ##今月
            count_years=int(today_year)-int(buyyear_result.get())
            if stock.loc[:f'{count_years}'].empty:
                label.config(text='超過該股發行日期，該股於'+stock.loc[f'{count_years}':]['date'][0]+'發行，請重新輸入買進年數!!',fg='red')
                return
            stock_gain=[]
            stock_pcs=[]
            stock_money=[]
            total=0
    
            for i in range(int(buyyear_result.get())):
                for k,j in enumerate(range(12)):
                    if today_month>=13:
                        today_month=1
                        count_years+=1
                    stock_price=stock.loc[f'{count_years+i}-{today_month}-{day_result.get()}':]['Close'][0].round(2)
                    total+=eval(cost_result.get())
                    pcs=total//stock_price
                    cost=stock_price*pcs
                    total-=cost
                    stock_gain.append(cost)
                    stock_pcs.append(pcs)
                    if f'{count_years+i}-{today_month}' in df[df['編號']==f'{stock_num_ent.get()}'].index:
                        money_deposit=eval(df[df['編號']==f'{stock_num_ent.get()}'].loc[f'{count_years+i}-{today_month}']['現金股利'][0])
                        get_deposit=sum(stock_pcs)*money_deposit
                        stock_money.append(get_deposit)
                    data.append([stock.loc[f'{count_years+i}-{today_month}-{day_result.get()}':]['date'][0],f'{stock_num_ent.get()}',df[df['編號']==f'{stock_num_ent.get()}']['股名'][0],f'{stock_price}X{pcs}',cost])
                    today_month+=1
                    if k==11 and today_month<=13:
                        count_years-=1
            gain=round((stock['Close'][-1]*sum(stock_pcs))-sum(stock_gain),2)     
            total_result.config(text=gain,font=('Arial',18),fg='blue')
            percent=round(gain/sum(stock_gain)*100,2)
            per.config(text=f'{percent}%',font=('Arial',18),fg='blue')
            money_result.config(text=round(sum(stock_money),1),font=('Arial',18),fg='blue')
            label.config(text='模擬試算完成!',fg='yellow')
        except Exception as e:
            print(e)
            label.config(text='請檢查資料庫(表)檔案!!',fg='red')

    #######開啟資料庫函式######
    def open_db():
        global conn
        global cursor
        conn=sqlite3.connect(file_name)
        try:
            cursor=conn.cursor()
            label.config(text=('聯結成功!'),fg='yellow')
        except Exception as e:
            print(e)
            label.config(text=('聯結失敗!'),fg='red')
    
    #######排行函式######
    def yield_rank():
        if conn==None:
            label.config(text='資料庫未開啟!',fg='red')
            return
        with sqlite3.connect('stock.db') as con:             
            df=pd.read_sql_query('select * from "deposit"',con=con)
        for value in df['除息日'].values:
            if '即將除息' in value:
                df=df.replace(value,value[0:-4])
            if '今日除息' in value:
                df=df.replace(value,value[0:-4])
        df.index=pd.to_datetime(df['除息日'],yearfirst=True)
        df=df.loc['2021']
        datas=[]
        # datas.append(['編號','股名','年均殖利率'])
        for value in set(df['編號']):
            avg=round(df[df['編號']==value]['現金殖利率'].sum(),2)
            datas.append([value,df[df['編號']==value]['股名'].values[0],avg])
        datas=sorted(datas,key=lambda x:x[2],reverse=True)
        df_avg=[['編號','股名','年均殖利率']]
        datas=df_avg+datas
        listVar3.set(datas[:10])
        label.config(text='搜尋完成!',fg='yellow')
    
    
    #######線圖函式######
    def stock_plot():
        if conn==None:
            label.config(text='資料庫未開啟!',fg='red')
            return
        if stock_num_ent.get()=='':
            label.config(text='請輸入股號!!',fg='red')
            return
        filename_pic='stock.png'
        path=r'./線圖資料/'
        if not os.path.exists(path):
            os.mkdir(path)
        with sqlite3.connect('stock.db') as con:             
            stock=pd.read_sql_query(f'select * from "{stock_num_ent.get()}"',con=con)
            df=pd.read_sql_query('select * from "deposit"',con=con)
        picture=tk.Toplevel()
        stock.index=pd.to_datetime(stock['date'])
        today_year=datetime.today().strftime('%Y-%m-%d')[0:4]   ##今年
        st_name=df[df['編號']==f'{stock_num_ent.get()}']['股名'].values[0]
        st=stock.loc[f'{today_year}':,['Close','Volume']]
        price_avg=[st['Close'].mean()]*len(st)
        picture.title('分析圖表') 
        
        plt.figure(figsize=(13,9))
        plt.suptitle(f'{stock_num_ent.get()}{st_name}',fontsize=(25))
        plt.subplot(2,1,1)
        plt.subplots_adjust(wspace=0.3,hspace=0.5)
        plt.grid(True)
        plt.xlabel('Date',labelpad=25,fontsize=(16))
        plt.ylabel('Price',rotation=0,labelpad=25,fontsize=(16))
        plt.plot(st.index,st['Close'])
        plt.plot(st.index,price_avg,'--')
        plt.annotate('Hight:'+str(round(st['Close'].max(),2)),xy=(st['Close'].idxmax(),st['Close'].max()),
                                  xytext=(st['Close'].idxmax(),st['Close'].max()),fontsize=20,arrowprops=dict(facecolor='blue',shrink=0.5,headwidth=5,headlength=40))
        plt.annotate('Low:'+str(round(st['Close'].min(),2)),xy=(st['Close'].idxmin(),st['Close'].min()),
                                  xytext=(st['Close'].idxmin(),st['Close'].min()),fontsize=20,arrowprops=dict(facecolor='red',shrink=0.5,headwidth=5,headlength=40))
        plt.text(st.index[0],st['Close'].mean()+0.1,'Avg:'+str(round(st['Close'].mean(),2)),fontsize=15)
        
        plt.subplot(2,1,2)
        plt.grid(True)
        plt.xlabel('Date',labelpad=25,fontsize=(16))
        plt.ylabel('Volume',rotation=0,labelpad=30,fontsize=(16))
        plt.bar(st.index,st['Volume'].div(10000))
        
        
        plt.savefig(f'{path}{filename_pic}')
        img_open=Image.open(f'{path}{filename_pic}')
        img_png=ImageTk.PhotoImage(img_open)
        label_img=tk.Label(picture,image=img_png,width=img_open.width,height=img_open.height)
        label_img.pack(anchor='center')

        
        
        label.config(text='輸出完成!',fg='yellow')
        picture.mainloop()
    
    
    #######匯出函式######
    def export():
        global data
        if conn==None:
            label.config(text='資料庫未開啟!',fg='red')
            return
        if stock_num_ent.get()=='':
            label.config(text='請輸入股號!!',fg='red')
            return
        if data==[]:
            label.config(text='請先使用Search以及試算!!',fg='red')
            return
        path=r'./存股資料/'
        if not os.path.exists(path):
            os.mkdir(path)
        inf=pd.DataFrame(data,columns=['日期','股號','股名','單價及股數','總價'])
        inf.to_csv(f'{path}{data[0][1]}.csv',encoding='utf-8-sig')
        label.config(text='匯出完成!',fg='yellow')
        data=[]
        
        
    #######結束函式######
    def end():
        if conn!=None:
            conn.close()
        win.destroy()
        
    
    font='華康華綜體W5' 
    #########Frame1#########    
    label=tk.Label(frame1,text='查詢狀況:',font=('華康正顏楷體W5',14),bg='black',fg='white'
                    ,anchor='w')
    label.pack(fill='x')
    
    #########Frame2#########
    #-------------------------------------#股號輸入
    stock_num_lab=tk.Label(frame2,text='輸入股號：',font=(font,18),fg='black')
    stock_num_ent=tk.Entry(frame2,font=('Arial',16),bg='#faf0e6',width=8,borderwidth=5)
    
    #-------------------------------------#Search按鈕
    search=tk.Button(frame2,text='Search',font=('Arial',16),fg='blue',borderwidth=5,command=stock_search)
    
    #-------------------------------------#股名
    stock_name=tk.Label(frame2,text='',font=('華康正顏楷體W5',24),fg='black')
    
    #-------------------------------------#股本
    equity_lab=tk.Label(frame2,text='股本：',font=(font,18),fg='black')
    equity_result=tk.Label(frame2,text='',font=('Arial',15),fg='black')
    
    #-------------------------------------#配息年數
    dividendyear_lab=tk.Label(frame2,text='配息年數：',font=(font,16),fg='black')
    dividendyear_result=tk.Label(frame2,text='',font=('Arial',15),fg='black')
    
    #-------------------------------------#昨日收盤價
    closeing_lab=tk.Label(frame2,text='昨日收盤價：',font=(font,16),fg='black')
    closeing_result=tk.Label(frame2,text='',font=('Arial',15),fg='black')
    
    #-------------------------------------#
    
    
    #--------------排版區域---------------#
    stock_num_lab.grid(row=0,column=0,pady=15)
    stock_num_ent.grid(row=0,column=1,pady=15)
    #-------------------------------------#
    search.grid(row=0,column=2,pady=15)
    #-------------------------------------#
    stock_name.grid(row=0,column=3,columnspan=3,padx=20,pady=15)
    #-------------------------------------#
    equity_lab.grid(row=1,column=0,pady=10)
    equity_result.grid(row=1,column=1,pady=10)
    #-------------------------------------#
    dividendyear_lab.grid(row=1,column=2,pady=10)
    dividendyear_result.grid(row=1,column=3,pady=10)
    #-------------------------------------#
    closeing_lab.grid(row=1,column=4,pady=10)
    closeing_result.grid(row=1,column=5,pady=10)
    #-------------------------------------#
    
    
    
    #########Frame3#########
    #-------------------------------------#董監持股
    shareholding_ratio=tk.Listbox(frame3,listvariable=listVar1,height=False,highlightthickness=10,font=('Arial',12),bg='dark red',fg='white')
    shareholding_ratio.pack(fill='x')
    
    #-------------------------------------#近3年最高及最低，殖利率
    high_to_down=tk.Listbox(frame3,listvariable=listVar2,height=False,highlightthickness=10,font=('Arial',12),bg='white',fg='black')
    high_to_down.pack(fill='x')
    
    #-------------------------------------#
    
    
    #########Frame4#########
    #-------------------------------------#投資金額輸入
    cost_lab=tk.Label(frame4,text='投資金額：',font=(font,18),fg='black')
    cost_result=tk.Entry(frame4,font=('Arial',16),bg='#faf0e6',width=6,borderwidth=5)
    
    #-------------------------------------#每月幾號輸入
    day_lab=tk.Label(frame4,text='每月幾號：',font=(font,18),fg='black')
    day_result=tk.Entry(frame4,font=('Arial',16),bg='#faf0e6',width=6,borderwidth=5)
    
    #-------------------------------------#買進年數輸入
    buyyear_lab=tk.Label(frame4,text='買進年數：',font=(font,18),fg='black')
    buyyear_result=tk.Entry(frame4,font=('Arial',16),bg='#faf0e6',width=6,borderwidth=5)
    
    #-------------------------------------#模擬試算按鈕
    calculate_butt=tk.Button(frame4,text='開始試算',font=(font,16),fg='blue',width=20,borderwidth=5,command=count)
    
    #-------------------------------------#損益顯示
    total_lab=tk.Label(frame4,text='總損益：',font=(font,18),fg='black')
    total_result=tk.Label(frame4,text='',font=('Arial',18),fg='black')
    
    #-------------------------------------#報酬率顯示
    per_lab=tk.Label(frame4,text='報酬率：',font=(font,18),fg='black')
    per=tk.Label(frame4,text='',font=('Arial',18),fg='black')
    
    #-------------------------------------#現金股利顯示
    money_lab=tk.Label(frame4,text='現金股利：',font=(font,18),fg='black')
    money_result=tk.Label(frame4,text='',font=('Arial',18),fg='black')
    
    #-------------------------------------#注意事項
    notice_lab=tk.Label(frame4,text='以上試算結果根據過去年數計算模擬，僅供參考!!',font=('標楷體',20,'underline'),fg='red')
    
    #-------------------------------------#
    
    
    #---------------排版區域---------------#
    cost_lab.grid(row=0,column=0,padx=5,pady=10)
    cost_result.grid(row=0,column=1,padx=5,pady=10)
    #-------------------------------------#
    day_lab.grid(row=0,column=2,padx=5,pady=10)
    day_result.grid(row=0,column=3,padx=5,pady=10)
    #-------------------------------------#
    buyyear_lab.grid(row=0,column=4,padx=5,pady=10)
    buyyear_result.grid(row=0,column=5,padx=5,pady=10)
    #-------------------------------------#
    calculate_butt.grid(row=1,column=2,columnspan=2,pady=10)
    #-------------------------------------#
    total_lab.grid(row=2,column=0,pady=10)
    total_result.grid(row=2,column=1,pady=10)
    #-------------------------------------#
    per_lab.grid(row=2,column=2,padx=5,pady=10)
    per.grid(row=2,column=3,padx=5,pady=10)
    #-------------------------------------#
    money_lab.grid(row=2,column=4,pady=10)
    money_result.grid(row=2,column=5,pady=10)
    #-------------------------------------#
    notice_lab.grid(row=3,column=0,columnspan=6,pady=5)
    #-------------------------------------#
    
    
    #########Frame5#########
    #-------------------------------------#按鈕
    button1=tk.Button(frame5,text='開啟',font=(font,18),fg='black',command=open_db)
    button1.grid(row=0,column=0,padx=10,pady=5)
    button2=tk.Button(frame5,text='年均殖利率排行',font=(font,18),fg='black',command=yield_rank)
    button2.grid(row=0,column=1,padx=10,pady=5)
    button3=tk.Button(frame5,text='該股價格線圖(年)',font=(font,18),fg='black',command=stock_plot)
    button3.grid(row=0,column=2,padx=10,pady=5)    
    button4=tk.Button(frame5,text='匯出',font=(font,18),fg='black',command=export)
    button4.grid(row=0,column=3,padx=10,pady=5)
    button5=tk.Button(frame5,text='離開',font=(font,18),fg='black',command=end)
    button5.grid(row=0,column=4,padx=10,pady=5)
    
    
    #########Frame6#########
    #-------------------------------------#年均殖利率排行
    yield_ranking=tk.Listbox(frame6,listvariable=listVar3,font=('Arial',12),bg='dark red',fg='white')
    yield_ranking.pack(fill='x')
    
    
    #########Frame7#########
    #-------------------------------------#資料來源聲明
    source_lab=tk.Label(frame7,text='開發日期:2021/11/19  資料來源:Yahoo股市、GoodInfo台灣股市資訊網',font=(font,14),bg='black',fg='yellow'
                    ,anchor='e')
    source_lab.pack(fill='x')
    
    #-------------------------------------#    
        
        
        
    
    
    #-------------------------------------#爬蟲函式
    def get_chrom(url,driver=r'C:\webdriver\chromedriver'):
        try:
            options=webdriver.ChromeOptions()
            options.add_argument('--headless')
            chrom=webdriver.Chrome(driver,options=options)
            chrom.implicitly_wait(1)
            chrom.get(url)
            return chrom
        except Exception as e:
            return None    
    
    
data=[]      
conn=None
cursor=None
file_name='stock.db'
listVar1=tk.StringVar() ##董監持股
listVar2=tk.StringVar() ##最高及最低
listVar3=tk.StringVar() ##殖利率排行
listVar1.set([])
listVar2.set([])
listVar3.set([])


showmain()
win.mainloop()