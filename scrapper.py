
import requests
from bs4 import BeautifulSoup
import xlsxwriter
#一區38球取6 二區8球取1
ids = []
dates = []
first_sections = []
second_sections = []
headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36'}
#這次要爬的網址 https://www.pilio.idv.tw/lto/listbbk.asp?indexpage=1&orderby=new
for i in range(1,53):#用for 迴圈抓52頁的數據
    url = f"https://www.pilio.idv.tw/lto/listbbk.asp?indexpage={i}&orderby=new"
    html = requests.get(url,headers=headers)
    html.encoding ='big5'
    html.raise_for_status()
    soup = BeautifulSoup(html.text, "html.parser")#抓出所有網頁碼
    code_datas1 = soup.select('b')[5:]#發現期數 日期 球號 都是<b> 抓出<b> 每頁前6項是我們不需要的
    item = 0
    #5次一循環 將資料抓到屬於自己的組別內
    for code_data1 in code_datas1:
        if item%5 == 0:#存期數
            ids.append(code_data1.getText())
        elif item%5 == 1:#存日期
            dates.append(code_data1.getText())
        elif item%5 == 2:#存第一區
            #first.append(code_data1.getText()) 抓出亂碼
            text = str(code_data1.getText()).replace(u"\xa0","")
            first_sections.append(text) 
        elif item%5 == 3:#存第二區
            second_sections.append(code_data1.getText())
        item+=1


print('期數',ids[0])
print('日期',dates[0])
print('第一區',first_sections[0])
#print('第二區',len(second_sections))

book = xlsxwriter.Workbook(".\data.xlsx")#建立excel檔
#--------------------------------------------------------------------
all_data = book.add_worksheet("all_data")#增加分頁用來放所有資料)
titles = ['期數','日期','第一區-1','第一區-2','第一區-3','第一區-4','第一區-5','第一區-6','第二區']
#儲存格樣式
title_format = book.add_format({'bold': True,'bg_color': '#262626','font_color': 'FFFFFF','border': True})
first_format = book.add_format({'bold': True,'font_color':'008000','top': True,'bottom': True})
second_format = book.add_format({'bold': True,'font_color': 'FF0000','border': True,'left': 3})
others_format = book.add_format({'bold': True,'bg_color': "#D0D3D9",'border': True})
all_data.set_column("B:B",12)
all_data.set_column("C:H",10)
#寫入標題
for col in range(0,9):
    all_data.write(0,col,titles[col],title_format)
#記錄行數
row = 1
#寫入資料
for data in range(0,len(ids)):
    all_data.write(row,0,ids[data],others_format)#寫入期數
    all_data.write(row,1,dates[data],others_format)#寫入日期
    for i in range(0,6):#寫入第一區
        all_data.write(row,i+2,first_sections[data].split(',')[i],first_format)
    all_data.write(row,8,second_sections[data],second_format)#寫入第二區
    row+=1
#-------------------------------------------------------------------
#增加新分頁放球出現次數
frequence = book.add_worksheet("frequency_of_number")
frequen_title = ['第一區球號','出現球數','第二區球數']
fre_fir_format = book.add_format({'bold' : True,'font_color' : '008000','bg_color': 'D0D3D9','border' : True})
fre_sec_format = book.add_format({'bold' : True,'font_color' : 'FF0000','bg_color': 'D0D3D9','border' : True})
fre_format = book.add_format({'bold' : True,'border' : True,'align':'left'})
#寫入標題
for i in range(0,6):
    if i % 2 == 0:
        frequence.write(0,i,frequen_title[0],title_format)
    else:
        frequence.write(0,i,frequen_title[1],title_format)
frequence.write(0,6,frequen_title[2],title_format)
frequence.write(0,7,frequen_title[1],title_format)
frequence.set_column("A:O",13)
#寫球數
col_index = 0
row_index = 1
#第一區
first_frequency = {}#記錄第一區球出現次數的字典
for ball in range(1,39):
    appearance = 0
    number = str(ball)
    if ball < 10:
        number = '0'+str(ball)
    if ball ==14:
        col_index = 2
        row_index = 1
    elif ball ==27:
        col_index = 4
        row_index = 1
    frequence.write(row_index,col_index,number,fre_fir_format)
    for i in range(0,len(first_sections)):
        appearance+= first_sections[i].count(f'{number}')
    frequence.write(row_index,col_index+1,int(appearance) ,fre_format)
    first_frequency[f'{number}'] = appearance
    row_index+=1

   
#第二區
col_index = 6
row_index = 1
second_frequency = {}#記錄第二區球出現次數的字典
for ball in range(1,9):
    appearance = 0
    number = '0'+str(ball)
    frequence.write(row_index,col_index,number,fre_sec_format)
    for i in range(0,len(second_sections)):
        appearance+= second_sections[i].count(f'{number}')
    frequence.write(row_index,col_index+1,appearance,fre_format)
    second_frequency[f'{number}']= appearance
    row_index+=1

#繪圖
#第一區分三張圖表較清楚
chartfirstone = book.add_chart({'type':'column'})
chartfirstone.set_x_axis({'name':'number'})
chartfirstone.set_y_axis({'min': 150, 'max': 300})
chartfirstone.set_title({'name':'Frequency Of Numbers\n(First Sections)'})
chartfirstone.add_series({'name':'Frequency','fill':   {'color': 'green'},'categories':'=frequency_of_number!$A$2:$A$14','values':'=frequency_of_number!$B$2:$B$14'})
frequence.insert_chart('A16',chartfirstone)

chartfirstwo = book.add_chart({'type':'column'})
chartfirstwo.set_x_axis({'name':'number'})
chartfirstwo.set_y_axis({'min': 150, 'max': 300})
chartfirstwo.set_title({'name':'Frequency Of Numbers\n(First Sections)'})
chartfirstwo.add_series({'name':'Frequency','fill':   {'color': 'green'},'categories':'=frequency_of_number!$C$2:$C$14','values':'=frequency_of_number!$D$2:$D$14&$F$2:$F$13'})
frequence.insert_chart('F16',chartfirstwo)

chartfirsthree = book.add_chart({'type':'column'})
chartfirsthree.set_x_axis({'name':'number'})
chartfirsthree.set_y_axis({'min': 150, 'max': 300})
chartfirsthree.set_title({'name':'Frequency Of Numbers\n(First Sections)'})
chartfirsthree.add_series({'name':'Frequency','fill':   {'color': 'green'},'categories':'=frequency_of_number!$E$2:$E$13','values':'=frequency_of_number!$F$2:$F$13'})
frequence.insert_chart('K16',chartfirsthree)
#第二區
chartsecond = book.add_chart({'type':'column'})
chartsecond.set_x_axis({'name':'number'})
chartsecond.set_title({'name':'Frequency Of Numbers\n(Second Sections)'})
chartsecond.add_series({'name':'Frequency','fill':   {'color': 'red'},'categories':'=frequency_of_number!$G$2:$G$9','values':'=frequency_of_number!$H$2:$H$9'})
frequence.insert_chart('K1',chartsecond)
#sum up

#寫出一區最多與最少出現的六個球數
#最多
frequence.set_column("P:P",21)
frequence.set_column("Q:Q",10)
sumuptitle = ['第二區最常出現號碼','第一區最常出現號碼','第一區最少出現號碼','出現次數']

frequence.write(0,15,sumuptitle[0],title_format)
frequence.write(9,15,sumuptitle[1],title_format)
frequence.write(16,15,sumuptitle[2],title_format)
frequence.write(0,16,sumuptitle[3],title_format)
frequence.write(9,16,sumuptitle[3],title_format)
frequence.write(16,16,sumuptitle[3],title_format)

first_less = sorted(first_frequency.items(),key=lambda x:x[1])
first_most = sorted(first_frequency.items(),key=lambda x:x[1],reverse=True)
for i in range(0,6):#第一區最多球數
    frequence.write(i+10,15,str(first_most[i][0]),fre_fir_format)
    frequence.write(i+10,16,str(first_most[i][1]),fre_format)

for i in range(0,6):#第一區最少球數
    frequence.write(i+17,15,str(first_less[i][0]),fre_fir_format)
    frequence.write(i+17,16,str(first_less[i][1]),fre_format)

second_most = sorted(second_frequency.items(),key=lambda x:x[1],reverse=True)
for i in range(0,8):#第二區最多球數
    frequence.write(i+1,15,str(second_most[i][0]),fre_sec_format)
    frequence.write(i+1,16,str(second_most[i][1]),fre_format)
#---------------------------------------------------------------------
#最久沒出現的球號
#寫標題
frequence.set_column("S:Z",16)
for i in range(0,6):
    if i % 2 == 0:
        frequence.write(0,i+18,frequen_title[0],title_format)
    else:
        frequence.write(0,i+18,'隔幾期沒出現',title_format)
frequence.write(0,24,frequen_title[2],title_format)
frequence.write(0,25,'隔幾期沒出現',title_format)
#記錄第一區的球有多久沒出現
first_incident = {}
for ball in range(1,39):
    number = str(ball)
    if ball < 10:
        number = '0'+str(ball)
    for id in range(0,len(first_sections)):
        if number in first_sections[id]:
            first_incident[number] = id
            break
first_incident = sorted(first_incident.items(),key=lambda x:x[1],reverse=True)
col_index = 18
row_index = 1
#寫入第一區的球有多久沒出現
for ball in range(1,39):
    if ball ==14:
        col_index = 20
        row_index = 1
    elif ball ==27:
        col_index = 22
        row_index = 1
    frequence.write(row_index,col_index,first_incident[ball-1][0],fre_fir_format)
    frequence.write(row_index,col_index+1,first_incident[ball-1][1],fre_format)
    row_index+=1
#記錄第二區的球有多久沒出現
second_incident = {}
for ball in range(1,9):
    number = '0'+str(ball)
    for id in range(0,len(second_sections)):
        if number in second_sections[id]:
            second_incident[number] = id
            break
second_incident = sorted(second_incident.items(),key=lambda x:x[1],reverse=True)
col_index = 24
row_index = 1
for ball in range(1,9):
    frequence.write(row_index,col_index,second_incident[ball-1][0],fre_sec_format)
    frequence.write(row_index,col_index+1,second_incident[ball-1][1],fre_format)
    row_index+=1

book.close()#保存excel