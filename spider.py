import requests
import json
import xlrd
import xlwt
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import math
import pandas as pd
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path) #打开此地址下的exl文档
    sheet_name = data_xsls.sheets()[0]  #进入第一张表
    count_nrows = sheet_name.nrows  #获取总行数
    #print(count_nrows)
    count_nocls = sheet_name.ncols  #获得总列数
    for i in range(0,31):
        if type(sheet_name.cell_value(i,3)) != type(0):
            city.append(sheet_name.cell_value(i,6))  
        else:
            pass
    for i in range(0,count_nrows):
        if type(sheet_name.cell_value(i,3)) != type(0):
            scenic.append(sheet_name.cell_value(i,3))  
        else:
            pass
    for i in range(0,31):
        if type(sheet_name.cell_value(i,3)) != type(0):
            cityname.append(sheet_name.cell_value(i,4))  
        else:
            pass
    for i in range(0,count_nrows):
        if type(sheet_name.cell_value(i,3)) != type(0):
            scenicname.append(sheet_name.cell_value(i,1))  
        else:
            pass
    for i in range(0,count_nrows):
        if type(sheet_name.cell_value(i,2)) != type(0):
            scenicmingzi.append(sheet_name.cell_value(i,2))  
        else:
            pass
    for i in range(0,31):
        if type(sheet_name.cell_value(i,5)) != type(0):
            citymingzi.append(sheet_name.cell_value(i,5))  
        else:
            pass
    for i in range(0,31):
        shiname.append([])
        if i>0:
            shitime[i] = shitime[i]+shitime[i-1]
        for j in range(7,len(sheet_name.row_values(i))):
            if sheet_name.cell_value(i,j) != '':
                shiname[i].append(sheet_name.cell_value(i,j)) 
                shitime[i]+=1
            else:
                pass

data_path = '22.xlsx'
city = []
cityname = []
citymingzi = []
scenic = []
scenicname = []
scenicmingzi = []
shiname = []
shitime = np.zeros(31)
read_xsls(data_path)
#print(shitime)
#for m in range(22,len(city)):
path='E:\\tour\\20181118.xls'
rb = xlwt.Workbook()
sheet = rb.add_sheet('sheet1') #新建sheet
sheet.write(0,0,'游客来源地')
for i in range(0,54):
    sheet.write(0,i+2,scenicmingzi[i])
for i in range(0,31):
    for j in range(0,len(shiname[i])):
        if i>0:
            sheet.write(j+int(shitime[i-1])+1,0,citymingzi[i])
            sheet.write(j+int(shitime[i-1])+1,1,shiname[i][j])
        else:
            sheet.write(j+1,0,citymingzi[i])
            sheet.write(j+1,1,shiname[i][j])
for m in range(0,len(city)):
    for n in range(0,len(scenic)):
        headers = {
        # 假装自己是浏览器
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Ubuntu Chromium/73.0.3683.75 Chrome/73.0.3683.75 Safari/537.36',
        # 把你刚刚拿到的Cookie塞进来
        'Cookie': 'JSESSIONID=eaa392aa-fa75-4ad0-bad5-88ed2dc41fc2',
        }
        session = requests.Session()
        response = session.get('http://218.94.79.6:8095/nanjingTourism/touristsource/queryOutPc/20181118/20181124/'+ str(city[m])+ '/'+ str(scenic[n]), headers=headers)

        a = response.text
        user_dict = json.loads(a)
        #print(user_dict['data'][0]['name'])
        #path='E:\\tour\\'+ str(scenicname[n]) +'1-' +  str(cityname[m]) + '.xls'
        if m>0:
            for i in range(0,len(user_dict['data'])):
                for j in range(0,len(shiname[m])):
                    if user_dict['data'][i]['name'] == shiname[m][j]:
                        sheet.write(j+1+shitime[m-1],n+2,user_dict['data'][i]['value'])
        else:
            for i in range(0,len(user_dict['data'])):
                for j in range(0,len(shiname[m])):
                    if user_dict['data'][i]['name'] == shiname[m][j]:
                        sheet.write(j+1,n+2,user_dict['data'][i]['value'])
        print(scenic[n])
    print(city[m],'输入完成')
rb.save(path)

data_path = '22.xlsx'
city = []
cityname = []
citymingzi = []
scenic = []
scenicname = []
scenicmingzi = []
shiname = []
shitime = np.zeros(31)
read_xsls(data_path)
#print(shitime)
#for m in range(22,len(city)):
path='E:\\tour\\20181125.xls'
rb = xlwt.Workbook()
sheet = rb.add_sheet('sheet1') #新建sheet
sheet.write(0,0,'游客来源地')
for i in range(0,54):
    sheet.write(0,i+2,scenicmingzi[i])
for i in range(0,31):
    for j in range(0,len(shiname[i])):
        if i>0:
            sheet.write(j+int(shitime[i-1])+1,0,citymingzi[i])
            sheet.write(j+int(shitime[i-1])+1,1,shiname[i][j])
        else:
            sheet.write(j+1,0,citymingzi[i])
            sheet.write(j+1,1,shiname[i][j])
for m in range(0,len(city)):
    for n in range(0,len(scenic)):
        headers = {
        # 假装自己是浏览器
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Ubuntu Chromium/73.0.3683.75 Chrome/73.0.3683.75 Safari/537.36',
        # 把你刚刚拿到的Cookie塞进来
        'Cookie': 'JSESSIONID=eaa392aa-fa75-4ad0-bad5-88ed2dc41fc2',
        }
        session = requests.Session()
        response = session.get('http://218.94.79.6:8095/nanjingTourism/touristsource/queryOutPc/20181125/20181201/'+ str(city[m])+ '/'+ str(scenic[n]), headers=headers)

        a = response.text
        user_dict = json.loads(a)
        #print(user_dict['data'][0]['name'])
        #path='E:\\tour\\'+ str(scenicname[n]) +'1-' +  str(cityname[m]) + '.xls'
        if m>0:
            for i in range(0,len(user_dict['data'])):
                for j in range(0,len(shiname[m])):
                    if user_dict['data'][i]['name'] == shiname[m][j]:
                        sheet.write(j+1+shitime[m-1],n+2,user_dict['data'][i]['value'])
        else:
            for i in range(0,len(user_dict['data'])):
                for j in range(0,len(shiname[m])):
                    if user_dict['data'][i]['name'] == shiname[m][j]:
                        sheet.write(j+1,n+2,user_dict['data'][i]['value'])
        print(scenic[n])
    print(city[m],'输入完成')
rb.save(path)


'''sheet.write(0,0,'游客来源')
sheet.write(0,1,'游客人数')
sheet.write(0,2,'占比')
for i in range(0,len(user_dict['data'])):
    sheet.write(i+1,0,user_dict['data'][i]['name'])
    sheet.write(i+1,1,user_dict['data'][i]['value'])
    sheet.write(i+1,2,user_dict['data'][i]['percent'])
rb.save(path)
#print(type(response.text))'''