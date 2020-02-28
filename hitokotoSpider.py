#!/usr/bin/python
# -*- coding: UTF-8 -*-
'''
HitokotoSpider V.1.5  Update in 2020-02-28 15:36
请提前安装所需的库:
pip install requests xlrd xlwt xlutils
请注意：本程序创建的表格为 Excel 2003 即XLS文件
        向表格中写入的数据不可超过65536行256列。
    如果您使用的是Python3.8以下的版本
        出现“module 'time' has no attribute 'perf_counter'”的问题时
        请将 time.perf_counter() 替换为 time.clock()  如果没有出现此问题，请反馈。
    如果您使用的是Python3以下的版本，重复率计算可能会有所误差，但不影响使用。
'''
import requests
import time,random
import xlrd,xlwt
from xlutils.copy import copy
print("引入所需库完成。")
print()
SLEEP_TIME = 0 #每次间隔秒数
path = input("保存的表格名字（无需后缀）:")
path = path + ".xls"#保存表格名字


try:
     f = open(path)
     f.close()
     print('%s文件存在，将对表格进行不覆盖追加读写。'%path)
     print()
except IOError:        
     workbook = xlwt.Workbook(encoding='utf-8')       #新建工作簿
     sheet1 = workbook.add_sheet("Hitokoto")          #新建sheet
     workbook.save(path)   #保存
     print('%s文件不存在，已创建。'%path)
     for k in range(0, 10):
          ClassHeader = ["id","hitokoto","type","from_a","from_who","creator","creator_uid","reviewer","uuid","created_at"]
          workbook = xlrd.open_workbook(path)  # 打开工作簿
          sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
          worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
          rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
          new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
          new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
          new_worksheet.write(0, k, ClassHeader[k])
          new_workbook.save(path)
     print()
     
COUNTS = input("要获取不重复一言的条数：") #需要采集的条数获取
COUNTS = int(COUNTS, base=10)   #将输入的条数化为整型
Need   = COUNTS
All    = 0
headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:73.0) Gecko/20100101 Firefox/73.0'}
print("=======准备开始抓取========")
print()
time.sleep(1)
Start_Time_Second = time.perf_counter()#记录程序开始运行时间(用于计算重复率)
def get_res(count,ids):
    global All#必须声明全局变量
    global Start_Time_Second
    url = 'https://v1.hitokoto.cn/'
    res = requests.get(url=url,headers=headers,)
    All = All + 1
    End_Time_Second = time.perf_counter()#记录程序结束运行时间(用于计算重复率)
    Spend_time = End_Time_Second-Start_Time_Second
    #展示花费的时间并转化为适当的单位
    if Spend_time > 60:
        Show_Spend_Time = str(round(Spend_time/60,2))+"min"
    else:
        Show_Spend_Time = str(round(Spend_time,2))+"s"
    #变量Spend_time 和 Speed 不要混淆！
    Speed = round(All/Spend_time,1)#每秒请求的条数
    Repeat = round(count/All*100,2)#计算重复率
   #print('目前正采集第%s条'%count,str('/%s条'%Need),'共获取%s条'%All,'不重复率%s%% '%Repeat,res.status_code)
    print("▢目前正采集第{}条/{}条 共获取{}条 不重复率{}% 用时{} QPS={} {}".format(count,Need,All,Repeat,Show_Spend_Time,Speed,res.status_code))
    if res.status_code != 200:
        print('返回状态码不为200，可能被限制，暂停30S')
        time.sleep(30)
    ret = res.json() # 将获取到的结果转为json字符串
    id = ret['id']
    if id in ids: #去重
        print('× 丢弃了一个重复句子。 ×')
        get_res(count,ids)
    else:
        ids.add(id)
    hitokoto = ret['hitokoto']
    type = ret['type']
    from_a = ret['from']
    from_who = ret['from_who']
    creator = ret['creator']
    creator_uid = ret['creator_uid']
    reviewer = ret['reviewer']
    uuid = ret['uuid']
    created_at = ret['created_at']
    con = [id,hitokoto,type,from_a,from_who,creator,creator_uid,reviewer,uuid,created_at]
    HitokotoWord = [hitokoto]
    print(hitokoto,"——",from_a)
    return con


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print()
   #print("=======储存数据成功========")
    print("===========================")
    print()

if __name__ == '__main__':
    ids = set() #去重id
    cons = [] #保存已经获取的内容
    for count in range(1,COUNTS+1):
        con = get_res(count,ids)
        cons.append(con)
        if len(cons)%1 == 0: #每1个回话保存一次
            write_excel_xls_append(path,cons)
            cons = []
           #time.sleep(random.randint(1, SLEEP_TIME))  # 随机间隔(SLEEP_TIME为整数)
            time.sleep(SLEEP_TIME)  # 间隔(SLEEP_TIME可不为整数)
        if count == COUNTS :
            write_excel_xls_append(path,cons)

