# -*- coding: utf-8 -*-
'''
  It IS ONLY USED FOR RECORD OR STUDY.
'''


import xlrd
import xlwt
from datetime import datetime,date
import time
import os
 
a,b,c ='9:30:00','18:00:00','9:00:00'

def read_excel():
  
  # 打开原始文件
  workbook = xlrd.open_workbook('E:\PythonWork\CD.xls')
  # 获取特定的sheet
  #print (workbook.sheet_names()) # [u'sheet1', u'sheetsource']
  #sheetsource_name = workbook.sheet_names()[1]
  
  # 根据sheet索引或者名称获取sheet内容,sheet索引从0开始,而我需要处理的sheet是第一个
  sheetsource = workbook.sheet_by_index(0) 
  
 
  # sheet的名称，行数，列数
  print (sheetsource.name,sheetsource.nrows,sheetsource.ncols)
 
 
  #打开一个workbook用于保存sheet信息
  wb_result = xlwt.Workbook()
  
  #用于保存目标sheet的行索引
  target_rowi = 0
  
  #设置保存sheet的名字，这里只写了一个sheet，如果必要可以写sheet2，sheet3...
  sheet1_result=wb_result.add_sheet('worktime_result',cell_overwrite_ok=True)
  for sheetsource_i in range(sheetsource.ncols):
    sheet1_result.write(0,sheetsource_i,sheetsource.row_values(0)[sheetsource_i])  
  
  
  #增加列，“工作时间”“原因”
  sheet1_result.write(target_rowi,sheetsource.ncols,"工作时间") 
  sheet1_result.write(target_rowi,sheetsource.ncols+1,"原因") 
  
  
 # 获取不正常的刷卡记录，这里为了简便，没做识别处理，假设六，八列分别是上班刷卡起始和结束时间
  for source_rowi in range(1,sheetsource.nrows):
    #结束时间没有刷卡记录
    if sheetsource.row(source_rowi)[7].value == '':
      target_rowi=target_rowi+1
      for sheetsource_i in range(sheetsource.ncols):
        sheet1_result.write(target_rowi,sheetsource_i,sheetsource.row_values(source_rowi)[sheetsource_i])
        sheet1_result.write(target_rowi,sheetsource.ncols+1,'无正常刷卡')
    else: 
      irows_begin = sheetsource.row(source_rowi)[5].value
      d1=datetime.strptime(irows_begin,'%H:%M:%S')
      irows_end = sheetsource.row(source_rowi)[7].value
      d2=datetime.strptime(irows_end,'%H:%M:%S')
    
      work_time = datetime.strptime(c,'%H:%M:%S')
      work_time = datetime.strptime(a,'%H:%M:%S')
      end_time = datetime.strptime(b,'%H:%M:%S')
      #print (d1,d2,d2-d1)
    
      if d2-d1 < end_time-work_time :
        #print ("~~~~")
        target_rowi=target_rowi+1
        for sheetsource_i in range(sheetsource.ncols):
          sheet1_result.write(target_rowi,sheetsource_i,sheetsource.row_values(source_rowi)[sheetsource_i])
        sheet1_result.write(target_rowi,sheetsource.ncols,str(d2-d1))
        sheet1_result.write(target_rowi,sheetsource.ncols+1,'小于8小时')
      if d1 > work_time and d2-d1> end_time-work_time:
        #print("超过9：30")
        #print (d1,work_time)
        target_rowi=target_rowi+1
        for sheetsource_i in range(sheetsource.ncols):
          sheet1_result.write(target_rowi,sheetsource_i,sheetsource.row_values(source_rowi)[sheetsource_i])
        sheet1_result.write(target_rowi,sheetsource.ncols,str(d2-d1))
        sheet1_result.write(target_rowi,sheetsource.ncols+1,'超过9：30')
      elif d2 < end_time and d2-d1> end_time-work_time:
        #print("不到18：00")
        target_rowi=target_rowi+1
        for sheetsource_i in range(sheetsource.ncols):
          sheet1_result.write(target_rowi,sheetsource_i,sheetsource.row_values(source_rowi)[sheetsource_i])
        sheet1_result.write(target_rowi,sheetsource.ncols,str(d2-d1))
        sheet1_result.write(target_rowi,sheetsource.ncols+1,'提前下班')
  
  #接下来遍历没有刷卡记录的情况
  #从表单中得到人员名单，当然如果整个月都没有刷卡的人不在此显示
  all_list = sheetsource.col_values(2)
  name_list = list(set(all_list))
  
  #删除无用元素 如titile
  name_list.remove('姓名')
  
  #得到工作日期
  all_date = sheetsource.col_values(4)
  month_date =list(set(all_date))
  #删除多余元素
  month_date.remove('日期')

  work_date = []
  for date_i in month_date:
    day=datetime.strptime(date_i,'%Y-%m-%d')
    if day.weekday() in range(5):
     #work_date 记录一个month中工作的日期
      work_date.append(date_i)
  
  #修改特殊的工作日，调休等，~~~根据情况处理~~
  work_date.remove('2016-10-5')
  work_date.append('2016-10-8')
  work_date.append('2016-10-9')
  
  #对每个人名，循环遍历表单
  for name_i in name_list:
    tmp_list=[]
    for line_i in range(1,sheetsource.nrows):
      if sheetsource.row_values(line_i)[2] == name_i:
        #tmp_value = datetime.strptime(sheetsource.row(line_i)[4].value,'%Y-%m-%d')
        tmp_value = sheetsource.row(line_i)[4].value
        if str(tmp_value) in work_date:
         
          #添加此人本月内的工作日
          tmp_list.append(sheetsource.row_values(line_i)[4])
  
    except_date = list(set(work_date).difference(set(tmp_list)))
    
    #记录没有刷卡的情况，写到result中
    for except_i in except_date:
      target_rowi = target_rowi +1
      sheet1_result.write(target_rowi,2,name_i)
      sheet1_result.write(target_rowi,4,except_i)
      sheet1_result.write(target_rowi,sheetsource.ncols+1,'无正常刷卡')
 
    #print(except_date)
		
		
  #可以去你的工作目录下查看结果了
  wb_result.save('result.xls')   

 


  
  #  另外，获取单元格内容的数据类型
  #print (sheetsource.cell(1,5).ctype)
 
if __name__ == '__main__':
  read_excel()