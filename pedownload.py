# -*- coding: utf-8 -*-
"""
Created on Wed Apr  6 19:59:09 2016

@author: xjc
"""

import zipfile
import xlrd
import urllib.request
import sqlite3
import datetime
import os


def getZipfile(date):
    global COUNT,DATE
    url=r"http://115.29.204.48/syl/bk"+date+".zip"
    try:
        temfile=urllib.request.urlretrieve(url)[0]
        COUNT=0
        DATE=date
        return temfile 
    except:
        COUNT=COUNT+1
        predate=datetime.datetime.strptime(date, "%Y%m%d").date()-datetime.timedelta(days = 1)
        DATE=predate.strftime('%Y%m%d')
        return getZipfile(DATE)   #递归用return
    
def zip2xls(zipf):
    z=zipfile.ZipFile(zipf, 'r') # 读取压缩包内的文件。z.namelist()是文件名列表
    z.extract(z.namelist()[0])
    data = xlrd.open_workbook(z.namelist()[0],encoding_override="utf-8")
    os.remove(z.namelist()[0])  #删除解压的excel文件
    return data

def insertData(date): 
    global DATE
    zipf=getZipfile(date)
    #zipf=getzip[0]   #注意getZipfile函数返回两个值，第一个是zipfile，第二个是date
    data=zip2xls(zipf)
    lyrpe=data.sheet_by_name("板块静态市盈率").col_values(1)
    ttmpe=data.sheet_by_name("板块滚动市盈率").col_values(1)
    pb=data.sheet_by_name("板块市净率").col_values(1)
    dyr=data.sheet_by_name("板块股息率").col_values(1)
    
    db=sqlite3.connect("pe.db")
    tablename=["huA","shenA","hsA","shenZhu","zhongxiao","cyb"]
    ln=len(tablename)
    for i in range(ln):
        sql="create table if not exists "+tablename[i]+" (date text primary key,\
                                    lyrpe real,\
                                    ttmpe real,\
                                    pb real,\
                                    dyr real);"
        db.cursor().executescript(sql)
        sql='insert or replace into '+tablename[i]+' (date, lyrpe,ttmpe,pb,dyr) values (?,?,?,?,?)'
        j=i+1
        db.execute(sql,(DATE,float(lyrpe[j]),float(ttmpe[j]),float(pb[j]),float(dyr[j])))
        print(DATE," ",COUNT)
    #db.execute('insert into huA (date, lyrpe,ttmpe,pb,dyr) values ("20160405",15.32,45.99,1.66,57.1)')              
    db.commit()
    db.close()

def isInTable(date):
    cur = db.execute('select title, text from entries order by id desc')

if __name__=="__main__":
    #date=datetime.date.today().strftime('%Y%m%d')
    COUNT=0
    DATE=""      #记录修改以后的date
    date=datetime.date.today().strftime('%Y%m%d')
    #date="20120507"
    while COUNT<10:
        insertData(date)
        predate=datetime.datetime.strptime(DATE, "%Y%m%d").date()-datetime.timedelta(days = 1)
        date=predate.strftime('%Y%m%d')
        
        
