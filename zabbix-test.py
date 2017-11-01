'''
Created on 2017年6月20日

@author: liuyanjie
'''
#!/usr/bin/python
#coding:utf-8
 
#import MySQLdb
import pymysql
# from aiohttp.hdrs import EXPECT

import time
import datetime
import xlsxwriter

ts_first = int(time.mktime(datetime.date(datetime.date.today().year,datetime.date.today().month-1,1).timetuple()))
lst_last = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
ts_last = int(time.mktime(lst_last.timetuple()))
# print("ts-first is :",ts_first)
# print("ts-last is :",ts_last)

print('ts_first time is : ',ts_first,', localtime is : ',time.asctime(time.localtime(ts_first)))

print('ts_last time is : ',ts_last,', localtime is : ',time.asctime(time.localtime(ts_last)))

win_templateid=10081
linux_templateid=10001
#zabbix数据库信息：
# zdbhost = '1.1.1.1'
# zdbuser = 'root'
# # zdbpass = '123456'
# zdbport = 3306
# zdbname = 'zabbix'

zdbhost = '1.1.1.1'
zdbuser = 'root'
zdbpass = '123456'
zdbport = 3306
zdbname = 'zabbix'

 
#生成文件名称：
# xlsfilename = 'damo.xls'
 
#需要查询的key列表 [名称，表名，key值，取值，格式化，数据整除处理]
keys_linux = [

    ['已用内存大小','trends_uint','vm.memory.size[used]','avg','',1],
    ['物理内存大小(单位G)','trends_uint','vm.memory.size[total]','avg','',1048576000],
#     ['可用平均内存(单位G)','trends_uint','vm.memory.size[available]','avg','',1048576000],
    ['内存占用率','trends','vm.memory.size[pused]','avg','',1],
    ['CPU5分钟负载','trends','system.cpu.load[percpu,avg5]','avg','%.2f',1],
    ['根分区使用率','trends','vfs.fs.size[/,pused]','avg','',1],
    ['home分区使用率','trends','vfs.fs.size[/boot,pused]','avg','',1],
#     ['根分区总大小(单位G)','trends_uint','vfs.fs.size[/,total]','avg','',1073741824],
#     ['根分区平均剩余(单位G)','trends_uint','vfs.fs.size[/,free]','avg','',1073741824],
]

keys_win = [
#     ['CPU利用率','trends','perf_counter[\\Processor(_Total)\\% Processor Time]','max','%.2f',1],
    ['可用平均内存(单位G)','trends_uint','vm.memory.size[free]','avg','',1048576000],
    ['物理内存大小(单位G)','trends_uint','vm.memory.size[total]','avg','',1048576000],
    ['内存使用率','trends','vm.memory.size[pused]','min','%.2f',1],
    ['CPU使用率','trends',r'perf_counter[\\Processor(_Total)\\% Processor Time]','max','',1],
    ['C盘空间使用率','trends','vfs.fs.size[C:,pused]','avg','%.2f',1],
    ['D盘空间使用率','trends','vfs.fs.size[D:,pused]','avg','%.2f',1],
]
 
 # 连接数据库
mydb = pymysql.connect(host=zdbhost,user=zdbuser,passwd=zdbpass,db=zdbname)
mycur =mydb.cursor()
groupname="ZZ-SERVERS"

# 根据主机ID获得主机ip
def get_host_ip(hostid):
    try:
        sql='''select host from hosts where status = 0 and hostid = %s''' % hostid
        mycur.execute(sql)
        hostip=mycur.fetchone()
    except Exception as e:
        print("mysql get host ip error",e)
    return hostip
    
# 根据模板ID获得相应模板下面主机ID
def get_hostid_list(templateid):
    hostid_list=[]
    try:
        sql='''SELECT hostid from hosts as a where `status`=0 and a.hostid IN ( SELECT hostid FROM `hosts_templates` where templateid=%d )''' % templateid
        mycur.execute(sql)
        hostidlist=mycur.fetchall()
    except Exception as e:
        print("get win hostidlist error",e)
    
    for i in hostidlist:
        hostid_list.append(i[0])
    
    return hostid_list

# 根据主机ID和键值获得对应的数据
def get_info(hostid,keys):
    host_info_list=[]
    for j in keys:
        try:
            sql = '''select itemid from items where hostid = %s and key_ = '%s' ''' % (hostid, j[2])
            mycur.execute(sql)
            itemid = mycur.fetchone()
        except Exception as e :
            print("itemid is error,",e)
        print("itemid is :",itemid)        
        
        try:
            sql = '''select value_%s as result from %s where itemid = %s and clock >= %s and clock <= %s''' % (j[3],j[1], itemid[0], ts_first, ts_last)
            mycur.execute(sql)
            result = mycur.fetchone()
        except Exception as e:
            print("trends is error",e)
        else:
            print("hostip is :",get_host_ip(hostid))
            hostip=get_host_ip(hostid)
            print("hostid is : %d ,itemid is : %d ,table is : %s , key is : %s ,value is : %s" % (hostid,itemid[0],j[1],j[2],result))
            host_info_list.append([hostip[0],hostid,itemid[0],j[1],j[2],result[0]])
    return host_info_list

# 取得windows平台主机id列表
win_hostid_list=get_hostid_list(win_templateid)
win_hostid_list.remove(10169)
win_hostid_list.remove(10119)
win_hostid_list.remove(10114)
#取得linux 平台主机ID列表
linux_hostid_list=get_hostid_list(linux_templateid)

# 输出测试
# print("windows hostid list is :",win_hostid_list)
# print("linux hostid list is :",linux_hostid_list)
# print("host ip is : ",get_host_ip("10109"))

# 定义excel表起始位置
row=0
col=0

import xlsxwriter as wx

# 创建excel文件对象
workbook1=wx.Workbook('test.xlsx')
# 创建excel文件sheet对象
worksheet1=workbook1.add_worksheet()

# 取windows 主机数值并保存到excel 中

for k in win_hostid_list:
    wininfo=get_info(k, keys_win)
    print('wininfo is :',wininfo)
    for i in wininfo:
        for j in i:
            worksheet1.write(row,6*(wininfo.index(i))+i.index(j ),j)
    row+=1

# 取linux 主机数据值并保存到excel中

for k in linux_hostid_list:
    lnxinfo=get_info(k, keys_linux)
    print('linuxinfo is :',lnxinfo)
    for i in lnxinfo:
        for j in i:
            worksheet1.write(row,6*(lnxinfo.index(i))+i.index(j ),j)
    row+=1

# 关闭excel文件
workbook1.close()

# 关闭数据库
mydb.close()
 