# version 1.1.1
# Auther landuoguo

import json
import os
import sys
import time
import pandas as pd
import glob
from pandas import DataFrame
import re

date = "2023-01-05"#修改为需统计的会议当天的日期
student_count = 61#修改为学生人数
path = input("输入文件路径：")

file_data = pd.read_excel(io = path, sheet_name = 1,header=None)
#print(file_data)
file_data_len = len(file_data)

namelist = pd.read_excel(io = "namelist.xlsx", sheet_name = 0,header=None)

if os.path.exists('time.json') == False:
    print('Err: 配置文件json丢失！')
    os.system("pause")
    sys.exit()
with open(r'time.json', 'r', encoding='utf8')as fp: conf = json.load(fp)

temp_d = []
for i in range(0,student_count+1):
    temp_d += [[]]

for i in range(10,file_data_len):
    is_student = "号" in file_data[0][i]
    if is_student:
        nick = re.search(r'\(([^)]+)\)[^)]*\Z', file_data[0][i]).group(1)
        sit = re.findall('[0-9]+\号', nick)[0].strip('号')
        
        #temp_d[int(sit)].append("新增元素")
        temp_d[int(sit)] = temp_d[int(sit)]+[{'start':file_data[1][i],'end':file_data[2][i]}]

#print(temp_d)

temp_d2 = {}
temp_d2['sit'] = {}
temp_d2['name'] = {}

limit_time = []
limit_time2 = []
limit_title_list = []

for i in range(1,student_count+1):
    temp_d2["sit"][i-1] = i
    temp_d2["name"][i-1] = namelist[1][i]

for i1 in conf:
    temp_d2[i1['name']] = {}
    xstart = time.strptime(date+" "+i1['start'],'%Y-%m-%d %H:%M:%S')
    xend = time.strptime(date+" "+i1['end'],'%Y-%m-%d %H:%M:%S')
    limit_time += [time.mktime(xend)-time.mktime(xstart)-60]
    limit_time2 += [time.mktime(xend)-time.mktime(xstart)-180]
    limit_title_list += [i1['name']]
    for i2 in range(1,student_count+1):
        duration = 0
        for i3 in temp_d[i2]:
            starttime = time.strptime(i3['start'],'%Y-%m-%d %H:%M:%S')
            if i3['end']=="--":
                endtime = time.strptime('2060-01-01 00:00:00','%Y-%m-%d %H:%M:%S')
            else:
                endtime = time.strptime(i3['end'],'%Y-%m-%d %H:%M:%S')
            #开始判断
            if(endtime>=xend and starttime<=xstart):
                duration += time.mktime(xend)-time.mktime(xstart)
            elif(endtime>=xend and starttime<xend and starttime>=xstart):
                duration += time.mktime(xend)-time.mktime(starttime)
            elif(starttime<=xstart and endtime<=xend and endtime>xstart):
                duration += time.mktime(endtime)-time.mktime(xstart)
            elif(starttime>=xstart and endtime<=xend):
                duration += time.mktime(endtime)-time.mktime(starttime)
        if (time.mktime(xend)-time.mktime(xstart))<duration:
            duration = time.mktime(xend)-time.mktime(xstart)
        temp_d2[i1['name']][i2-1] = duration
            

#print(temp_d2)

outpath = "new提取 "+path

writer = DataFrame(temp_d2)
#高亮显示时间不足的学生
writer = writer.style.highlight_between(axis=1,left=limit_time2,right=limit_time,subset=limit_title_list,inclusive="right",props='font-weight:bold;color:#ffffff;background-color:orange')\
.highlight_between(axis=1,right=limit_time2,subset=limit_title_list,inclusive="both",props='font-weight:bold;color:#ffffff;background-color:red')
#输出excel
writer.to_excel(outpath, sheet_name='1', index=False, header=True)
    
print(outpath+"   已生成")
print("=======Success=======")
#os.system("pause")
