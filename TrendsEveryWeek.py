#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import win32com,sys
from win32com.client import Dispatch, constants

filepath = input(r"路径（默认为当前路径）：") or sys.path[0]+'\\'
basedoc = input("模版（默认为'base.doc'）：") or r'base.doc'
classes=['电气1401','电气1402','电气1403','电气1404','电气1405','电气1406','电气1407','电气1408','电气1409','电气1410','电气1411','药学1401','药学1402','药学1403']
names=['李鹏飞','陈赛珍','杨昌林','邹世博','郑济元','卢睿','陈昊文','张思耕','董春彤','唐建东','史枭迪','余慧','谢弈晖','虞文远']
year=input('年：') or '2015'
month_end=input('结束月份：') or '11'
month_start=input('开始月份（默认等于结束月份）：') or month_end
day_end=input('结束日：') or '16'
day_start=input('开始日（默认等于结束日-6）：') or str(int(day_end)-6)


msword = Dispatch('Word.Application')
msword.DisplayAlerts = 0
msword.Visible = 0

#newdoc = msword.Documents.Add()
basedoc = msword.Documents.Open(FileName = filepath + basedoc)

# 正文文字替换
msword.Selection.Find.ClearFormatting()
msword.Selection.Find.Replacement.ClearFormatting()

i=-1;
for s in classes:
    i+=1
j=i
#替换
while i>=0:
    msword.Selection.Find.Execute(r'(%y%)', False, False, False, False, False, True, 1, True, year, 2)
    msword.Selection.Find.Execute(r'(%ms%)', False, False, False, False, False, True, 1, True, month_start, 2)
    msword.Selection.Find.Execute(r'(%me%)', False, False, False, False, False, True, 1, True, month_end, 2)
    msword.Selection.Find.Execute(r'(%ds%)', False, False, False, False, False, True, 1, True, day_start, 2)
    msword.Selection.Find.Execute(r'(%de%)', False, False, False, False, False, True, 1, True, day_end, 2)
    i-=1
    
Oldname = r'(%name%)'
Oldclass = r'(%class%)'
while j>=0:
    msword.Selection.Find.Execute(Oldname, False, False, False, False, False, True, 1, True, names[j], 2)
    Oldname = names[j]
    msword.Selection.Find.Execute(Oldclass, False, False, False, False, False, True, 1, True, classes[j], 2)
    Oldclass = classes[j]
    basedoc.SaveAs(filepath + classes[j] +'.doc')
    j-=1

#close word
basedoc.Close()
msword.Quit()
