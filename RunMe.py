#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2018-05-05 09:50:06
# @Author  : Frazier (frozen_tearz@163.com)
# @Link    : http://frozentearz.github.io
# @Version : $Python3$

import os
import xlrd
import datetime

schedule = [
    ("08:00 AM", "09:40 AM"),  # 1,2
    ("10:00 AM", "11:40 AM"),  # 3,4
    ############ 中 午 ############
    ("02:30 PM", "04:10 PM"),  # 5,6
    ("04:30 PM", "06:10 PM"),  # 7,8
    ############ 傍 晚 ############
    ("07:30 PM", "09:10 PM"),  # 9,10
]

subjects = []

class Subject():
    '''
    Reference: https://support.google.com/calendar/answer/37118?hl=zh-Hans
    '''
    def __init__(self,
                 subject,
                 start_date,
                 start_time,
                 end_date,
                 end_time,
                 all_day_event,
                 description,
                 location,
                 private):
        self.subject = subject # 课程名
        self.start_date = start_date # 开始日期
        self.start_time = start_time # 开始时间
        self.end_date = end_date # 结束日期
        self.end_time = end_time # 结束时间
        self.all_day_event = all_day_event # 是否全天事件
        self.description = description  # tercher name
        self.location = location # 上课地点
        self.private = private # 是否私人可见

    def __str__(self):
        return "%s, %s, %s, %s, %s, %s, %s, %s, %s" % (
            self.subject,
            self.start_date,
            self.start_time,
            self.end_date,
            self.end_time,
            self.all_day_event,
            self.description, 
            self.location,
            self.private,
        )

def get_date(item, c):
    year = 2018
    month = 3
    day = 6
    semester_start_date = datetime.date(year, month, day)
    # item = '20102 (1-16周)'
    week = item[7:].rstrip('周)')
    if "-" in item:
        start_week = int(week.split("-")[0])
        end_week = int(week.split("-")[1])
    else:
        start_week = int(week)
        end_week = start_week
    weeks = list(range(start_week, end_week+1))
    date = []
    for week in weeks:
        date.append(semester_start_date + datetime.timedelta(weeks=week-1) + datetime.timedelta(days=c-2))
    return date



def get_time(item, r):
    return [schedule[r-3][0], schedule[r-3][1]]

def ParseXls():
    xlspath = './xlsx/ClassSchedule.xlsx'
    xlrd.Book.encoding = "utf-8"
    data = xlrd.open_workbook(xlspath)

    table = data.sheets()[1]  #通过索引顺序获取
    # table = data.sheet_by_index(sheet_indx) # 通过索引顺序获取
    # table = data.sheet_by_name(sheet_name) # 通过名称获取
    
    for r in range(3, table.nrows):
        for c in range(1,table.ncols):
            item = table.cell_value(r, c)
            if len(item) == 0:
                continue
            parseItem(item, c, r)

def parseItem(item, c, r):
    if len(item) == 0:
        return []
    item = item.split("</br>")
    item = item[0].split('\r\n')

    for x in range(len(get_date(item[2], c))):
        subject = "课程：" + item[0]
        start_date = get_date(item[2], c)[x]
        time = get_time(item, r)
        start_time = time[0]
        end_date = start_date  # 一节课只在当天持续
        end_time = time[1]
        description = item[1][:-1]
        location = item[3]

        subject = Subject(subject=subject,
                      start_date=start_date,
                      start_time=start_time,
                      end_date=end_date,
                      end_time=end_time,
                      all_day_event=False,
                      description=description,
                      location=location,
                      private=True
                     )
        subjects.append(str(subject))
    return subjects

def savecvs():
    subjects
    f = open("课表.cvs", "a+", encoding="utf-8")
    f.write("Subject, Start Date, Start Time, End Date, End Time, All Day Event, Description, Location, Private\n")
    for subject in subjects:
        print(subject)
        f.write(subject + "\n")

def main():
    ParseXls()
    savecvs()

if __name__ == '__main__':
    main()