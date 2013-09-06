import csv
import calendar
import xlwt
import xlrd

from config import *
from pprint import pprint

def runAllReports(list_stu):
    wbk = xlwt.Workbook()
    
    wbk = runTopicReport(list_stu, wbk)
    wbk = runOSReport(list_stu, wbk)   
    wbk = runBrowserReport(list_stu, wbk)
    wbk = runCourseReportByID(list_stu, wbk)
    wbk = runCourseReportBySection(list_stu, wbk)
    wbk = runJavaReport(list_stu, wbk)
    wbk = runCollegeReport(list_stu, wbk)
    
    wbk.save('glennisawesome1.xls') 
    
    return


def runTopicReport(list_stu, wbk):

    sheet = wbk.add_sheet("TOPIC",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][1]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'TOPIC REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk


def runOSReport(list_stu, wbk):

    sheet = wbk.add_sheet("OS",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][2]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'OS REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk


def runBrowserReport(list_stu, wbk):
    sheet = wbk.add_sheet("BROWSER",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][3]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'BROWSER REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk


def runCourseReportByID(list_stu, wbk):

    sheet = wbk.add_sheet("COURSEID",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][4][0].rsplit('-', 1)[0]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'COURSE ID REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk

def runCourseReportBySection(list_stu, wbk):
    
    sheet = wbk.add_sheet("SECTION",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][4][0]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'COURSE SECTION REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk


def runJavaReport(list_stu, wbk):
    # confirmed working 11/15
    
    sheet = wbk.add_sheet("JAVA",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][5]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'JAVA VERSION REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk

def runCollegeReport(list_stu, wbk):
    # confirmed working 11/15
    
    sheet = wbk.add_sheet("COLLEGE",cell_overwrite_ok=True)
    
    final_report = {}
    month_list = []
    item_list = {}

    for student in list_stu:
        date = list_stu[student][0]
        item = list_stu[student][6]
        
        month = calendar.month_abbr[int(date.split('/')[0])]
        monthyear = str(str(date.split('/')[2][0:4] + '/' + date.split('/')[0])) + month

        if monthyear not in month_list:
            month_list.append(monthyear)

        if item not in item_list:
            item_list[item] = []
        
        if monthyear not in final_report:   
            final_report[monthyear] = {item:1}

        elif item not in final_report[monthyear]:
            final_report[monthyear].update({item:1})

        else:
            final_report[monthyear][item] += 1

    sheet.write(0,0,'COLLEGE REPORT')

    month_list.sort()

    col = 0
    row = 2

    for each_item in sorted(item_list.iterkeys()):
        sheet.write(row, col, each_item)
        row += 1

    col = 1

    for each_month in sorted(final_report.iterkeys()):
        row = 1
        sheet.write(row, col, each_month)
        row +=1

        for each_item in sorted(item_list.iterkeys()):
            if (each_item not in final_report[each_month]):
                sheet.write(row, col,0)

            else:
                sheet.write(row, col, final_report[each_month][each_item])

            row +=1

        col +=1
       
    
    return wbk 
