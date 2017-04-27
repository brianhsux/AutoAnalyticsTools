#!/usr/bin/python
from xlrd import open_workbook
from datetime import datetime
from xlutils.copy import copy
from xlwt import easyxf
from csv import reader
from apiclient.discovery import build

import xlrd
import xlwt
import rpy2.robjects as robjects
import os, sys
import numpy
import numpy as np
import json
import argparse
import httplib2

def getSecondDecimalPlace(number):
    return float('{:.2f}'.format(number))

def getIssueWbInfo():
    cwd = os.getcwd()
    file_path = cwd + '\\' + 'Issue_analytics_result.xls'
    rb = open_workbook(file_path, formatting_info=True)
    r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
    issue_wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
    issue_wb_sheet = issue_wb.get_sheet(0) # the sheet to write to within the writable copy
    return [issue_wb, issue_wb_sheet]

def get_analytics_cdn_data(filename, month_index):
    issue_info_results = getIssueWbInfo()
    issue_wb = issue_info_results[0]
    issue_wb_sheet = issue_info_results[1]

    wb = open_workbook(filename)
    str_name = os.path.splitext(filename)[0]
    log_info = "analytics Total CDN data: {0}".format(str_name)
    print(log_info)

    time_str = str_name[8:14]

    sum_of_json_ok_edge_hits_counts_new = 0
    sum_of_json_ok_edge_hits_counts_old = 0
    sum_of_json_ok_edge_hits_counts_total = 0
    sum_of_json_edge_volume_mb_new = 0
    sum_of_json_edge_volume_mb_old = 0
    sum_of_json_edge_volume_mb_total = 0

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        theme_str = "ThemeData"
        json_str = ".json"
        gz_str = ".gz"
        themepack_str = "com.asus.themes"       
        json_ok_edge_hits_counts_new = 0    
        json_edge_volume_mb_new = 0
        json_ok_edge_hits_counts_old = 0    
        json_edge_volume_mb_old = 0

        rows = []
        for row in range(1, number_of_rows):
            values = []
            value_temp  = (sheet.cell(row,0).value)

            try:
                if (value_temp.find(theme_str) > 0 and value_temp.find(json_str) > 0):
                    if (value_temp.find(gz_str) > 0):
                        json_ok_edge_hits_counts_new = json_ok_edge_hits_counts_new + sheet.cell(row,2).value
                        json_edge_volume_mb_new = json_edge_volume_mb_new + sheet.cell(row,4).value
                    else:
                        json_ok_edge_hits_counts_old = json_ok_edge_hits_counts_old + sheet.cell(row,2).value
                        json_edge_volume_mb_old = json_edge_volume_mb_old + sheet.cell(row,4).value
            except BaseException:
                pass

        sum_of_json_ok_edge_hits_counts_new += json_ok_edge_hits_counts_new    
        sum_of_json_edge_volume_mb_new += json_edge_volume_mb_new
        sum_of_json_ok_edge_hits_counts_old += json_ok_edge_hits_counts_old    
        sum_of_json_edge_volume_mb_old += json_edge_volume_mb_old

        print('time_str: {0}'.format(time_str))
        print('sum_of_json_ok_edge_hits_counts_new: {0}'.format(sum_of_json_ok_edge_hits_counts_new))
        print('sum_of_json_ok_edge_hits_counts_old: {0}'.format(sum_of_json_ok_edge_hits_counts_old))
        print('sum_of_json_edge_volume_mb_new: {0}'.format(sum_of_json_edge_volume_mb_new))        
        print('sum_of_json_edge_volume_mb_old: {0}'.format(sum_of_json_edge_volume_mb_old))

    sum_of_json_ok_edge_hits_counts_total = (sum_of_json_ok_edge_hits_counts_new + 
        sum_of_json_ok_edge_hits_counts_old)
    sum_of_json_edge_volume_mb_total = sum_of_json_edge_volume_mb_new + sum_of_json_edge_volume_mb_old

    sum_of_json_ok_edge_hits_counts_new_percent = (str(getSecondDecimalPlace((sum_of_json_ok_edge_hits_counts_new / 
        sum_of_json_ok_edge_hits_counts_total) * 100)) + '%')
    sum_of_json_ok_edge_hits_counts_old_percent = (str(getSecondDecimalPlace((sum_of_json_ok_edge_hits_counts_old / 
        sum_of_json_ok_edge_hits_counts_total) * 100)) + '%')
    sum_of_json_edge_volume_mb_new_percent = (str(getSecondDecimalPlace((sum_of_json_edge_volume_mb_new / 
        sum_of_json_edge_volume_mb_total) * 100)) + '%')
    sum_of_json_edge_volume_mb_old_percent = (str(getSecondDecimalPlace((sum_of_json_edge_volume_mb_old / 
        sum_of_json_edge_volume_mb_total) * 100)) + '%')

    print(sum_of_json_ok_edge_hits_counts_new_percent)
    print(sum_of_json_ok_edge_hits_counts_old_percent)
    print(sum_of_json_edge_volume_mb_new_percent)
    print(sum_of_json_edge_volume_mb_old_percent)

    return ([time_str, 
        getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_new), 
        getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_old),
        getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_total),
        sum_of_json_ok_edge_hits_counts_new_percent, 
        sum_of_json_ok_edge_hits_counts_old_percent,
        getSecondDecimalPlace(sum_of_json_edge_volume_mb_new),
        getSecondDecimalPlace(sum_of_json_edge_volume_mb_old),
        getSecondDecimalPlace(sum_of_json_edge_volume_mb_total),
        sum_of_json_edge_volume_mb_new_percent,
        sum_of_json_edge_volume_mb_old_percent])

def get_analytics_console_data(filename, month_index):
    # installs_com.asus.themeapp_201703_app_version.xlsx
    print(filename[31:33])

    wb = open_workbook(filename)
    sheet = wb.sheets()[0]
    print(sheet.name)

    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    total_user_installs_str = "Total User Installs"
    active_device_installs_str = "Active Device Installs"
    themeapp_str = "com.asus.themeapp"
    rows = []

    enable_json_mechenism = 1510600192    
    totalUserInstalls_enable = 0
    totalUserInstalls_disable = 0
    totalUserInstalls_enable_percent = 0
    totalUserInstalls_disable_percent = 0
    totalUserInstalls_sum = 0

    activeDeviceInstalls_enable = 0
    activeDeviceInstalls_disable = 0
    activeDeviceInstalls_enable_percent = 0
    activeDeviceInstalls_disable_percent = 0
    activeDeviceInstalls_sum = 0

    currentYear = xlrd.xldate_as_tuple(sheet.cell(number_of_rows - 1,0).value, 0)[0]
    currentMonth = xlrd.xldate_as_tuple(sheet.cell(number_of_rows - 1,0).value, 0)[1]
    if (currentMonth < 10):
        time_str = str(currentYear) + '0' + str(currentMonth)
    else:
        time_str = str(currentYear) + str(currentMonth)

    lastDayInmonth = xlrd.xldate_as_tuple(sheet.cell(number_of_rows - 1,0).value, 0)[2]
    print(lastDayInmonth)

    for row in range(1, number_of_rows):
        if (xlrd.xldate_as_tuple(sheet.cell(row,0).value, 0)[2] == lastDayInmonth):
            # AppVersion: sheet.cell(row,2)
            appVersion = int(sheet.cell(row,2).value)
            # cause the xlsx fomat before 2016/12 is different, need to add logic
            if (currentYear == 2016 and currentMonth <= 11):
                totalUserInstalls = int(sheet.cell(row,8).value)
                activeDeviceInstalls = int(sheet.cell(row,11).value)
            else:
                # Total User Installs: sheet.cell(row,6)
                totalUserInstalls = int(sheet.cell(row,6).value)
                # Active Device Installs: sheet.cell(row,9)
                activeDeviceInstalls = int(sheet.cell(row,9).value)

            if (appVersion >= enable_json_mechenism):
                # print('enable appVersion: {0}'.format(appVersion))
                totalUserInstalls_enable = totalUserInstalls_enable + totalUserInstalls
                activeDeviceInstalls_enable = activeDeviceInstalls_enable + activeDeviceInstalls
            else:
                # print('disable appVersion: {0}'.format(appVersion))
                totalUserInstalls_disable = totalUserInstalls_disable + totalUserInstalls
                activeDeviceInstalls_disable = activeDeviceInstalls_disable + activeDeviceInstalls

    totalUserInstalls_sum = totalUserInstalls_enable + totalUserInstalls_disable
    activeDeviceInstalls_sum = activeDeviceInstalls_enable + activeDeviceInstalls_disable

    totalUserInstalls_enable_percent = (str(getSecondDecimalPlace(totalUserInstalls_enable / 
        totalUserInstalls_sum * 100)) + '%')
    totalUserInstalls_disable_percent = (str(getSecondDecimalPlace(totalUserInstalls_disable / 
        totalUserInstalls_sum * 100)) + '%')
    activeDeviceInstalls_enable_percent = (str(getSecondDecimalPlace(activeDeviceInstalls_enable / 
        activeDeviceInstalls_sum  * 100)) + '%')
    activeDeviceInstalls_disable_percent = (str(getSecondDecimalPlace(activeDeviceInstalls_disable / 
        activeDeviceInstalls_sum * 100)) + '%')

    return ([time_str, 
        activeDeviceInstalls_enable, activeDeviceInstalls_disable, activeDeviceInstalls_sum,
        activeDeviceInstalls_enable_percent, activeDeviceInstalls_disable_percent,
        totalUserInstalls_enable, totalUserInstalls_disable, totalUserInstalls_sum,
        totalUserInstalls_enable_percent, totalUserInstalls_disable_percent])

def getSecondDecimalPlace(number):
    return float('{:.2f}'.format(number))

def prepare_data_title():
    cdn_wb = xlwt.Workbook()
    cdn_wb_total_sheet = cdn_wb.add_sheet('Issue analytics', cell_overwrite_ok=True)
    cdn_wb_total_sheet.write(0, 0, "應用Json壓縮機制對於流量的影響\nAppversion>=1.6.0.38_161017")

    active_device_index = 1    
    cdn_wb_total_sheet.write(active_device_index, 0, "Active Device Installs(N)")
    cdn_wb_total_sheet.write(active_device_index + 1, 0, "Active Device Installs(O)")
    cdn_wb_total_sheet.write(active_device_index + 2, 0, "Active Device Installs(Total)")
    cdn_wb_total_sheet.write(active_device_index + 3, 0, "Active Device Installs %(N)")
    cdn_wb_total_sheet.write(active_device_index + 4, 0, "Active Device Installs %(O)")

    total_user_index = 6
    cdn_wb_total_sheet.write(total_user_index, 0, "Total User Installs(N)")
    cdn_wb_total_sheet.write(total_user_index + 1, 0, "Total User Installs(O)")
    cdn_wb_total_sheet.write(total_user_index + 2, 0, "Total User Installs(Total)")
    cdn_wb_total_sheet.write(total_user_index + 3, 0, "Total User Installs %(N)")
    cdn_wb_total_sheet.write(total_user_index + 4, 0, "Total User Installs %(O)")

    access_cdn_index = 11
    cdn_wb_total_sheet.write(access_cdn_index, 0, "成功存取CDN次數(N)")
    cdn_wb_total_sheet.write(access_cdn_index + 1, 0, "成功存取CDN次數(O)")
    cdn_wb_total_sheet.write(access_cdn_index + 2, 0, "成功存取CDN次數(Total)")
    cdn_wb_total_sheet.write(access_cdn_index + 3, 0, "成功存取CDN次數 %(N)")
    cdn_wb_total_sheet.write(access_cdn_index + 4, 0, "成功存取CDN次數 %(O)")

    access_cdn_volume_index = 16
    cdn_wb_total_sheet.write(access_cdn_volume_index, 0, "存取CDN流量(MB)(N)")
    cdn_wb_total_sheet.write(access_cdn_volume_index + 1, 0, "存取CDN流量(MB)(O)")
    cdn_wb_total_sheet.write(access_cdn_volume_index + 2, 0, "存取CDN流量(MB)(Total)")
    cdn_wb_total_sheet.write(access_cdn_volume_index + 3, 0, "存取CDN流量(MB) %(N)")
    cdn_wb_total_sheet.write(access_cdn_volume_index + 4, 0, "存取CDN流量(MB) %(O)")

    cdn_wb.save('Issue_analytics_result.xls')

def storage2xls():
    cwd = os.getcwd()
    file_path = cwd + '\\' + 'Issue_analytics_result.xls'
    rb = open_workbook(file_path, formatting_info=True)
    r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
    wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
    w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy

    console_data_result = []
    cdn_data_result = []

    files = []
    path = "."
    for f in os.listdir(path):
            if os.path.isfile(f):
                    files.append(f)

    app_version_str = 'app_version'
    csv_str = '.csv'
    xlsx_str = '.xlsx'
    month_index = 1                
    for f in files:
        os.path.splitext(f)
        if (os.path.splitext(f)[0].find(app_version_str) >= 0 and os.path.splitext(f)[1] == xlsx_str):
            console_data_result = get_analytics_console_data(f, month_index)  
            print(console_data_result)
            print(console_data_result[0])
            # write time_str
            w_sheet.write(0, month_index, console_data_result[0])

            # write activeDeviceInstalls
            w_sheet.write(1, month_index, console_data_result[1])
            w_sheet.write(2, month_index, console_data_result[2])
            w_sheet.write(3, month_index, console_data_result[3])
            w_sheet.write(4, month_index, console_data_result[4])
            w_sheet.write(5, month_index, console_data_result[5])

            # write totalUserInstalls
            w_sheet.write(6, month_index, console_data_result[6])
            w_sheet.write(7, month_index, console_data_result[7])
            w_sheet.write(8, month_index, console_data_result[8])
            w_sheet.write(9, month_index, console_data_result[9])
            w_sheet.write(10, month_index, console_data_result[10])
            month_index = month_index + 1            

    files = []
    path = "."
    for f in os.listdir(path):
            if os.path.isfile(f):
                    files.append(f)
    cdn_cal_str = '流量計算'
    xlsx_str = '.xlsx'
    month_index = 1                
    for f in files:
        os.path.splitext(f)
        if (os.path.splitext(f)[0].find(cdn_cal_str) >= 0 and os.path.splitext(f)[1] == xlsx_str):      
            cdn_data_result = get_analytics_cdn_data(f, month_index)
            # write access_cdn
            w_sheet.write(11, month_index, cdn_data_result[1])
            w_sheet.write(12, month_index, cdn_data_result[2])
            w_sheet.write(13, month_index, cdn_data_result[3])
            w_sheet.write(14, month_index, cdn_data_result[4])
            w_sheet.write(15, month_index, cdn_data_result[5])
            # write access_cdn_volume
            w_sheet.write(16, month_index, cdn_data_result[6])
            w_sheet.write(17, month_index, cdn_data_result[7])
            w_sheet.write(18, month_index, cdn_data_result[8])
            w_sheet.write(19, month_index, cdn_data_result[9])
            w_sheet.write(20, month_index, cdn_data_result[10])
            month_index = month_index + 1    

    wb.save('Issue_analytics_result.xls')

def storage_console_data():
    prepare_data_title()
    storage2xls()

def main():
    storage_console_data()

if __name__ == '__main__':
    main()