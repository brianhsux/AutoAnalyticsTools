#!/usr/bin/python
from xlrd import open_workbook
# import xlrd
import xlwt
from datetime import datetime
import rpy2.robjects as robjects
import os, sys
import numpy
import numpy as np

def analytics_CDN_data(str, wb_CDNDataArrangeTotal, ws_CDNDataPerFile):
	wb = open_workbook(str)
	str_name = os.path.splitext(str)[0]
	log_info = "Filtering CDN data: {0}".format(str_name)
	print(log_info)
	ws_CDNDataPerFile = wb_CDNDataArrangeTotal.add_sheet(os.path.splitext(str)[0], cell_overwrite_ok=True)	
	i = 1
	sum_of_json_ok_edge_hits_counts = 0
	sum_of_json_edge_volume_mb = 0
	sum_of_themepack_ok_edge_hits_counts = 0
	sum_of_themepack_edge_volume_mb = 0

	ws_CDNDataPerFile.write(0, 0, str)            
	ws_CDNDataPerFile.write(1, 0, "JSON OK EDGE HITS")
	ws_CDNDataPerFile.write(2, 0, "THEMEPACK OK EDGE HITS:")
	ws_CDNDataPerFile.write(3, 0, "JSON EDGE VOLUME(MB)")	
	ws_CDNDataPerFile.write(4, 0, "THEMEPACK EDGE VOLUME(MB)")

	for sheet in wb.sheets():
	    number_of_rows = sheet.nrows
	    number_of_columns = sheet.ncols
	    theme_str = "ThemeData"
	    json_str = "json"
	    themepack_str = "com.asus.themes"

	    # items = []
	   
	    json_ok_edge_hits_counts = 0    
	    json_edge_volume_mb = 0
	    themepack_ok_edge_hits_counts = 0    
	    themepack_edge_volume_mb = 0    

	    rows = []
	    for row in range(1, number_of_rows):
	        values = []
	        # for col in range(number_of_columns):
	            # value  = (sheet.cell(row,col).value)
	        value_temp  = (sheet.cell(row,0).value)

	        try:
	            value_temp = str(int(value_temp))
	        except ValueError:
	            pass
	        finally:
	            # if any(str_ in value_temp for str_ in ("ThemeData", "json")):
	            if (value_temp.find(theme_str) > 0 and value_temp.find(json_str) > 0):
	            # if ("ThemeData" in value_temp):
	                # json_edge_hits_counts = json_edge_hits_counts + sheet.cell(row,1).value
	                json_ok_edge_hits_counts = json_ok_edge_hits_counts + sheet.cell(row,2).value
	                json_edge_volume_mb = json_edge_volume_mb + sheet.cell(row,4).value
	                # print("find value themedata json")
	                # values.append(value_temp)                
	                # print(edge_hits_counts)
	            if (value_temp.find(theme_str) > 0 and value_temp.find(themepack_str) > 0):
	            # if ("ThemeData" in value_temp):
	                # json_edge_hits_counts = json_edge_hits_counts + sheet.cell(row,1).value
	                themepack_ok_edge_hits_counts = themepack_ok_edge_hits_counts + sheet.cell(row,2).value
	                themepack_edge_volume_mb = themepack_edge_volume_mb + sheet.cell(row,4).value

	    json_ok_edge_hits_counts_in_million = json_ok_edge_hits_counts / 1000000
	    themepack_ok_edge_hits_counts_in_million = themepack_ok_edge_hits_counts / 1000000
	    json_edge_volume_tb = json_edge_volume_mb / 1000000
	    themepack_edge_volume_tb = themepack_edge_volume_mb / 1000000

	    if (json_ok_edge_hits_counts != 0):
		    ws_CDNDataPerFile.write(0, i, sheet.name)
		    ws_CDNDataPerFile.write(1, i, getSecondDecimalPlace(json_ok_edge_hits_counts_in_million))
		    ws_CDNDataPerFile.write(2, i, getSecondDecimalPlace(themepack_ok_edge_hits_counts_in_million))
		    ws_CDNDataPerFile.write(3, i, getSecondDecimalPlace(json_edge_volume_tb))
		    ws_CDNDataPerFile.write(4, i, getSecondDecimalPlace(themepack_edge_volume_tb))
		    i = i + 1

	    sum_of_json_ok_edge_hits_counts += json_ok_edge_hits_counts
	    sum_of_json_edge_volume_mb += json_edge_volume_mb 
	    sum_of_themepack_ok_edge_hits_counts += themepack_ok_edge_hits_counts
	    sum_of_themepack_edge_volume_mb += themepack_edge_volume_mb

	sum_of_json_ok_edge_hits_counts_in_million = sum_of_json_ok_edge_hits_counts / 1000000
	sum_of_themepack_ok_edge_hits_counts_in_million = sum_of_themepack_ok_edge_hits_counts / 1000000
	sum_of_json_edge_volume_tb = sum_of_json_edge_volume_mb / 1000000
	sum_of_themepack_edge_volume_tb = sum_of_themepack_edge_volume_mb / 1000000

	ws_CDNDataPerFile.write(0, i, "Total")
	ws_CDNDataPerFile.write(1, i, getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_in_million))
	ws_CDNDataPerFile.write(2, i, getSecondDecimalPlace(sum_of_themepack_ok_edge_hits_counts_in_million))
	ws_CDNDataPerFile.write(3, i, getSecondDecimalPlace(sum_of_json_edge_volume_tb))
	ws_CDNDataPerFile.write(4, i, getSecondDecimalPlace(sum_of_themepack_edge_volume_tb))
	wb_CDNDataArrangeTotal.save('CDN_analytics_result.xls')

def analytics_CDN_total_data(str, i, wb_CDNDataArrangeTotal, ws_CDNDataPerFile):
	wb = open_workbook(str)
	str_name = os.path.splitext(str)[0]
	log_info = "Handling CDN data: {0}".format(str_name)
	print(log_info)

	sum_of_json_ok_edge_hits_counts = 0
	sum_of_json_edge_volume_mb = 0
	sum_of_themepack_ok_edge_hits_counts = 0
	sum_of_themepack_edge_volume_mb = 0

	for sheet in wb.sheets():
	    number_of_rows = sheet.nrows
	    number_of_columns = sheet.ncols
	    theme_str = "ThemeData"
	    json_str = "json"
	    themepack_str = "com.asus.themes"

	    items = []
	   
	    json_ok_edge_hits_counts = 0    
	    json_edge_volume_mb = 0
	    themepack_ok_edge_hits_counts = 0    
	    themepack_edge_volume_mb = 0    

	    rows = []
	    for row in range(1, number_of_rows):
	        values = []
	        value_temp  = (sheet.cell(row,0).value)

	        try:
	            value_temp = str(int(value_temp))
	        except ValueError:
	            pass
	        finally:
	            if (value_temp.find(theme_str) > 0 and value_temp.find(json_str) > 0):
	                json_ok_edge_hits_counts = json_ok_edge_hits_counts + sheet.cell(row,2).value
	                json_edge_volume_mb = json_edge_volume_mb + sheet.cell(row,4).value
	            if (value_temp.find(theme_str) > 0 and value_temp.find(themepack_str) > 0):
	                themepack_ok_edge_hits_counts = themepack_ok_edge_hits_counts + sheet.cell(row,2).value
	                themepack_edge_volume_mb = themepack_edge_volume_mb + sheet.cell(row,4).value

	    sum_of_json_ok_edge_hits_counts += json_ok_edge_hits_counts
	    sum_of_json_edge_volume_mb += json_edge_volume_mb 
	    sum_of_themepack_ok_edge_hits_counts += themepack_ok_edge_hits_counts
	    sum_of_themepack_edge_volume_mb += themepack_edge_volume_mb

	sum_of_json_ok_edge_hits_counts_in_million = sum_of_json_ok_edge_hits_counts / 1000000
	sum_of_themepack_ok_edge_hits_counts_in_million = sum_of_themepack_ok_edge_hits_counts / 1000000
	sum_of_json_edge_volume_tb = sum_of_json_edge_volume_mb / 1000000	
	sum_of_themepack_edge_volume_tb = sum_of_themepack_edge_volume_mb / 1000000

	sum_of_ok_edge_hits_counts_in_million = sum_of_json_ok_edge_hits_counts_in_million + sum_of_themepack_ok_edge_hits_counts_in_million
	sum_of_edge_volume_tb = sum_of_json_edge_volume_tb + sum_of_themepack_edge_volume_tb
	
	ws_CDNDataPerFile.write(0, 0, "CDN流量計算")
	ws_CDNDataPerFile.write(0, i, str_name[8:14])
	ws_CDNDataPerFile.write(1, i, getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_in_million))	
	ws_CDNDataPerFile.write(2, i, getSecondDecimalPlace(sum_of_themepack_ok_edge_hits_counts_in_million))
	ws_CDNDataPerFile.write(3, i, getSecondDecimalPlace(sum_of_ok_edge_hits_counts_in_million))
	ws_CDNDataPerFile.write(4, i, getSecondDecimalPlace(sum_of_json_edge_volume_tb))
	ws_CDNDataPerFile.write(5, i, getSecondDecimalPlace(sum_of_themepack_edge_volume_tb))
	ws_CDNDataPerFile.write(6, i, getSecondDecimalPlace(sum_of_edge_volume_tb))
	wb_CDNDataArrangeTotal.save('CDN_analytics_result.xls')

def getSecondDecimalPlace(number):
	return float('{:.2f}'.format(number))

def main():
	wb_CDNDataArrangeTotal = xlwt.Workbook()
	ws_CDNDataPerFile = wb_CDNDataArrangeTotal.add_sheet('Total CDN data', cell_overwrite_ok=True)
	ws_CDNDataPerFile.write(1, 0, "JSON OK EDGE HITS(million)")
	ws_CDNDataPerFile.write(2, 0, "THEMEPACK OK EDGE HITS(million)")
	ws_CDNDataPerFile.write(3, 0, "Total EDGE HITS(million)")
	ws_CDNDataPerFile.write(4, 0, "JSON EDGE VOLUME(TB)")
	ws_CDNDataPerFile.write(5, 0, "THEMEPACK EDGE VOLUME(TB)")
	ws_CDNDataPerFile.write(6, 0, "Total EDGE VOLUME(TB)")	

	files = []
	path = "."
	for f in os.listdir(path):
	        if os.path.isfile(f):
	                files.append(f)
	cdn_cal_str = '流量計算'
	xlsx_str = '.xlsx'
	i = 1                
	for f in files:
		os.path.splitext(f)
		if (os.path.splitext(f)[0].find(cdn_cal_str) >= 0 and os.path.splitext(f)[1] == xlsx_str):		
			analytics_CDN_total_data(f, i, wb_CDNDataArrangeTotal, ws_CDNDataPerFile)
			analytics_CDN_data(f, wb_CDNDataArrangeTotal, ws_CDNDataPerFile)
			i = i + 1
			
main()