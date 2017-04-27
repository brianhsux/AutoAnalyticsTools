#!/usr/bin/python
from xlrd import open_workbook
from datetime import datetime
from xlutils.copy import copy
from xlwt import easyxf
import xlwt
import rpy2.robjects as robjects
import os, sys
import numpy
import numpy as np

def getCdnWbInfo():
	cwd = os.getcwd()
	file_path = cwd + '\\' + 'CDN_analytics_result.xls'
	rb = open_workbook(file_path, formatting_info=True)
	r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
	cdn_wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
	cdn_wb_total_sheet = cdn_wb.get_sheet(0) # the sheet to write to within the writable copy
	return [cdn_wb, cdn_wb_total_sheet]

def analytics_CDNdata(str, month_index):
	cnd_info_results = getCdnWbInfo()
	cdn_wb = cnd_info_results[0]
	cdn_wb_total_sheet = cnd_info_results[1]

	wb = open_workbook(str)
	str_name = os.path.splitext(str)[0]
	log_info = "analytics Total CDN data: {0}".format(str_name)
	print(log_info)

	cdn_wb_individual_sheet = cdn_wb.add_sheet(os.path.splitext(str)[0], cell_overwrite_ok=True)	
	col_index = 1

	sum_of_json_ok_edge_hits_counts = 0
	sum_of_json_edge_volume_mb = 0
	sum_of_themepack_ok_edge_hits_counts = 0
	sum_of_themepack_edge_volume_mb = 0

	cdn_wb_individual_sheet.write(0, 0, str)            
	cdn_wb_individual_sheet.write(1, 0, "JSON OK EDGE HITS")
	cdn_wb_individual_sheet.write(2, 0, "THEMEPACK OK EDGE HITS:")
	cdn_wb_individual_sheet.write(3, 0, "JSON EDGE VOLUME(MB)")	
	cdn_wb_individual_sheet.write(4, 0, "THEMEPACK EDGE VOLUME(MB)")

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
			value_temp  = sheet.cell(row,0).value
			try:
				if (value_temp.find(theme_str) > 0 and value_temp.find(json_str) > 0):
					json_ok_edge_hits_counts = json_ok_edge_hits_counts + sheet.cell(row,2).value
					json_edge_volume_mb = json_edge_volume_mb + sheet.cell(row,4).value
				if (value_temp.find(theme_str) > 0 and value_temp.find(themepack_str) > 0):
					themepack_ok_edge_hits_counts = themepack_ok_edge_hits_counts + sheet.cell(row,2).value
					themepack_edge_volume_mb = themepack_edge_volume_mb + sheet.cell(row,4).value
			except BaseException:
				pass

		json_ok_edge_hits_counts_in_million = json_ok_edge_hits_counts / 1000000
		themepack_ok_edge_hits_counts_in_million = themepack_ok_edge_hits_counts / 1000000
		json_edge_volume_tb = json_edge_volume_mb / 1000000
		themepack_edge_volume_tb = themepack_edge_volume_mb / 1000000

		if (json_ok_edge_hits_counts != 0):
			cdn_wb_individual_sheet.write(0, col_index, sheet.name)
			cdn_wb_individual_sheet.write(1, col_index, getSecondDecimalPlace(json_ok_edge_hits_counts_in_million))
			cdn_wb_individual_sheet.write(2, col_index, getSecondDecimalPlace(themepack_ok_edge_hits_counts_in_million))
			cdn_wb_individual_sheet.write(3, col_index, getSecondDecimalPlace(json_edge_volume_tb))
			cdn_wb_individual_sheet.write(4, col_index, getSecondDecimalPlace(themepack_edge_volume_tb))
			col_index = col_index + 1

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
	
	cdn_wb_total_sheet.write(0, 0, "CDN流量計算")
	cdn_wb_total_sheet.write(0, month_index, str_name[8:14])
	cdn_wb_total_sheet.write(1, month_index, getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_in_million))	
	cdn_wb_total_sheet.write(2, month_index, getSecondDecimalPlace(sum_of_themepack_ok_edge_hits_counts_in_million))
	cdn_wb_total_sheet.write(3, month_index, getSecondDecimalPlace(sum_of_ok_edge_hits_counts_in_million))
	cdn_wb_total_sheet.write(4, month_index, getSecondDecimalPlace(sum_of_json_edge_volume_tb))
	cdn_wb_total_sheet.write(5, month_index, getSecondDecimalPlace(sum_of_themepack_edge_volume_tb))
	cdn_wb_total_sheet.write(6, month_index, getSecondDecimalPlace(sum_of_edge_volume_tb))

	cdn_wb_individual_sheet.write(0, col_index, "Total")
	cdn_wb_individual_sheet.write(1, col_index, getSecondDecimalPlace(sum_of_json_ok_edge_hits_counts_in_million))
	cdn_wb_individual_sheet.write(2, col_index, getSecondDecimalPlace(sum_of_themepack_ok_edge_hits_counts_in_million))
	cdn_wb_individual_sheet.write(3, col_index, getSecondDecimalPlace(sum_of_json_edge_volume_tb))
	cdn_wb_individual_sheet.write(4, col_index, getSecondDecimalPlace(sum_of_themepack_edge_volume_tb))

	cdn_wb.save('CDN_analytics_result.xls')

def getSecondDecimalPlace(number):
	return float('{:.2f}'.format(number))

def prepare_data_title():
	cdn_wb = xlwt.Workbook()
	cdn_wb_total_sheet = cdn_wb.add_sheet('Total CDN data', cell_overwrite_ok=True)
	cdn_wb_total_sheet.write(1, 0, "JSON OK EDGE HITS(million)")
	cdn_wb_total_sheet.write(2, 0, "THEMEPACK OK EDGE HITS(million)")
	cdn_wb_total_sheet.write(3, 0, "Total EDGE HITS(million)")
	cdn_wb_total_sheet.write(4, 0, "JSON EDGE VOLUME(TB)")
	cdn_wb_total_sheet.write(5, 0, "THEMEPACK EDGE VOLUME(TB)")
	cdn_wb_total_sheet.write(6, 0, "Total EDGE VOLUME(TB)")
	cdn_wb.save('CDN_analytics_result.xls')	

def main():
	prepare_data_title()

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
			analytics_CDNdata(f, month_index)
			month_index = month_index + 1
			
if __name__ == '__main__':
    main()