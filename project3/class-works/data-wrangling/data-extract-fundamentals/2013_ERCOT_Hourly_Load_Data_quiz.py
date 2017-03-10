#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Thu Mar  9 21:24:29 2017

@author: franklin
"""
"""
Your task is as follows:
- read the provided Excel file
- find and return the min, max and average values for the COAST region
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

import xlrd

datafile = "data/2013_ERCOT_Hourly_Load_Data.xls"

def row_value(sheet, value, col=0):
    row = False
    for r in range(sheet.nrows):
        if sheet.cell_value(r, col) == value:
            row = r
            break
    
    return row
    
    
def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    
    
    #print "Number of cols: "+str(sheet.ncols)
    total_value = sum(sheet.col_values(1, start_rowx=1))
    max_value = max(sheet.col_values(1, start_rowx=1, end_rowx=sheet.nrows))
    min_value = min(sheet.col_values(1, start_rowx=1, end_rowx=sheet.nrows))
    avg_value = total_value/ float(sheet.nrows)
    
    #print total_value
    #print max_value
    #print min_value
    #print avg_value
    
    
    #min_time_excel = min(sheet.col_values(0, start_rowx=1))
    max_time_excel = sheet.cell_value(row_value(sheet, max_value, 1), 0)
    #print max_time_excel
    
    #max_time_excel = max(sheet.col_values(0, start_rowx=1))
    min_time_excel = sheet.cell_value(row_value(sheet, min_value, 1), 0)
    #print max_time_excel
    
    #print xlrd.xldate_as_tuple(max_time_excel, 0)
    
    #print row_value(sheet, max_value, 1)

    ### example on how you can get the data
    #sheet_data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)] for r in range(sheet.nrows)]

    ### other useful methods:
    # print "\nROWS, COLUMNS, and CELLS:"
    # print "Number of rows in the sheet:", 
    # print sheet.nrows
    # print "Type of data in cell (row 3, col 2):", 
    # print sheet.cell_type(3, 2)
    # print "Value in cell (row 3, col 2):", 
    # print sheet.cell_value(3, 2)
    # print "Get a slice of values in column 3, from rows 1-3:"
    # print sheet.col_values(3, start_rowx=1, end_rowx=4)

    # print "\nDATES:"
    # print "Type of data in cell (row 1, col 0):", 
    # print sheet.cell_type(1, 0)
    # exceltime = sheet.cell_value(1, 0)
    # print "Time in Excel format:",
    # print exceltime
    # print "Convert time to a Python datetime tuple, from the Excel float:",
    # print xlrd.xldate_as_tuple(exceltime, 0)
    
    
    data = {
            'maxtime': xlrd.xldate_as_tuple(max_time_excel, 0),
            'maxvalue': max_value,
            'mintime': xlrd.xldate_as_tuple(min_time_excel, 0),
            'minvalue': min_value,
            'avgcoast': avg_value
    }
    return data


def test():
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)


test()