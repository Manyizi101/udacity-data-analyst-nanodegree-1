#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Thu Mar  9 23:10:41 2017

@author: franklin
"""

import xlrd

datafile = "data/2013_ERCOT_Hourly_Load_Data.xls"

def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)] for r in range(sheet.nrows)]
    
    cv = sheet.col_values(1, start_rowx=1)
    
    maxval = max(cv)
    minval = min(cv)
    
    maxpos = cv.index(maxval) + 1
    minpos = cv.index(minval) + 1
    
    maxtime = sheet.cell_value(maxpos, 0)
    max_realtime = xlrd.xldate_as_tuple(maxtime, 0)
    
    mintime = sheet.cell_value(minpos, 0)
    min_realtime = xlrd.xldate_as_tuple(mintime, 0)
    
    
    data = {
            'maxtime':max_realtime,
            'maxvalue': maxval,
            'mintime':min_realtime,
            'minvalue': minval,
            'avgcoast': sum(cv) / float(len(cv))
    }
    return data


def test():
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)


test()