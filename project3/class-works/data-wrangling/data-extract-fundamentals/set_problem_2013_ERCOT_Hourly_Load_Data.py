#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Wed Mar 15 13:30:33 2017

@author: franklin
"""

'''
Find the time and value of max load for each of the regions
COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
and write the result out in a csv file, using pipe character | as the delimiter.

An example output can be seen in the "example.csv" file.
'''

import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "data/2013_ERCOT_Hourly_Load_Data.xls"
outfile = "data/2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    
    header_columns = []
    
    #print "Number of cols: "+str(sheet.ncols)
    
    data = []
    
    for col in range(sheet.ncols):
        header_columns.append(sheet.cell_value(0, col))
        
    
    data.append(header_columns)
    
    i = 1
    for col in range(sheet.ncols-1):
        station_data = []
        if col >= 1:
            station_data.append(header_columns[i])
            max_val_col = max(sheet.col_values(col, start_rowx=1))
            index_max_cell_col = sheet.col_values(col, start_rowx=1).index(max_val_col)+1
            #print max_val_col
            datetime_value_max = xlrd.xldate_as_tuple(sheet.cell_value(index_max_cell_col, 0), 0)
            #print xlrd.xldate_as_tuple(datetime_value_max, 0)
            
            station_data.append(datetime_value_max[0])
            station_data.append(datetime_value_max[1])
            station_data.append(datetime_value_max[2])
            station_data.append(datetime_value_max[3])
            station_data.append(max_val_col)
            i += 1
        data.append(station_data)
            
    #print data
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    return data

def save_file(data, filename):
    titles_csv = ['Station', 'Year', 'Month', 'Day', 'Hour', 'Max Load']
    with open (filename, 'wb') as final_file:
        writer = csv.writer(final_file, delimiter='|')       
        i = 0
        for line in data:
            if i > 0:
                writer.writerow(line)
            else:
                 writer.writerow(titles_csv)
            i += 1
    
    
def test():
    #open_zip(datafile)
    data = parse_file(datafile)
    save_file(data, outfile)

    number_of_rows = 0
    stations = []

    ans = {'FAR_WEST': {'Max Load': '2281.2722140000024',
                        'Year': '2013',
                        'Month': '6',
                        'Day': '26',
                        'Hour': '17'}}
    correct_stations = ['COAST', 'EAST', 'FAR_WEST', 'NORTH',
                        'NORTH_C', 'SOUTHERN', 'SOUTH_C', 'WEST']
    fields = ['Year', 'Month', 'Day', 'Hour', 'Max Load']

    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            print line
            station = line['Station']
            if station == 'FAR_WEST':
                for field in fields:
                    # Check if 'Max Load' is within .1 of answer
                    if field == 'Max Load':
                        max_answer = round(float(ans[station][field]), 1)
                        max_line = round(float(line[field]), 1)
                        assert max_answer == max_line

                    # Otherwise check for equality
                    else:
                        assert ans[station][field] == line[field]

            number_of_rows += 1
            stations.append(station)

        # Output should be 8 lines not including header
        assert number_of_rows == 8

        # Check Station Names
        assert set(stations) == set(correct_stations)

        
if __name__ == "__main__":
    test()
