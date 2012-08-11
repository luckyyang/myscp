#!/usr/bin/env python 

"""This command use to import date from 2012-01-01 to your current date into a file with format csv/excel.

"""

import csv
import sys
import getopt
import datetime
from xlrd import open_workbook,cellname
from tempfile import TemporaryFile
from xlwt import Workbook

def usage():
    print "usage: python %s --format [excel|csv] --output filename" %sys.argv[0]

def main():
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hf:o:", ["help", "format=", "output="])
    except getopt.GetoptError, msg:
        print msg
        print "for help use [-h|--help]"
        sys.exit(2)
    outputFormat = None
    outputFileName = None
    # process options
    for o, a in opts:
        if o in ("-h", "--help"):
            print __doc__;
            usage()
            sys.exit(0);
        elif o in ("-f", "--format"):
            outputFormat = a;
        elif o in ("-o", "--output"):
            outputFileName = a;

    today = datetime.date.today()
    currentYear = today.year
    currentMonth = today.month
    currentDay = today.day

    firstDateThisYear = datetime.date(currentYear, 1, 1)
    tempDate = firstDateThisYear

    csvFile = open(outputFileName, 'wb')

    CSVFIELDS = ['#', 'Year', 'Month', 'Day']

    writer = csv.writer(csvFile)

    writer.writerow(CSVFIELDS)



    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')

    printDate = []
    today = datetime.date.today()
    printDate = str(today).split('-')
    print printDate

    book.add_sheet('Sheet 2')
    sheet1.write(0,0,'A1')
    sheet1.write(0,1,'B1')
    row1 = sheet1.row(1)
    row1.write(0,'A2')
    row1.write(1,'B2')
    sheet1.col(0).width = 3000
    book.save('format.xls')
    book.save(TemporaryFile())

    print datetime.date.today()

    #book = open_workbook('data.xls')
    book = open_workbook("format.xls")
    sheet = book.sheet_by_index(0)
    print sheet.name
    print sheet.nrows
    print sheet.ncols
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            print cellname(row_index,col_index),'-',
                print sheet.cell(row_index,col_index).value




    index = 1

    try:
        while (tempDate <= today):
            toListDate = tempDate.isoformat().split('-')
            toPrintDate = toListDate.insert(0, str(index))
            tempDate += datetime.timedelta(days=1)
            index += 1
            writer.writerow(toListDate)
    finally:
        csvFile.close()

if __name__ == "__main__": 
    main()  
