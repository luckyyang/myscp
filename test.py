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
    if sys.argv[1:] is None:
        print "are you kidding? please enter some params to let me to parse, this is why I am a program!"
        print "for help use [-h|--help]"
        sys.exit(2)
    outputFormat = None
    outputFileName = None
    # process options
    for o, a in opts:
        if o in ("-h", "--help"):
            print __doc__
            usage()
            sys.exit(0)
        elif o in ("-f", "--format"):
            outputFormat = a
            if getattr(output_format, output_%s %outputFormat) is None:
                print "can not deal with your format, will use csv format for default"
                usage()
        elif o in ("-o", "--output"):
            outputFileName = a

    today = datetime.date.today()
    currentYear = today.year
    currentMonth = today.month
    currentDay = today.day

    firstDateThisYear = datetime.date(currentYear, 1, 1)
    tempDate = firstDateThisYear

    #outputHead= ['#', 'Year', 'Month', 'Day']
    storeDateIntoList = ['#', 'Year', 'Month', 'Day']

    index = 1

    while (tempDate <= today):
        convertDateToRow = tempDate.isoformat().split('-')
        convertDateToRow.insert(0, str(index))
        storeDateIntoList.append(convertDateToRow)
        tempDate += datetime.timedelta(days=1)
        index += 1

    output_function(outputFormat)

    def output_function(data, outputFileName, outputFormat="csv"):
        output_function = getattr(output_format, output_%s %outputFormat, output_format.output_csv)
        return output_function(data)

class output_format(data):
    def output_csv(data):
        csvFile = open(outputFileName, 'wb')

        CSVFIELDS = ['#', 'Year', 'Month', 'Day']

        writer = csv.writer(csvFile)

        writer.writerow(CSVFIELDS)

        index = 1

        try:
            while (tempDate <= today):
                storeDateIntoList = tempDate.isoformat().split('-')
                toPrintDate = toListDate.insert(0, str(index))
                tempDate += datetime.timedelta(days=1)
                index += 1
                writer.writerow(toListDate)
        finally:
            csvFile.close()

    def output_excel(data):
        book = Workbook()
        sheet1 = book.add_sheet('Sheet 1')

        sheet1.write(0,0,'A1')
        sheet1.write(0,1,'B1')
        row1 = sheet1.row(1)
        row1.write(0,'A2')
        row1.write(1,'B2')
        sheet1.col(0).width = 3000
        book.save('format.xls')
        book.save(TemporaryFile())

        book = open_workbook("format.xls")
        sheet = book.sheet_by_index(0)
        print sheet.name
        print sheet.nrows
        print sheet.ncols
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                print cellname(row_index,col_index),'-',
                    print sheet.cell(row_index,col_index).value


if __name__ == "__main__": 
    main()  
