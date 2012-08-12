#!/usr/bin/env python 

"""This command use to import date from 2012-01-01 to your current date into a file with format csv/excel.
"""
import csv
import sys
import getopt
import datetime
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
    if sys.argv[1:] == []:
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
            if getattr(output_format, "output_%s" %outputFormat) is 'output_None':
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

    storeDateIntoList = ['#', 'Year', 'Month', 'Day']

    index = 1

    while (tempDate <= today):
        convertDateToRow = tempDate.isoformat().split('-')
        convertDateToRow.insert(0, str(index))
        storeDateIntoList.extend(convertDateToRow)
        tempDate += datetime.timedelta(days=1)
        index += 1

    of = output_format()
    output_function(of, storeDateIntoList, outputFileName, outputFormat)

def output_function(classInstance, data, outputFileName, outputFormat="csv"):
    output_function = getattr(classInstance, "output_%s" %outputFormat, output_format.output_csv)
    return output_function(data, outputFileName)

class output_format:
    def output_csv(self, data, outputFileName):
        csvFile = open(outputFileName, 'wb')
        writer = csv.writer(csvFile)

        index = 0
        try:
            while (index <= len(data)):
                writer.writerow(data[index:(index + 4)])
                index += 4
        finally:
            csvFile.close()

    def output_excel(self, data, outputFileName):
        book = Workbook()
        sheet1 = book.add_sheet('Sheet 1')

        index = 0
        rowNumber = 0
        try:
            while (index < len(data)):
                for i in range(4):
                    sheet1.write(rowNumber, i, data[index+i])
                index += 4
                rowNumber += 1
        finally:
            book.save(outputFileName)
            book.save(TemporaryFile())

if __name__ == "__main__": 
    main()  
