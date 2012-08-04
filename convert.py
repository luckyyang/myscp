#!/usr/bin/env python 

"""This command use to import date from 2012-01-01 to your current date into a file with format csv/excel.

usage: python --format [excel|csv] --output filename
"""

import csv
import sys
import getopt
import datetime

def

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

    csvFile = open('test.xls', 'wb')

    CSVFIELDS = ['#', 'Year', 'Month', 'Day']

    writer = csv.writer(csvFile)

    writer.writerow(CSVFIELDS)

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
