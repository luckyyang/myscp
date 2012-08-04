#!/usr/bin/env python  

import csv
import datetime

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
    #while (tempDate <= datetime.date(currentYear, 2, 2)):
        toListDate = tempDate.isoformat().split('-')
        toPrintDate = toListDate.insert(0, str(index))
        tempDate += datetime.timedelta(days=1)
        index += 1
        writer.writerow(toListDate)
        #print toListDate
finally:
    csvFile.close()

