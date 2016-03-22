from openpyxl import load_workbook

def test():
    wb = load_workbook('M:\chartData\Book1.xlsx')
    sheet1 = wb['Sheet1']

    datetimes = []
    for row in sheet1.rows:
        datetime = row[0].value
        a = roundTime(datetime, 60)
        datetimes.append(datetime)

    print 'a'

import datetime
def roundTime(dt, roundTo):
   """Round a datetime object to any time laps in seconds
   dt : datetime.datetime object, default now.
   roundTo : Closest number of seconds to round to, default 1 minute.
   Author: Thierry Husson 2012 - Use it as you want but don't blame me.
   """
   if dt == None :
       dt = datetime.datetime.now()
   seconds = (dt - dt.min).seconds
   # // is a floor division, not a comment on following line:
   rounding = (seconds+roundTo/2) // roundTo * roundTo
   return dt + datetime.timedelta(0,rounding-seconds,-dt.microsecond)

test()