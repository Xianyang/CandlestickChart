import sys
import platform
import copy
from os import listdir
from os.path import isfile, join
import os
from openpyxl import load_workbook
import csv
import plotly
from plotly.tools import FigureFactory
from datetime import datetime, timedelta
import plotly.graph_objs as go

dataRows = []               # data list for all data
datetime_data = []          # datetime list for candlestick
datetime_xAxis = []         # datetime list for showing on the xAxis
extensionLists = []         # extension list
chartTitle = ''             # title of this chart
timeInterval = 0            # the minimum time interval
daysToShow = 0              # how many days the user want to show
extensionCount = 0          # how many extension the chart has
parameterRowNumber = 5      # how many parameter row


def cleanTheData():
    global extensionLists
    global chartTitle
    global timeInterval
    global daysToShow

    extensionLists = []
    chartTitle = ''
    timeInterval = 0
    daysToShow = 0


def roundTime(dt, roundTo):
    """
    Round a datetime object to any time laps in seconds
    dt : datetime.datetime object, default now.
    roundTo : Closest number of seconds to round to, default 1 minute.
    Author: Thierry Husson 2012 - Use it as you want but don't blame me.
    """
    if dt is None:
        dt = datetime.datetime.now()

    import types
    if type(dt) is types.StringType:
        dt = datetime.strptime(dt, '%m/%d/%Y %H:%M')

    seconds = (dt - dt.min).seconds
    # // is a floor division, not a comment on following line:
    rounding = (seconds+roundTo/2) // roundTo * roundTo
    return dt + timedelta(0, rounding-seconds, -dt.microsecond)


class DataRow:
    """
    initialize a data row using an entry of data
    there have to be date, open, high, low, close data
    the extension data is up to the user
    """
    def __init__(self, row, indexOfRow):
        self.index = indexOfRow
        self.date = (roundTime(row[0], 60) if indexOfRow != -1 else None)
        self.open = (row[1] if indexOfRow != -1 else None)
        self.high = (row[2] if indexOfRow != -1 else None)
        self.low = (row[3] if indexOfRow != -1 else None)
        self.close = (row[4] if indexOfRow != -1 else None)
        self.extensionData = []
        for i in range(6, len(row)):
            self.extensionData.append(float(row[i]) if indexOfRow != -1 and row[i] != '' else None)

    def checkData(self):
        if self.high < self.low:
            print 'The high data is smaller than low data, please check %d data' % (self.index + 1)
            sys.exit()


def readData(filename):
    cleanTheData()
    global timeInterval, datetime_xAxis, datetime_data, chartTitle, daysToShow, extensionCount
    try:
        rowsToShow = []
        with open(filename, 'rb') as csvfile:
            spamreader = csv.reader(csvfile)
            for row in spamreader:
                rowsToShow.append(row)

        chartTitle = rowsToShow[0][0]
        timeInterval = timedelta(minutes=int(rowsToShow[1][1]))
        daysToShow = int(rowsToShow[2][1])
        extensionCount = int(rowsToShow[3][1])
        for i in range(0, extensionCount):
            extensionLists.append([])
        #print 'time interval is %d' %(timeInterval)
    except ValueError:
        print "Value error, please check the data"
        sys.exit()
    except IOError as exc:
        print 'IOError: %s not found, please check file name' % filename
        sys.exit()
    except Exception as exc:
        print 'unexpected error: ', exc.message
        sys.exit()

    if daysToShow > 0:
        indexToCheck = 5
        temp = copy.deepcopy(rowsToShow)
        rowsToShow = []
        count = daysToShow
        while count and len(temp) > parameterRowNumber:
            row = temp.pop()
            rowsToShow.insert(0, row)
            if row[indexToCheck] != '':
                count -= 1
        for i in range(0, parameterRowNumber):
            rowsToShow.insert(0, temp[parameterRowNumber - 1 - i])

    for indexOfRow, row in enumerate(rowsToShow):
        # Step1 ------ save the candlestick data to list
        if indexOfRow < parameterRowNumber:
            continue

        datarow = DataRow(row, indexOfRow - parameterRowNumber)
        datarow.checkData()
        dataRows.append(datarow)

        # Step2 ------ add extensions
        # the data here is like date[day1_begin, day1_end, day1_end + 1 minute, ...], high[high1, high1, None, ...]
        # the None value in the list is to create a gap
        for index, data in enumerate(datarow.extensionData):
            if data:
                extensionLists[index].append([datarow.date, data])
                extensionLists[index].append([datarow.date + timeInterval, data])
                extensionLists[index].append([datarow.date + timeInterval + timedelta(seconds=1), None])

    # Step3 ------ delete the big time interval
    indexToInsert = []
    for i in xrange(1, len(dataRows)):
        timeGap = dataRows[i].date - dataRows[i - 1].date
        # this means that the gap is bigger than one time interval
        if timeGap > 1 * timeInterval + timedelta(minutes=1):
            indexToInsert.append(i)

    emptyDataRow = DataRow([None] * 13, -1)
    for i in indexToInsert:
        dataRows.insert(i + indexToInsert.index(i), emptyDataRow)

    datetime_data = [daterow.date for daterow in dataRows]
    datetime_xAxis = [dataRows[0].date + i * timeInterval for i in xrange(0, len(dataRows))]

    # Step5 ------ Process datetime for extensions
    for extensionList in extensionLists:
        for i in xrange(0, len(extensionList):



def transFromOriginDateToXaxisData(dateTime):
    if dateTime in datetime_data:
        index = datetime_data.index(dateTime)
        return datetime_xAxis[index]
    else:
        dateTime -= timeInterval
        return transFromOriginDateToXaxisData(dateTime)


def addLineExtension(figure, xData, yData, name, dashStyle, color):
    extension = go.Scatter(
        x=xData,
        y=yData,
        mode='lines',
        name=name,
        line=dict(
            color=color,
            dash=dashStyle
        )
    )

    figure['data'].extend([extension])


def addDotExtension(figure, xData, yData, name, symbol, color):
    extension = go.Scatter(
        x=xData,
        y=yData,
        name=name,
        mode='markers',
        marker=go.Marker(
            color=color,
            symbol=symbol
        )
    )

    figure['data'].extend([extension])


def createFigure(filename, isAutoOpen = True, isUploadToServer = False):
    print 'a'


def main():
    path = os.path.dirname(os.path.abspath(__file__))

    filenames = [path+'\\'+f for f in listdir(path) if isfile(join(path, f)) and f.endswith('.csv')]
    for filename in filenames:
        isAutoOpen = True
        isUploadToServer = False

        readData(filename)
        createFigure(filename.replace(path, ''), isAutoOpen, isUploadToServer)

if __name__ == '__main__':
    print "****** Candle Stick Generator ******"
    main()

