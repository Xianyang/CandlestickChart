import sys
import platform
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
activations = []            # activation list
limitations = []            # limitation list
takeProfits = []            # top profit list
stopLosses = []             # stop lost list
timeInterval = 0            # the minimum time interval


def cleanTheData():
    global dataRows
    global datetime_data
    global datetime_xAxis
    global activations
    global limitations
    global takeProfits
    global stopLosses
    global timeInterval

    dataRows = []
    datetime_data = []
    datetime_xAxis = []
    activations = []
    limitations = []
    takeProfits = []
    stopLosses = []
    timeInterval = 0


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
    def __init__(self, row, indexOfRow):
        self.index = indexOfRow
        self.date = (roundTime(row[0], 60) if indexOfRow != -1 else None)
        self.open = (row[1] if indexOfRow != -1 else None)
        self.high = (row[2] if indexOfRow != -1 else None)
        self.low = (row[3] if indexOfRow != -1 else None)
        self.close = (row[4] if indexOfRow != -1 else None)
        self.activation = (float(row[5]) if indexOfRow != -1 and len(row) > 5 and row[5] else None)
        self.limitation = (float(row[6]) if indexOfRow != -1 and len(row) > 6 and row[6] else None)
        self.takeProfit = (float(row[7]) if indexOfRow != -1 and len(row) > 7 and row[7] else None)
        self.stopLoss = (float(row[8]) if indexOfRow != -1 and len(row) > 8 and row[8] else None)

    def checkData(self):
        if self.high < self.low:
            print 'The high data is smaller than low data, please check %d data' % (self.index + 1)
            sys.exit()


def readData(filename, lastDays=0):
    cleanTheData()
    global timeInterval, datetime_xAxis, datetime_data
    try:
        rowsToShow = []
        with open(filename, 'rb') as csvfile:
            spamreader = csv.reader(csvfile)
            for row in spamreader:
                rowsToShow.append(row)
        timeInterval = timedelta(minutes=15)
        print 'time interval is %d' %(timeInterval.seconds / 60)
    except ValueError:
        print "Value error, please check the data"
        sys.exit()
    except IOError as exc:
        print 'IOError: %s not found, please check file name' % filename
        sys.exit()
    except Exception as exc:
        print 'unexpected error: ', exc.message
        sys.exit()

    for indexOfRow, row in enumerate(rowsToShow):
        # Step1 ------ save the candlestick data to list
        if 'DATE' in str(row[0]) or row[0] is None:
            continue

        datarow = DataRow(row, indexOfRow)
        datarow.checkData()
        dataRows.append(datarow)

        # Step2 ------ add extensions
        # the data here is like date[day1_begin, day1_end, day1_end + 1 minute, ...], high[high1, high1, None, ...]
        # the None value in the list is to create a gap

        if datarow.activation:
            activations.append([datarow.date, datarow.activation])
            activations.append([datarow.date + timeInterval, datarow.activation])
            # this None data is used to add a gap
            activations.append([datarow.date + timeInterval + timedelta(seconds=1), None])

        if datarow.limitation:
            limitations.append([datarow.date, datarow.limitation])
            limitations.append([datarow.date + timeInterval, datarow.limitation])
            # this None data is used to add a gap
            limitations.append([datarow.date + timeInterval + timedelta(seconds=1), None])

        if datarow.takeProfit:
            takeProfits.append([datarow.date, datarow.takeProfit])
            takeProfits.append([datarow.date + timeInterval, datarow.takeProfit])
            # this None data is used to add a gap
            takeProfits.append([datarow.date + timeInterval + timedelta(seconds=1), None])

        if datarow.stopLoss:
            stopLosses.append([datarow.date, datarow.stopLoss])
            stopLosses.append([datarow.date + timeInterval, datarow.stopLoss])
            # this None data is used to add a gap
            stopLosses.append([datarow.date + timeInterval + timedelta(seconds=1), None])

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

    for i in xrange(0, len(activations)):
        if activations[i][1]:
            activations[i][0] = transFromOriginDateToXaxisData(activations[i][0])
            if i % 3 == 1 and activations[i][0] == activations[i - 1][0]:
                activations[i][0] += timeInterval
        else:
            activations[i][0] = activations[i - 1][0] + timedelta(seconds=1)

    for i in xrange(0, len(limitations)):
        if limitations[i][1]:
            limitations[i][0] = transFromOriginDateToXaxisData(limitations[i][0])
            if i % 3 == 1 and limitations[i][0] == limitations[i - 1][0]:
                limitations[i][0] += timeInterval
        else:
            limitations[i][0] = limitations[i - 1][0] + timedelta(seconds=1)

    for i in xrange(0, len(takeProfits)):
        if takeProfits[i][1]:
            takeProfits[i][0] = transFromOriginDateToXaxisData(takeProfits[i][0])
            if i % 3 == 1 and takeProfits[i][0] == takeProfits[i - 1][0]:
                takeProfits[i][0] += timeInterval
        else:
            takeProfits[i][0] = takeProfits[i - 1][0] + timedelta(seconds=1)

    for i in xrange(0, len(stopLosses)):
        if stopLosses[i][1]:
            stopLosses[i][0] = transFromOriginDateToXaxisData(stopLosses[i][0])
            if i % 3 == 1 and stopLosses[i][0] == stopLosses[i - 1][0]:
                stopLosses[i][0] += timeInterval
        else:
            stopLosses[i][0] = stopLosses[i - 1][0] + timedelta(seconds=1)

    print "The data is prepared"


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


def createFigure(filename, lastDays = 0, isAutoOpen = True, machine = 'other'):
    # ------------Custom Candlestick Colors------------
    # Make increasing ohlc sticks and customize their color and name
    opens = [datarow.open for datarow in dataRows]
    highs = [datarow.high for datarow in dataRows]
    lows = [datarow.low for datarow in dataRows]
    closes = [datarow.close for datarow in dataRows]

    fig_increasing = FigureFactory.create_candlestick(opens, highs, lows, closes, dates=datetime_xAxis,
        direction='increasing',
        marker=go.Marker(color='rgb(61, 153, 112)'),
        line=go.Line(color='rgb(61, 153, 112)'))

    # Make decreasing ohlc sticks and customize their color and name
    fig_decreasing = FigureFactory.create_candlestick(opens, highs, lows, closes, dates=datetime_xAxis,
        direction='decreasing',
        marker=go.Marker(color='rgb(255, 65, 54)'),
        line=go.Line(color='rgb(255, 65, 54)'))

    # Initialize the figure
    fig = fig_increasing
    fig['data'].extend(fig_decreasing['data'])

    # ------------Add extensions------------
    addLineExtension(fig, [activation[0] for activation in activations], [activation[1] for activation in activations], 'Activation', 'line', 'orange')
    addLineExtension(fig, [limitation[0] for limitation in limitations], [limitation[1] for limitation in limitations], 'Limit Order', 'line', 'blue')
    addLineExtension(fig, [takeProfit[0] for takeProfit in takeProfits], [takeProfit[1] for takeProfit in takeProfits], 'Take Profit', 'line', 'green')
    addLineExtension(fig, [stopLoss[0] for stopLoss in stopLosses], [stopLoss[1] for stopLoss in stopLosses], 'Stop Loss', 'line', 'red')

    # ------------Update layout And Draw the chart------------
    filename = filename[0:-4]
    fig['layout'].update(
        title=filename,
        xaxis=dict(
            ticktext=[datarow.date for datarow in dataRows],
            tickvals=datetime_xAxis,
            showticklabels=False,
            showgrid=True,
            showline=True
        ),
    )

    if machine == 'bloomberg':
        path = 'M:\\Matsuba_Charts\\'
        plotly.offline.plot(fig, filename=path + filename + '.html', auto_open=isAutoOpen)
    else:
        plotly.offline.plot(fig, filename=filename + '.html', auto_open=isAutoOpen)
    # plotly.plotly.plot(fig, filename=filename, sharing='public')
    # py.image.save_as(fig,'data.png')
    print "----------The chart has been generated----------\n"


def main():
    machine = platform.node()

    if machine == 'bloomberg':
        path = 'M:\\Chester.Luo\\Evine\\'
    elif machine == 'KOI-PC6':
        path = 'M:\\Evine\\'
    else:
        path = ''

    filenames = [path+f for f in listdir(path) if isfile(join(path, f)) and f.endswith('.csv')]
    for filename in filenames:
        isAutoOpen = True
        machine = 'other'

        readData(filename)
        createFigure(filename.replace(path, ''), 0, isAutoOpen, machine)


if __name__ == "__main__":
    print "****** Candle Stick Generator ******"
    main()
