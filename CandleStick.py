import sys
import os
from openpyxl import load_workbook
import plotly
from plotly.tools import FigureFactory
from datetime import datetime, timedelta
import plotly.graph_objs as go


dataRows = []               # data list for all data
datetime_data = []          # datetime list for candlestick
datetime_xAxis = []         # datetime list for showing on the xAxis
extension1 = []             # extension1 list
extension2 = []             # extension2 list
timeInterval = 0            # the minimum time interval

filesForBloomBerg = {'5min': 'M:\\Chester.Luo\\chartData\\charts5min.xlsx',
                     '15min': 'M:\\Chester.Luo\\chartData\\charts15min.xlsx',
                     '60min': 'M:\\Chester.Luo\\chartData\\charts1hour.xlsx'}
filesForOthers = {'5min': 'M:\\chartData\\charts5min.xlsx',
                  '15min': 'M:\\chartData\\charts15min.xlsx',
                  '60min': 'M:\\chartData\\charts1hour.xlsx'}


def cleanTheData():
    global dataRows
    global datetime_data
    global datetime_xAxis
    global extension1
    global extension2
    global timeInterval

    dataRows = []
    datetime_data = []
    datetime_xAxis = []
    extension1 = []
    extension2 = []
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

    seconds = (dt - dt.min).seconds
    # // is a floor division, not a comment on following line:
    rounding = (seconds+roundTo/2) // roundTo * roundTo
    return dt + timedelta(0, rounding-seconds, -dt.microsecond)


class DataRow:
    def __init__(self, row, indexOfRow):
        self.index = indexOfRow
        self.date = (roundTime(row[0].value, 60) if indexOfRow != -1 else None)
        self.open = (row[1].value if indexOfRow != -1 else None)
        self.high = (row[2].value if indexOfRow != -1 else None)
        self.low = (row[3].value if indexOfRow != -1 else None)
        self.close = (row[4].value if indexOfRow != -1 else None)
        self.hhx = (row[5].value if indexOfRow != -1 else None)
        self.hlx = (row[6].value if indexOfRow != -1 else None)
        self.extension1High = (row[7].value if indexOfRow != -1 and len(row) > 9 and row[7].value else None)
        self.extension1Low = (row[8].value if indexOfRow != -1 and len(row) > 9 and row[8].value else None)
        self.extension1LastFor = (row[9].value if indexOfRow != -1 and len(row) > 9 and row[9].value else None)
        self.extension2High = (row[10].value if indexOfRow != -1 and len(row) > 12 and row[10].value else None)
        self.extension2Low = (row[11].value if indexOfRow != -1 and len(row) > 12 and row[11].value else None)
        self.extension2LastFor = (row[12].value if indexOfRow != -1 and len(row) > 12 and row[12].value else None)

    def checkData(self):
        if self.high < self.low:
            print 'The high data is smaller than low data, please check %d data' % (self.index + 1)
            sys.exit()


def readData(filename, lastDays = 0):
    cleanTheData()
    global timeInterval, datetime_xAxis, datetime_data
    try:
        wb = load_workbook(filename)
        sheet1 = wb['Sheet1']
        sheet2 = wb['Sheet2']
        timeInterval = timedelta(seconds=sheet2.rows[0][1].value)
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

    rowsToShow = []
    if lastDays > 0:
        # for timeinterval 5 mins, 10 mins, 15 mins and 30 mins
        indexToCheck = 0
        if timeInterval in [timedelta(minutes=5), timedelta(minutes=10), timedelta(minutes=15), timedelta(minutes=30)]:
            indexToCheck = 10
        elif timeInterval == timedelta(minutes=60):
            indexToCheck = 7
        temp = (list)(sheet1.rows)
        count = lastDays
        while count and len(temp) > 0:
            row = temp.pop()
            rowsToShow.insert(0, row)
            if row[indexToCheck].value:
                count -= 1
    else:
        rowsToShow = sheet1.rows

    for indexOfRow, row in enumerate(rowsToShow):
        # Step1 ------ save the candlestick data to list
        if 'date' in str(row[0].value) or row[0].value is None:
            continue

        datarow = DataRow(row, indexOfRow)
        datarow.checkData()
        dataRows.append(datarow)

        # Step2 ------ add extensions
        # the data here is like date[day1_begin, day1_end, day1_end + 1 minute, ...], high[high1, high1, None, ...]
        # the None value in the list is to create a gap

        if datarow.extension1High:
            extension1.append([datarow.date, datarow.extension1High, datarow.extension1Low])
            extension1.append([datarow.date + datarow.extension1LastFor * timeInterval, datarow.extension1High, datarow.extension1Low])
            # this None data is used to add a gap
            extension1.append([datarow.date + datarow.extension1LastFor * timeInterval + timedelta(minutes=1), None, None])

        if datarow.extension2High:
            extension2.append([datarow.date, datarow.extension2High, datarow.extension2Low])
            extension2.append([datarow.date + datarow.extension2LastFor * timeInterval, datarow.extension2High, datarow.extension2Low])
            # this None data is used to add a gap
            extension2.append([datarow.date + datarow.extension2LastFor * timeInterval + timedelta(minutes=1), None, None])

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
    for i in xrange(0, len(extension1)):
        if extension1[i][1] and extension1[i][2]:
            extension1[i][0] = transFromOriginDateToXaxisData(extension1[i][0])
        else:
            extension1[i][0] = extension1[i - 1][0] + timedelta(minutes=1)

    if timeInterval in [timedelta(minutes=5), timedelta(minutes=10), timedelta(minutes=15), timedelta(minutes=30)]:
        for i in xrange(0, len(extension2)):
            if extension2[i][1] and extension2[i][2]:
                extension2[i][0] = transFromOriginDateToXaxisData(extension2[i][0])
            else:
                extension2[i][0] = extension2[i - 1][0] + timedelta(minutes=1)

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


def createFigure(lastDays = 0, isAutoOpen = True, machine = 'other'):
    # fig = FF.create_candlestick(open_data, high_data, low_data, close_data, dates=datetime_data)
    if lastDays == 0:
        print 'Start Drawing %d minute candlesticks' % (timeInterval.seconds / 60)
    else:
        print 'Start Drawing %d minute candlesticks(%d days)' % (timeInterval.seconds / 60, lastDays)

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
    if timeInterval in [timedelta(minutes=5), timedelta(minutes=10), timedelta(minutes=15), timedelta(minutes=30)]:
        addLineExtension(fig, [extension[0] for extension in extension1], [extension[1] for extension in extension1], '1 hour high extension', 'dash', 'green')
        addLineExtension(fig, [extension[0] for extension in extension1], [extension[2] for extension in extension1], '1 hour low extension', 'dash', 'red')
        addLineExtension(fig, [extension[0] for extension in extension2], [extension[1] for extension in extension2], '1 day high extension', 'dot', 'orange')
        addLineExtension(fig, [extension[0] for extension in extension2], [extension[2] for extension in extension2], '1 day low extension', 'dot', 'blue')

    elif timeInterval == timedelta(minutes=60):
        addDotExtension(fig, datetime_xAxis, [datarow.hhx for datarow in dataRows], '1 hour high extension', 'triangle-down', 'orange')
        addDotExtension(fig, datetime_xAxis, [datarow.hlx for datarow in dataRows], '1 hour low extension', 'triangle-up', 'orange')
        addLineExtension(fig, [extension[0] for extension in extension1], [extension[1] for extension in extension1], '1 day high extension', 'dash', 'green')
        addLineExtension(fig, [extension[0] for extension in extension1], [extension[2] for extension in extension1], '1 day low extension', 'dash', 'red')

    # ------------Update layout And Draw the chart------------
    if lastDays > 0:
        filename = '%dMin-Candlesticks(%ddays)' % (timeInterval.seconds / 60, lastDays)
    else:
        filename = '%dMin-Candlesticks' % (timeInterval.seconds / 60)
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


def transferToServer():
    print 'waiting........the files are transferring to server........'
    import subprocess
    #os.system('candlestickTransferToServer.py')
    subprocess.call('M:\\Matsuba_Charts\\candlestickTransferToServer.py', shell=True)


def readCommand( argv ):
    """
    :param argv: the input from commandl line
    :return: args
    """
    from optparse import OptionParser
    usageStr = """
    USAGE:          python CandleStick.py <options>
    EXAMPLES :      (1) python CandleStick.py start all the charts generating
                    (2) python CandleStick.py --timeinterval 5 --last 5
                        - start generating the last 5 days candlestick chart with time interval 5 mins
    """
    parser = OptionParser(usageStr)

    parser.add_option('-t', '--timeinterval', dest='timeInterval', type='int', help='define which chart to show. If 0, generate all the charts')
    parser.add_option('-l', '--last', dest='lastDays', type='int', help='define how many days to show. If 0, generate all the dates')
    parser.add_option('-o', '--open', dest='autoOpen', type='int', help='define whether the charts should be opened. If 1, open the charts. If 0, not open')
    parser.add_option('-m', '--machine', dest='machine', type='str', help='define which machine the code is running on')

    options, otherjunk = parser.parse_args(argv)
    if len(otherjunk) != 0:
        raise Exception('Command line input not understood: ' + str(otherjunk))

    args = dict()
    args['timeInterval'] = options.timeInterval
    if args['timeInterval'] is None: args['timeInterval'] = 0

    args['lastDays'] = options.lastDays
    if args['lastDays'] is None: args['lastDays'] = 0

    args['isAutoOpen'] = options.autoOpen
    if args['isAutoOpen'] is None: args['isAutoOpen'] = 1

    args['machine'] = options.machine
    if args['machine'] is None: args['machine'] = 'other'

    return args


def runChart(timeInterval, lastDays, isAutoOpen, machine):
    filesname = []
    if machine == 'bloomberg':
        if timeInterval is 5:
            filesname.append(filesForBloomBerg['5min'])
        elif timeInterval is 15:
            filesname.append(filesForBloomBerg['15min'])
        elif timeInterval is 60:
            filesname.append(filesForBloomBerg['60min'])
        elif timeInterval is 0:
            filesname = filesForBloomBerg.values()
        else:
            raise Exception('Wrong time interval, please check -t')

    elif machine == 'other':
        if timeInterval is 5:
            filesname.append(filesForOthers['5min'])
        elif timeInterval is 15:
            filesname.append(filesForOthers['15min'])
        elif timeInterval is 60:
            filesname.append(filesForOthers['60min'])
        elif timeInterval is 0:
            filesname = filesForOthers.values()
        else:
            raise Exception('Wrong time interval, please check -t')

    if len(filesname) == 0:
        raise Exception('No file is found')

    for filename in filesname:
        print '----------start read data----------'
        print filename

        if lastDays == 0:
            readData(filename, lastDays)
            createFigure(lastDays, isAutoOpen, machine)

            readData(filename, 5)
            createFigure(5, isAutoOpen, machine)

        else:
            readData(filename, lastDays)
            createFigure(lastDays, isAutoOpen, machine)

    if machine == 'bloomberg': transferToServer()


def main():
    args = readCommand(sys.argv[1:])
    try:
        runChart(**args)
    except Exception as exc:
        print exc


if __name__ == "__main__":
    print "****** Candle Stick Generator ******"
    main()
