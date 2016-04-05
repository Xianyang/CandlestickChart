import sys
from openpyxl import load_workbook
import plotly
from plotly.tools import FigureFactory
from datetime import datetime, timedelta
import plotly.graph_objs as go

plotly.plotly.sign_in('xianyang', 'feoedskn0z')

datetime_data = []          # datetime list for candlestick
datetime_hhl = []           # datetime list for 1 hour high and 1 hour low extension
datetime_xAxis = []         # datetime list for showing on the xAxis
open_data = []              # open price data list
high_data = []              # high price data list
low_data = []               # low price data list
close_data = []             # close price data list
hhx_data = []               # 1 hour high extension price data list
hlx_data = []               # 1 hour low extension price data list
extension1_xDate = []       # datetime list for the first extension
extension1High_yData = []   # high value list for the first extension
extension1Low_yData = []    # low value list for the first extension
extension2_xDate = []       # datetime list for the second extension
extension2High_yData = []   # high value list for the second extension
extension2Low_yData = []    # low value list for the second extension
timeInterval = 0            # the minimum time interval


def cleanTheData():
    global datetime_data
    global datetime_hhl
    global datetime_xAxis
    global open_data
    global high_data
    global low_data
    global close_data
    global hhx_data
    global hlx_data
    global extension1_xDate
    global extension1High_yData
    global extension1Low_yData
    global extension2_xDate
    global extension2High_yData
    global extension2Low_yData
    global timeInterval

    datetime_data = []
    datetime_hhl = []
    datetime_xAxis = []
    open_data = []
    high_data = []
    low_data = []
    close_data = []
    hhx_data = []
    hlx_data = []
    extension1_xDate = []
    extension1High_yData = []
    extension1Low_yData = []
    extension2_xDate = []
    extension2High_yData = []
    extension2Low_yData = []
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

def readData(filename):
    # filename = 'M:\Tanya\charts input.xlsx'
    # filename = 'M:\chartData\charts input - Copy.xlsx'
    # filename = 'M:\chartData\charts input.xlsx'
    # filename = 'M:\chartData\Book1.xlsx'
    # filename = 'M:\chartData\charts10min.xlsx'
    # filename = 'M:\chartData\charts30min.xlsx'
    # filename = 'M:\chartData\charts1hour.xlsx'
    cleanTheData()
    global timeInterval
    try:
        wb = load_workbook(filename)
        sheet1 = wb['Sheet1']
        sheet2 = wb['Sheet2']
        timeInterval = timedelta(seconds=sheet2.rows[0][1].value)
    except ValueError:
        print "Value error, please check the data"
        sys.exit()
    except Exception:
        sys.exit()

    for row in sheet1.rows:
        # Step1 ------ save the candlestick data to list
        if 'date' in str(row[0].value) or row[0].value is None:
            continue
        datetime_data.append(roundTime(row[0].value, 60))
        open_data.append(row[1].value)
        high_data.append(row[2].value)
        low_data.append(row[3].value)
        close_data.append(row[4].value)
        hhx_data.append(row[5].value)
        hlx_data.append(row[6].value)

        # Step2 ------ add extensions
        # the data here is like date[day1_begin, day1_end, day1_end + 1 minute, ...], high[high1, high1, None, ...]
        # the None value in the list is to create a gap
        if row[7].value:
            extension1_xDate.append(roundTime(row[0].value, 60))
            extension1High_yData.append(row[7].value)
            extension1Low_yData.append(row[8].value)
            extension1_xDate.append(roundTime(row[0].value, 60) + row[9].value * timeInterval)
            extension1High_yData.append(row[7].value)
            extension1Low_yData.append(row[8].value)
            extension1_xDate.append(roundTime(row[0].value, 60) + row[9].value * timeInterval + timedelta(minutes=1))
            extension1High_yData.append(None)
            extension1Low_yData.append(None)

        if len(row) > 10 and row[10].value:
            extension2_xDate.append(roundTime(row[0].value, 60))
            extension2High_yData.append(row[10].value)
            extension2Low_yData.append(row[11].value)
            extension2_xDate.append(roundTime(row[0].value, 60) + row[12].value * timeInterval)
            extension2High_yData.append(row[10].value)
            extension2Low_yData.append(row[11].value)
            extension2_xDate.append(roundTime(row[0].value, 60) + row[12].value * timeInterval + timedelta(minutes=1))
            extension2High_yData.append(None)
            extension2Low_yData.append(None)

    # Step3 ------ delete the big time interval
    addPointCount = 0
    indexToInsert = []
    for i in range(0, len(datetime_data)):
        if i > 0:
            timeGap = datetime_data[i] - datetime_data[i - 1]
            if timeGap > 1 * timeInterval + timedelta(minutes=1):
                addPointCount += 1
                indexToInsert.append(i)

    for i in range(0, len(datetime_data) + addPointCount):
        datetime_xAxis.append(datetime_data[0] + i * timeInterval)

    for i in indexToInsert:
        high_data.insert(i + indexToInsert.index(i), None)
        low_data.insert(i + indexToInsert.index(i), None)
        open_data.insert(i + indexToInsert.index(i), None)
        close_data.insert(i + indexToInsert.index(i), None)
        hhx_data.insert(i + indexToInsert.index(i), None)
        hlx_data.insert(i + indexToInsert.index(i), None)
        datetime_data.insert(i + indexToInsert.index(i), '')

    # Step4 ------ Process datetime for 1hhx and 1hlx
    # lastDateTime = timeInterval + datetime_data[-1]
    '''
    for i in range(0, len(hhx_data) - 1):
        if hhx_data[i + 1] != None or (i + 1) == len(hhx_data) - 1:
            datetime_hhl.append(datetime_xAxis[i + 1])
        else:
            datetime_hhl.append(datetime_xAxis[i + 2])
    datetime_hhl.append(datetime_hhl[-1] + timeInterval)
    '''

    # Step5 ------ Process datetime for extensions
    for i in range(0, len(extension1_xDate)):
        if extension1High_yData[i] == 41.27 and extension1Low_yData[i] == 41.28:
            if extension1High_yData[i]:
                extension1_xDate[i] = transFromOriginDateToXaxisData(extension1_xDate[i])
            else:
                extension1_xDate[i] = extension1_xDate[i - 1] + timedelta(minutes=1)
        else:
            if extension1High_yData[i]:
                extension1_xDate[i] = transFromOriginDateToXaxisData(extension1_xDate[i])
            else:
                extension1_xDate[i] = extension1_xDate[i - 1] + timedelta(minutes=1)

    if timeInterval == timedelta(minutes=10) or timeInterval == timedelta(minutes=30):
        for i in range(0, len(extension2_xDate)):
            if extension2High_yData[i]:
                extension2_xDate[i] = transFromOriginDateToXaxisData(extension2_xDate[i])
            else:
                extension2_xDate[i] = extension2_xDate[i - 1] + timedelta(minutes=1)

    print "The data is prepared"


def transFromOriginDateToXaxisData(dateTime):
    if dateTime in datetime_data:
        index = datetime_data.index(dateTime)
        return datetime_xAxis[index]
    else:
        dateTime = dateTime - timeInterval
        return transFromOriginDateToXaxisData(dateTime)


def addLineExtension(figure, xData, yData, name, dashStyle, color):
    extension = go.Scatter(
        x=xData,
        y=yData,
        mode='lines',
        name=name,
        line=dict(
            dash=dashStyle,
            color=color
        )
    )

    figure['data'].extend([extension])

def addDotExtension(figure, xData, yData, name, color, symbol):
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


def createFigure():
    # fig = FF.create_candlestick(open_data, high_data, low_data, close_data, dates=datetime_data)

    # ------------Custom Candlestick Colors------------
    # Make increasing ohlc sticks and customize their color and name
    fig_increasing = FigureFactory.create_candlestick(open_data, high_data, low_data, close_data, dates=datetime_xAxis,
        direction='increasing',
        marker=go.Marker(color='rgb(61, 153, 112)'),
        line=go.Line(color='rgb(61, 153, 112)'))


    # Make decreasing ohlc sticks and customize their color and name
    fig_decreasing = FigureFactory.create_candlestick(open_data, high_data, low_data, close_data, dates=datetime_xAxis,
        direction='decreasing',
        marker=go.Marker(color='rgb(255, 65, 54)'),
        line=go.Line(color='rgb(255, 65, 54)'))

    # Initialize the figure
    fig = fig_increasing
    fig['data'].extend(fig_decreasing['data'])

    # ------------Add extensions------------
    if timeInterval == timedelta(minutes=10) or timeInterval == timedelta(minutes=30):
        addLineExtension(fig, extension1_xDate, extension1High_yData, '1 hour high extension', 'dash', 'green')
        addLineExtension(fig, extension1_xDate, extension1Low_yData, '1 hour low extension', 'dash', 'red')
        addLineExtension(fig, extension2_xDate, extension2High_yData, '1 day high extension', 'dot', 'orange')
        addLineExtension(fig, extension2_xDate, extension2Low_yData, '1 day low extension', 'dot', 'blue')

    elif timeInterval == timedelta(minutes=60):
        addDotExtension(fig, datetime_xAxis, hhx_data, '1 hour high extension', 'orange', 'square')
        addDotExtension(fig, datetime_xAxis, hlx_data, '1 hour low extension', 'blue', 'square')
        addLineExtension(fig, extension1_xDate, extension1High_yData, '1 day high extension', 'dash', 'green')
        addLineExtension(fig, extension1_xDate, extension1Low_yData, '1 day low extension', 'dash', 'red')

    # ------------Update layout And Draw the chart------------
    filename = ''
    if timeInterval == timedelta(minutes=10):
        filename = '10 Min Candlesticks'
    elif timeInterval == timedelta(minutes=30):
        filename = '30 Min Candlesticks'
    elif timeInterval == timedelta(minutes=60):
        filename = '60 Min Candlesticks'

    fig['layout'].update(
        title=filename,
        xaxis=dict(
            ticktext=datetime_data,
            tickvals=datetime_xAxis,
            showticklabels=False,
            showgrid=True,
            showline=True
        )
    )

    # plotly.offline.plot(fig)

    plotly.plotly.plot(fig, filename=filename, sharing='public')
    # py.image.save_as(fig,'data.png')
    print "The chart has been generated"
    # py.iplot(fig, filename='finance/aapl-candlestick-lxy', validate=False)

    '''
    # ------------Adding some dots to a Candlestick Chart------------
    add_hhx = Scatter(
        x = datetime_xAxis,
        y = hhx_data,
        name = '1 hour high extension',
        mode = 'markers',
        marker = Marker(
            color = 'orange',
            symbol = 'square'
        )
    )

    add_hlx = Scatter(
        x = datetime_xAxis,
        y = hlx_data,
        name = '1 hour low extension',
        mode = 'markers',
        marker = Marker(
            color = 'blue',
            symbol = 'square'
        )
    )

    #fig['data'].extend([add_hhx])
    #fig['data'].extend([add_hlx])

    # ------------Define the name of extension--------------
    name1, name2, name3, name4 = '', '', '', ''
    if timeInterval == timedelta(minutes=10) or timeInterval == timedelta(minutes=30):
        # 10 minutes
        name1 = '1 hour high extension'
        name2 = '1 hour low extension'
        name3 = '1 day high extension'
        name4 = '1 day low extension'
    elif timeInterval == timedelta(minutes=60):
        name1 = '1 day high extension'
        name2 = '1 day low extension'
        name3 = ''
        name4 = ''

    # ------------Adding the extension1 to the Candlestick Chart------------
    add_extension1High = Scatter(
        x = extension1_xDate,
        y = extension1High_yData,
        mode = 'lines',
        name = name1,
        line = dict(
            dash = 'dash'
        )
    )

    add_extension1Low = Scatter(
        x = extension1_xDate,
        y = extension1Low_yData,
        mode = 'lines',
        name = name2,
        line = dict(
            dash = 'dash'
        )
    )

    #fig['data'].extend([add_extension1High])
    #fig['data'].extend([add_extension1Low])

    # ------------Adding the extension2 to the Candlestick Chart------------
    if len(extension2_xDate):
        add_extension2High = Scatter(
            x = extension2_xDate,
            y = extension2High_yData,
            mode = 'lines',
            name = name3,
            line = dict(
                dash = 'dot'
            )
        )

        add_extension2Low = Scatter(
            x = extension2_xDate,
            y = extension2Low_yData,
            mode = 'lines',
            name = name4,
            line = dict(
                dash = 'dot'
            )
        )

        #fig['data'].extend([add_extension2High])
        #fig['data'].extend([add_extension2Low])
    '''


def main():
    # Read data
    print 'start read data'
    filenames = ['M:\Chester.Luo\chartData\charts10min.xlsx', 'M:\Chester.Luo\chartData\charts30min.xlsx', 'M:\Chester.Luo\chartData\charts1hour.xlsx']
    # filename = 'M:\chartData\charts10min.xlsx'
    for filename in filenames:
        # filename = ''
        readData(filename)
        createFigure()



if __name__ == "__main__":
    print "***** Candle Stick Generator *****"
    # try:
    main()
    # except:
    #    print "Fail to generate the chart"

"""
# ------------Customizing the Figure with Text and Annotations------------
fig['layout'].update({
    'title': 'The Great Recession',
    'yaxis': {'title': 'AAPL Stock'},
    'shapes': [{
        'x0': '2007-12-01', 'x1': '2007-12-01',
        'y0': 0, 'y1': 1, 'xref': 'x', 'yref': 'paper',
        'line': {'color': 'rgb(30,30,30)', 'width': 1}
    }],
    'annotations': [{
        'x': '2007-12-01', 'y': 0.05, 'xref': 'x', 'yref': 'paper',
        'showarrow': False, 'xanchor': 'left',
        'text': 'Official start of the recession'
    }]
})
"""
