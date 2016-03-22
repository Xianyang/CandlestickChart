import plotly.plotly as py
import plotly
import plotly.graph_objs as go
from plotly import tools
import numpy as np

plotly.tools.set_credentials_file(username="xianyang", api_key="feoedskn0z")
"""
trace1 = go.Scatter(
        x=[1, 2, 3],
        y=[5, 6, 7],
        mode='markers',
        marker=go.Marker(
            color='red',
            symbol='square'
        )
    )
data = go.Data([trace1])
#plotly.offline.plot(data)

N = 100
random_x = np.linspace(0, 1, N)
random_y0 = np.random.randn(N)+5
random_y1 = np.random.randn(N)
random_y2 = np.random.randn(N)-5

# Create traces
trace0 = go.Scatter(
    x = random_x,
    y = random_y0,
    mode = 'lines',
    name = 'lines'
)
trace1 = go.Scatter(
    x = random_x,
    y = random_y1,
    mode = 'lines+markers',
    name = 'lines+markers',
    line = dict(
        color = 'red',
        dash = 'dash',
        shape = 'vh'
    ),
    marker = dict(
        color = 'yellow'
    )
)
trace2 = go.Scatter(
    x = random_x,
    y = random_y2,
    mode = 'markers',
    name = 'markers'
)
data = [trace0, trace1, trace2]

# Plot and embed in ipython notebook!
plotly.offline.plot(data, filename='line-mode')
"""

"""
trace0 = go.Scatter(
    x=[1, 2, 3, 3.1, 20, 30, 40],
    y=[2, 3, 4, None, 5, 5, 5]
)
trace1 = go.Scatter(
    x=[2, 3, 4],
    y=[5, 5, 5],
)
trace2 = go.Scatter(
    x=[200, 300, 400],
    y=[4, 4, 4],
)
trace3 = go.Scatter(
    x=[4000, 5000, 6000],
    y=[3, 3, 3],
)

data = [trace0]

layout = go.Layout(
    xaxis=dict(
        type='category',
        autorange=True
    )
)
fig = go.Figure(data=data, layout=layout)
#fig['layout'].update(height=600, width=600, title='Multiple Subplots with Shared Y-Axes')
#plot_url = py.plot(fig, filename='multiple-subplots-shared-yaxes')
plotly.offline.plot(fig)
"""

import plotly.plotly as py
from plotly.graph_objs import Scatter, Data, Layout
import datetime

data = Data([Scatter(x=[0,1,2,3,4,5], y=[0,1,2,30,30,30])])

now = datetime.datetime.now()

layout = dict(
        yaxis = dict(
            #type = 'log',
            #tickvals = [ 1.5, 2.53, 5.99999 ]
        ),
        xaxis = dict(
            ticktext = [0, 1, 2, 30, 40, 50],
            tickvals = [0, 1, 2, 3, 4, 5],
            autorange = True
        )
    )

fig = { 'data':data, 'layout':layout }
plotly.offline.plot(fig)
