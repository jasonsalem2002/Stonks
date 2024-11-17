# dashboard.py
import plotly.graph_objs as go
from dash import Dash, dcc, html

def create_dashboard(symbols, hist_data):
    app = Dash(__name__)
    
    # Prepare data for the dashboard
    traces = []
    for symbol in symbols:
        if symbol in hist_data and not hist_data[symbol].empty:
            traces.append(go.Scatter(
                x=hist_data[symbol].index,
                y=hist_data[symbol]['Close'],
                mode='lines',
                name=symbol
            ))
    
    app.layout = html.Div([
        html.H1("Stock Analysis Dashboard"),
        dcc.Graph(
            id='comparative-closing-prices',
            figure={
                'data': traces,
                'layout': go.Layout(
                    title='Comparative Closing Prices',
                    xaxis={'title': 'Date'},
                    yaxis={'title': 'Price ($)'},
                    hovermode='closest'
                )
            }
        ),
        # Additional components (e.g., for technical indicators) can be added here
    ])
    
    # Run the Dash app
    app.run_server(debug=False)
