# Import necessary libraries 
import dash
from dash import Dash, html, Input, Output, ctx, dcc, State
import dash_bootstrap_components as dbc

import time
from datetime import datetime

import openpyxl
import threading
import  pandas as pd
import numpy as np

# Import data from excel files
dt = pd.read_excel(r'Data.xlsx')                             ######### Needs URL updated
df = pd.read_excel(r'Agents.xlsx')

# Assign list of names
df['Name'] = df['Name'].str.upper()
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

# Establishes timestamp
now = datetime.now()
timestamp_num = datetime.timestamp(now)
timestamp = datetime.fromtimestamp(timestamp_num)


# Define the page layout
layout = dbc.Container([
    

    dbc.Row([
        html.Center(html.H1("Add/Remove Agent")),
        html.Br(),
        html.Hr(),
        dbc.Col([
            html.P("Select the name of the Agent you want to Add"), 
            
            #Add agent
            html.Div(dcc.Input('', id='input-on-submit', type='text')),
            html.Button('Submit', id='submit-name-btn', n_clicks=0),
            html.Div(id='container-button-basic', children='Enter a Name and press submit'),


        ]), 
        dbc.Col([
            html.P("Select the name of the Agent you want to Remove"), 
            
            #Remove agent
            html.Div(dcc.Dropdown(name_list, first_name_list, id='demo-dropdown', clearable=False)),
            html.Div(id='dd-output-container'),
            html.Button('Remove', id='btn-nclicks-1', n_clicks=0),

        ])
    ])
])



