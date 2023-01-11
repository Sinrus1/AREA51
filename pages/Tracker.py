# Import necessary libraries 
import dash
from dash import Dash, html, Input, Output, ctx, dcc, State
import dash_bootstrap_components as dbc

import openpyxl
import threading
import  pandas as pd
import numpy as np


df = pd.read_excel(r'/home/densadam/AREA51/Agents.xlsx')
dt = pd.read_excel(r'/home/densadam/AREA51/Data.xlsx')

df['Name'] = df['Name'].str.upper()
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

uet_worked_today = df.at[0, 'UET_Worked']
uet_worked_total = dt.at[0, 'Total_UET_Tickets']



### Add the page components here 
table_header = [
    html.Thead(html.Tr([html.Th("First Name"), html.Th("Last Name")]))
]

row1 = html.Tr([html.Td("Arthur"), html.Td("Dent")])
row2 = html.Tr([html.Td("Ford"), html.Td("Prefect")])


table_body = [html.Tbody([row1, row2])]

page2_table = dbc.Table(table_header + table_body, bordered=True)

# Define the final page layout
layout = dbc.Container([

    dbc.Row([
        html.Center(html.H1("Ticket Tracker")),
        html.Br(),
        html.Hr(),
        dbc.Col([
            html.P("Select the Agents who are working today"), 
            
            #Agent list mutlti-drop down
            dcc.Dropdown(name_list, name_list, id='work-list', multi=True, persistence= True, persistence_type='session'),
            html.Div(id='mdd-output-container'),

            #Auto-Assign MBM Ticket Button
            html.Br(),
            html.Label('MBM Assign'),
            html.Br(),
            dbc.Button("Auto-Assign MBM Case", color="secondary", id='ambm-btn', n_clicks=0),

            # Manually assign MBM
            html.Br(),
            dcc.Dropdown(name_list, first_name_list,  id='mbm_dropdown', clearable=False),
            html.Button('Manually Assign MBM case', id='mmbm-btn', n_clicks=0),
            html.Div(id='mbm-output-container'),
            
            #Auto-Assign UET Ticket Button
            html.Br(),
            html.Label('UET Assign'),
            html.Br(),
            dbc.Button("Auto-Assign UET Ticket", color="secondary", id='auet-btn', n_clicks=0),

            # Manually assign UET
            html.Br(),
            dcc.Dropdown(name_list, first_name_list,  id='uet_dropdown', clearable=False),
            html.Button('Manually Assign UET ticket', id='muet-btn', n_clicks=0),
            html.Div(id='uet-output-container'),


        ]), 
        dbc.Col([
            html.P("This is column 2."), 



            html.Table([
            #MBM cases display
            html.Tr([html.Td('Next MBM Case is assigned to: '), html.Td(id='mbm-assignee-output')]),            
            html.Tr([html.Td('MBM Cases assigned today = '), html.Td(id='mbm-day-count-output')]),
            html.Tr([html.Td('Total MBM Cases worked = '), html.Td(id='mbm-count-output')]),

            #UET tickets display
            html.Tr([html.Td('Next UET ticket is assigned to: '), html.Td(id='uet-assignee-output')]),
            html.Tr([html.Td('UET Tickets assigned today = '), html.Td(id='uet-day-count-output')]),
            html.Tr([html.Td('Total UET Tickets worked = '), html.Td(id='uet-count-output')]),


            ])
        ])
    ])
])