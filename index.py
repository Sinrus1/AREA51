# Import necessary libraries 
import dash
from dash import dash, html, Input, Output, ctx, dcc, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc

import time
import threading
from datetime import datetime
import xlsxwriter
import openpyxl
import pandas as pd
import numpy as np

# Connect to main app.py file
from app import app

# Connect to your app pages
from pages import Tracker, Add_Remove_Agent

# Connect the navbar to the index
from components import navbar



#define the navbar
nav = navbar.Navbar()

# import name and data files
dt = pd.read_excel(r'Data.xlsx')
df = pd.read_excel(r'Agents.xlsx')

#Assigns value of total tickets worked
mbm_total_output = dt.at[0, 'Total_MBM_Cases']


df['Name'] = df['Name'].str.upper()
name_list = df['Name']
first_name_list = df.loc[0, 'Name']


# Establishes timestamp
def timestamp():
    now = datetime.now()
    timestamp_num = datetime.timestamp(now)
    timestamp = datetime.fromtimestamp(timestamp_num)
    return timestamp

    


# Sets up today's date and yesterdays data
timestamp = timestamp()

day = timestamp.strftime("%d")
today = day
yesterday = dt.loc[0,'Date']
dt.loc[0, 'Date'] = today
timestamp_now = pd.to_datetime('today').strftime("%Y-%m-%d %H:%M:%S")


# checks if time is different day, if true, resets tickets worked to zero
if int(today) != yesterday:
    print('Different day, tickets worked was reset.')
    df['MBM_Worked'] = 0
    df['UET_Worked'] = 0
    
    df.to_excel('Agents.xlsx', index = False)
    dt.to_excel('Data.xlsx', index = False)

else:
    pass






# Define the index page layout
app.layout = html.Div([


    # Core Homepage
    dcc.Location(id='url', refresh=True),
    nav, 
    html.Div(id='page-content', children=[]), 
])

@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/Tracker':
        return Tracker.layout
    if pathname == '/Add_Remove_Agent':
        return Add_Remove_Agent.layout

    else:
        return "Please choose a link"





#######################
#### Add New Agent ####
#######################


@app.callback(
    Output('container-button-basic', 'children'),
    Input('submit-name-btn', 'n_clicks'),
    State('input-on-submit', 'value'),
    prevent_initial_call=True
)

def update_add_remove_agent(n_clicks, value):
    
    df = pd.read_excel(r'Agents.xlsx')
    
    #checks if value is blank
    if value.isalpha() == False:
        msg = 'No spaces, numbers or special characters are allowed'

    else:
        if "submit-name-btn" == ctx.triggered_id:
            value_name = value
            selected_name = value_name.upper()

            # Creates seperate list to compare if duplicates exist.
            df = pd.read_excel(r'Agents.xlsx')                            ######### Needs URL updated
            dup_check = True
            name_list = df['Name']
            name_list.loc[len(name_list)] = selected_name
        
            #If duplicates exist on seperate list, error message generated. Else, add name to main list.     
            if len(name_list) != len(set(name_list)):
                msg = 'Duplicate name, please enter a different name!'
                dup_check = True
            else:
                msg = 'Adding {} to the list of agents!'.format(selected_name)
                dup_check = False


            #Checks if duplicates status is false, then adds name to DF and DT dataframes.    
            if dup_check == False:
                add_name = [selected_name, 0, 0, 0, 0, False]
                df.loc[len(df)] = add_name
        
                writer = pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', engine='xlsxwriter')   ########Need to change directory
                with pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx') as writer:                   ########Need to change directory

                    dict = {selected_name+'_MBM_Worked': [timestamp], "Action": ['Agent Created']}
                    da_m = pd.DataFrame(dict)
                    dict = {selected_name+'_UET_Worked': [timestamp], "Action": ['Agent Created']}
                    da_u = pd.DataFrame(dict)
                        
                    da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                    da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)

                df['Name'] = df['Name'].str.upper()
                name_list = df['Name']
                first_name_list = df.loc[0, 'Name']

                df.to_excel('Agents.xlsx', index = False)                      ########Need to change directory




                
        else:
            pass

    return html.Div(msg)
    



######################
#### Remove Agent ####                                                         
######################


@app.callback(
    Output('dd-output-container', 'children'),
    Input('demo-dropdown', 'value'),
    Input('btn-nclicks-1', 'n_clicks'),
    prevent_initial_call=True
)
def update_output(value, n_clicks):
    value_name = value
    msg1 = 'Select name of the agent who you want to remove from the list of agents'

    if "btn-nclicks-1" == ctx.triggered_id:
        
        df = pd.read_excel(r'Agents.xlsx')     # Directory needs to be updated

        msg1 = 'You have removed {} from the list of agents'.format(value_name)
        df = df[df["Name"].str.contains(value_name) == False]

        df.to_excel('Agents.xlsx', index = False)  # Directory needs to be updated


    else:
        pass

    return html.Div(msg1)



#################################
##### Change Working State ######
#################################

@app.callback(
    Output('mdd-output-container', 'children'),
    Input('work-list', 'value'),
    prevent_initial_call=False    
    )

def update_working(value1):

    df = pd.read_excel(r'Agents.xlsx')  # Directory needs to be updated
    dt = pd.read_excel(r'Data.xlsx')
    
    # Sets up today's date and yesterdays data
    now = datetime.now()
    timestamp_num = datetime.timestamp(now)
    timestamp = datetime.fromtimestamp(timestamp_num)

    day = timestamp.strftime("%d")
    today = day
    yesterday = dt.loc[0,'Date']
    dt.loc[0, 'Date'] = today
    timestamp_now = pd.to_datetime('today').strftime("%Y-%m-%d %H:%M:%S")


    # checks if time is different day, if true, resets tickets worked to zero
    if int(today) != yesterday:
        
        df['MBM_Worked'] = 0
        df['UET_Worked'] = 0
        df.to_excel('Agents.xlsx', index = False) # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False) # Directory needs to be updated
    else:
        pass




    #Checks if there are any workers selected
    if not value1:
        df['Working'] = False
        msg2 = "There aren't any agents selected, please select an Agent first!"
        
    else:
        df = pd.read_excel(r'Agents.xlsx')  # Directory needs to be updated
        # Updates selected agent values and updates working list. 
        list_value = value1
        df_value1 = pd.Series(list_value)
        working_value = df['Name'].isin(df_value1)
        df['Working'] = working_value
        invalid_workers = False
        msg2 = 'Agents working today has been updated'
    
        df.to_excel('Agents.xlsx', index = False) # Directory needs to be updated
   
    
    return  msg2
    


###########################################
### Manual/Automatic Assign MBM Cases  ####
###########################################


@app.callback(
    #Auto/Mnl MBM assign
    Output('mbm-output-container', 'children'),
    Output('mbm-count-output', 'children'),
    Output('mbm-assignee-output', 'children'),
    Output('mbm-day-count-output', 'children'),



    Input('ambm-btn', 'n_clicks'),
    Input('mmbm-btn', 'n_clicks'),
    Input('mbm_dropdown', component_property='options'),
    State("mbm_dropdown", "value"),  
    prevent_initial_call=False
)

#Manual assign MBM ticket and setups button call for Automatic assign MBM case 
def update_mbm(button1, button2, options, value):

    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')

    #Assigns value of total tickets worked
    mbm_total_output = dt.at[0, 'Total_MBM_Cases']
    message1 = "Either manually or automatically assign agent a ticket"
    assignee = value 

    #Gets sum of total tickets worked for the day
    mbm_day_output = df['MBM_Worked'].sum()

    #Waits for button click to be triggered
    triggered_id = ctx.triggered_id
    
    if triggered_id == 'mmbm-btn':
         
        #Reads/creates dataframes from Excel files
        df = pd.read_excel(r'Agents.xlsx')
        dt = pd.read_excel(r'Data.xlsx')
        
        #Gets agent name from dropdown and finds index value
        selected_name = value                                                      # df['Name'].str.capitalize() )
        mbm_selected_index = df[df['Name']==selected_name].index.values

        #Assigns value of total tickets worked
        mbm_total_output = dt.at[0, 'Total_MBM_Cases']

        assignee = selected_name

        #Gets sum of total tickets worked for the day
        mbm_day_output = df['MBM_Worked'].sum()

        #Checks if selected agent is working today then returns True/False value
        working_true = df.at[int(mbm_selected_index), 'Working']

        #Checks if agent is working today value is False
        if working_true == False:
            
            message1 = "Are you sure you selected the right agent? This agent isn't set to work today!"

            return message1, mbm_total_output, assignee, mbm_day_output
    
        else:
            #Adds +1 ticket worked to selected agent
            df.at[int(mbm_selected_index), 'MBM_Worked'] += 1
            
            #Add +1 to selection count value to all agents
            df['MBM_Selected'] += 1

            #Resets selection value count of selected agent to zero
            df.at[int(mbm_selected_index), 'MBM_Selected'] = 0

            #Add +1 to total worked tickets
            dt['Total_MBM_Cases'] += 1

            #Reads personal agent Excel files            
            da_m = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', sheet_name=selected_name+'_MBM_Worked') ####*******Will need to redefine directory
            da_u = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', sheet_name=selected_name+'_UET_Worked')  ####*******Will need to redefine directory


            #Makes copy of agent data and adds new data
            da_m1 = pd.DataFrame(da_m[[selected_name+'_MBM_Worked', 'Action']])
            da_m2 = pd.DataFrame({selected_name+'_MBM_Worked': [timestamp], "Action": ['Manually Assigned Ticket']})
            da_m = pd.concat([da_m1, da_m2])

            #writes data to excel file
            writer = pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', engine='xlsxwriter')  ####*******Will need to redefine directory
            with pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)
            
            #Assigns value of total tickets worked
            mbm_total_output = dt.at[0, 'Total_MBM_Cases']
            
            #Gets sum of total tickets worked for the day
            mbm_day_output = df['MBM_Worked'].sum()

            #Returns message of changes made
            message1 = '{} was manually assigned the next MBM Case'.format(selected_name, mbm_total_output)
            assignee = selected_name

            #Writes data to Excel files
            df.to_excel('Agents.xlsx', index = False)
            dt.to_excel('Data.xlsx', index = False)

            return message1, mbm_total_output, assignee, mbm_day_output
        


        
    #calls on auto_mbm assign button function if pressed
    elif triggered_id == 'ambm-btn':
         return Auto_MBM()

    return message1, mbm_total_output, assignee, mbm_day_output


#Auto Assign ticket based on a selection value counter. The agent with highest value is selected for next ticket. Then selection value is reset for selected agent
#And all other agents gain +1 to selection value.
def Auto_MBM():

    df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated

    #Checks if any agents are working
    if df['Working'].values.sum() == 0:
        
        #imports files from Excel files
        df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated
        dt = pd.read_excel(r'Data.xlsx')              # Directory needs to be updated
            
        message1 = 'No agents selected to work today!'
        
        #Finds agent with lowest selection count (last agent picked)
        min_work_index = df['MBM_Selected'].idxmin()
        min_agent_name_uet = df.at[min_work_index, 'Name']
        assignee = min_agent_name_uet

        #Gets sum of total tickets worked for the day
        mbm_day_output = df['MBM_Worked'].sum()

        #Assigns value of total tickets worked
        mbm_total_output = dt.at[0, 'Total_MBM_Cases']

        #Saves changes to Excel files
        df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated
        
        return message1, mbm_total_output, assignee, mbm_day_output
        
    else:

        #imports files from Excel files
        df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated
        dt = pd.read_excel(r'Data.xlsx')              # Directory needs to be updated

        #Creates list of agents who are working today
        work_list = df[df['Working'] == True]
        
        #Creates Data Frame of agents who are working today, then finds the agent with highest slection count value
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['MBM_Selected'].idxmax()
        max_agent_name_mbm = df.at[max_work_index, 'Name']
        max_agent_name_column_mbm = max_agent_name_mbm + '_MBM_Cases_Worked'
        
        #Reads and creates Data Frame of personal agent Excel file
        da_m = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+max_agent_name_mbm+'.xlsx', sheet_name=max_agent_name_mbm+'_MBM_Worked')  ####*******Will need to redefine directory
        da_u = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+max_agent_name_mbm+'.xlsx', sheet_name=max_agent_name_mbm+'_UET_Worked') ####*******Will need to redefine directory


        #Adds +1 count to all agent's selection count value
        df['MBM_Selected'] += 1

        #Adds +1 total number of worked tickets
        df.at[max_work_index, 'MBM_Worked'] += 1

        #Gets name of agent with highest selection count value then returns it in message
        assignee = df.at[max_work_index, 'Name']
        message1 = "{} was automatically assigned the next MBM case".format(assignee)

        
        #Resets value of highest selection count agent to 0
        df.at[max_work_index, 'MBM_Selected'] = 0

        #Adds +1 to total ever worked tickets
        dt['Total_MBM_Cases'] += 1

        #Gets sum of total tickets worked for the day
        mbm_day_output = df['MBM_Worked'].sum()
        
        #Assigns value of total tickets worked
        mbm_total_output = dt.at[0, 'Total_MBM_Cases']

        #Makes copy of selected agent data frame, writes timestamp to copy of data frame and then merges them together
        da_m1 = pd.DataFrame(da_m[[max_agent_name_mbm+'_MBM_Worked', 'Action']])
        da_m2 = pd.DataFrame({max_agent_name_mbm+'_MBM_Worked': [timestamp], "Action": ['Automatically Assigned Ticket']})
        da_m = pd.concat([da_m1, da_m2])

        #writes data to agent's personal excel file
        writer = pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+max_agent_name_mbm+'.xlsx', engine='xlsxwriter')  ####*******Will need to redefine directory
        with pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+max_agent_name_mbm+'.xlsx') as writer:
            da_m.to_excel(writer, sheet_name=max_agent_name_mbm+'_MBM_Worked', index=False)
            da_u.to_excel(writer, sheet_name=max_agent_name_mbm+'_UET_Worked', index=False)

        #Saves changes to Excel files
        df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated


        return message1, mbm_total_output, assignee, mbm_day_output



###########################################
### Manual/Automatic Assign UET Ticket ####
###########################################


@app.callback(
    #Auto/Mnl UET assign
    Output('uet-output-container', 'children'),
    Output('uet-count-output', 'children'),
    Output('uet-assignee-output', 'children'),
    Output('uet-day-count-output', 'children'),

    Input('auet-btn', 'n_clicks'),
    Input('muet-btn', 'n_clicks'),
    Input('uet_dropdown', component_property='options'),
    State("uet_dropdown", "value"),  
    prevent_initial_call=False
)

#Manual assign UET ticket and setups button call for Automatic assign UET ticket 
def update_uet(button1, button2, options, value):

    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')

    #Assigns value of total tickets worked
    uet_total_output = dt.at[0, 'Total_UET_Tickets']
    message1 = "Either manually or automatically assign agent a ticket"
    assignee = value 

    #Gets sum of total tickets worked for the day
    uet_day_output = df['UET_Worked'].sum()

    #Waits for button click to be triggered
    triggered_id = ctx.triggered_id

    if triggered_id == 'muet-btn':
         
        #Reads/creates dataframes from Excel files
        df = pd.read_excel(r'Agents.xlsx')
        dt = pd.read_excel(r'Data.xlsx')

        #Gets agent name from dropdown and finds index value
        selected_name = value                                                      # df['Name'].str.capitalize() )
        uet_selected_index = df[df['Name']==selected_name].index.values

        #Assigns value of total tickets worked
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        assignee = selected_name

        #Gets sum of total tickets worked for the day
        uet_day_output = df['UET_Worked'].sum()

        #Checks if selected agent is working today then returns True/False value
        working_true = df.at[int(uet_selected_index), 'Working']

        #Checks if agent is working today value is False
        if working_true == False:
            
            message1 = "Are you sure you selected the right agent? This agent isn't set to work today!"

            return message1, uet_total_output, assignee, uet_day_output
    
        else:
            #Adds +1 ticket worked to selected agent
            df.at[int(uet_selected_index), 'UET_Worked'] += 1
            
            #Add +1 to selection count value to all agents
            df['UET_Selected'] += 1

            #Resets selection value count of selected agent to zero
            df.at[int(uet_selected_index), 'UET_Selected'] = 0

            #Add +1 to total worked tickets
            dt['Total_UET_Tickets'] += 1
            
            #Reads personal agent Excel files            
            da_m = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', sheet_name=selected_name+'_MBM_Worked') ####*******Will need to redefine directory
            da_u = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', sheet_name=selected_name+'_UET_Worked')  ####*******Will need to redefine directory


            #Makes copy of agent data and adds new data
            da_u1 = pd.DataFrame(da_u[[selected_name+'_UET_Worked', 'Action']])
            da_u2 = pd.DataFrame({selected_name+'_UET_Worked': [timestamp], "Action": ['Manually Assigned Ticket']})
            da_u = pd.concat([da_u1, da_u2])

            #writes data to excel file
            writer = pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx', engine='xlsxwriter')  ####*******Will need to redefine directory
            with pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+selected_name+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)
        


            #Assigns value of total tickets worked
            uet_total_output = dt.at[0, 'Total_UET_Tickets']

            #Gets sum of total tickets worked for the day
            uet_day_output = df['UET_Worked'].sum()

            #Returns message of changes made
            message1 = '{} was manually assigned the next UET ticket'.format(selected_name, uet_total_output)
            assignee = selected_name
            
            #Writes data to Excel files
            df.to_excel('Agents.xlsx', index = False)
            dt.to_excel('Data.xlsx', index = False)
        
            return message1, uet_total_output, assignee, uet_day_output

        
    #calls on auto_uet assign button function if pressed
    elif triggered_id == 'auet-btn':
         return Auto_UET()

    return message1, uet_total_output, assignee, uet_day_output


#Auto Assign ticket based on a selection value counter. The agent with highest value is selected for next ticket. Then selection value is reset for selected agent
#And all other agents gain +1 to selection value.
def Auto_UET():

    #imports files from Excel files
    df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated

    #Checks if any agents are working
    if df['Working'].values.sum() == 0:
            
        message1 = 'No agents selected to work today!'
        
        #Finds agent with lowest selection count (last agent picked)
        min_work_index = df['UET_Selected'].idxmin()
        min_agent_name_uet = df.at[min_work_index, 'Name']
        assignee = min_agent_name_uet
        
        #Assigns value of total tickets worked
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        #Gets sum of total tickets worked for the day
        uet_day_output = df['UET_Worked'].sum()
        
        return message1, uet_total_output, assignee, uet_day_output
        
    else:

        #Creates list of agents who are working today
        work_list = df[df['Working'] == True]
        
        #Creates Data Frame of agents who are working today, then finds the agent with highest slection count value
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['UET_Selected'].idxmax()
        max_agent_name_uet = df.at[max_work_index, 'Name']
        max_agent_name_column_uet = max_agent_name_uet + '_UET_Cases_Worked'
        
        #Reads and creates Data Frame of personal agent Excel file
        da_m = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+max_agent_name_uet+'.xlsx', sheet_name=max_agent_name_uet+'_MBM_Worked')  ####*******Will need to redefine directory
        da_u = pd.read_excel('/home/densadam/AREA51/Agent_Data/'+max_agent_name_uet+'.xlsx', sheet_name=max_agent_name_uet+'_UET_Worked') ####*******Will need to redefine directory


        #Adds +1 count to all agent's selection count value
        df['UET_Selected'] += 1

        #Adds +1 total number of worked tickets
        df.at[max_work_index, 'UET_Worked'] += 1

        #Gets name of agent with highest selection count value then returns it in message
        assignee = df.at[max_work_index, 'Name']
        message1 = "{} was automatically assigned the next UET ticket".format(assignee)
        
        #Resets value of highest selection count agent to 0
        df.at[max_work_index, 'UET_Selected'] = 0

        #Adds +1 to total ever worked tickets
        dt['Total_UET_Tickets'] += 1

        #Makes copy of selected agent data frame, writes timestamp to copy of data frame and then merges them together
        da_u1 = pd.DataFrame(da_u[[max_agent_name_uet+'_UET_Worked', 'Action']])
        da_u2 = pd.DataFrame({max_agent_name_uet+'_UET_Worked': [timestamp], "Action": ['Automatically Assigned Ticket']})
        da_u = pd.concat([da_u1, da_u2])

        #writes data to agent's personal excel file
        writer = pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+max_agent_name_uet+'.xlsx', engine='xlsxwriter')  ####*******Will need to redefine directory
        with pd.ExcelWriter('/home/densadam/AREA51/Agent_Data/'+max_agent_name_uet+'.xlsx') as writer:
            da_m.to_excel(writer, sheet_name=max_agent_name_uet+'_MBM_Worked', index=False)
            da_u.to_excel(writer, sheet_name=max_agent_name_uet+'_UET_Worked', index=False)

        #Assigns value of total tickets worked
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        #Gets sum of total tickets worked for the day
        uet_day_output = df['UET_Worked'].sum()

        #Saves changes to Excel files
        df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated

        return message1, uet_total_output, assignee, uet_day_output




# Run the app on localhost:8050
if __name__ == '__main__':
    app.run_server(debug=True)



