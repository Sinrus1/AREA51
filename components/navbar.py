# Import necessary libraries
from dash import html
import dash_bootstrap_components as dbc


# Define the navbar structure
def Navbar():

    layout = html.Div([
        dbc.NavbarSimple(
            children=[
                dbc.NavItem(dbc.NavLink("Tracker", href="/Tracker")),
                dbc.NavItem(dbc.NavLink("Add/Remove Agent", href="/Add_Remove_Agent")),
            
            ] ,
            brand="Agent Rotation Errand Assiger v5.1 (AREA 51)",
            brand_href="/Tracker",
            color="green",
            
            dark=True,
        ), 
    ])

    return layout
