import dash
from dash import html, dcc
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import ThemeChangerAIO
import SBW

# stylesheet with the .dbc class to style dcc components with a Bootstrap theme
dbc_css = "https://cdn.jsdelivr.net/gh/AnnMarieW/dash-bootstrap-templates/dbc.min.css"

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME, dbc_css])

color_mode_switch = html.Span(
    [
        dbc.Label(className="fa fa-moon", html_for="switch"),
        dbc.Switch(id="switch", value=True, className="d-inline-block ms-1", persistence=True),
        dbc.Label(className="fa fa-sun", html_for="switch"),
    ]
)

theme_controls = html.Div(
    [ThemeChangerAIO(aio_id="theme"), color_mode_switch],
    className="hstack gap-3 mt-2"
)

header = html.H1("Arbetsmiljö - Hjälpmedel", className="bg-primary text-white p-2 mb-2 text-center")

app.layout = dbc.Container(
    [
        header,
        dbc.Row([
            dbc.Col([
                theme_controls,
            ], width=12),
        ]),
        dbc.Row([
            dbc.Col([
                dcc.Tabs([
                    dcc.Tab(label="Risk och Konsekvensanalys", children=[
                        SBW.layout
                    ]),
                    dcc.Tab(label="Funktion under uppbyggnad", children=[
                        html.Div("Funktion under uppbyggnad")
                    ])
                ])
            ], width=12),
        ]),
    ],
    fluid=True,
    className="dbc",
)

# Register SBW callbacks
SBW.register_callbacks(app)

# updates the Bootstrap global light/dark color mode
app.clientside_callback(
    """
    function(switchOn) {       
       document.documentElement.setAttribute('data-bs-theme', switchOn ? 'light' : 'dark');  
       return window.dash_clientside.no_update
    }
    """,
    dash.Output("switch", "id"),
    dash.Input("switch", "value"),
)

if __name__ == "__main__":
    app.run_server(debug=True)
