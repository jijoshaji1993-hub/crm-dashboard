import os
import pandas as pd
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
from dash import dash_table
import plotly.express as px

# === Load all sheets dynamically ===
file_path = "CRM_Analysis_Report.xlsx"
xls = pd.ExcelFile(file_path)
sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

# === Initialize app ===
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
app.title = "CRM Analysis Dashboard"

# === Color scheme ===
COLORS = {
    'header': '#2C3E50',
    'primary': '#1F77B4',
    'secondary': '#6C757D',
    'accent': '#FF7F0E',
    'background': '#d8d4d8',  # Updated background color
    'table_header_bg': '#1F77B4',
    'table_header_color': 'white'
}

# === Helper: Pretty sheet names ===
def format_sheet_name(sheet_name):
    return sheet_name.replace("_", " ").title()

# === Helper: DataTable ===
def make_table(df, table_id, page_size=10):
    return dash_table.DataTable(
        id=table_id,
        columns=[{"name": col, "id": col} for col in df.columns],
        data=df.to_dict('records'),
        page_size=page_size,
        page_action='native',
        filter_action='native',
        sort_action='native',
        sort_mode='multi',
        style_table={'overflowX': 'auto'},
        style_header={
            'backgroundColor': COLORS['table_header_bg'],
            'color': COLORS['table_header_color'],
            'fontWeight': 'bold'
        },
        style_cell={'padding': '5px', 'textAlign': 'left', 'minWidth': '100px', 'whiteSpace': 'normal'},
        style_data_conditional=[
            {'if': {'row_index': 'even'}, 'backgroundColor': '#f2f2f2'},
            {'if': {'row_index': 'odd'}, 'backgroundColor': 'white'}
        ]
    )

# === Navigation buttons ===
def nav_buttons():
    return dbc.ButtonGroup([
        dbc.Button("🏠 Home", id='btn_home', color='info', href='/'),
        dbc.Button("🔙 Back", id='btn_back', color='secondary', n_clicks=0),
    ], className="mb-3")

# === Generic sheet layout ===
def layout_generic(sheet_name, prev_pathname=None):
    df = sheets[sheet_name]
    table = make_table(df, table_id=f"table_{sheet_name.lower()}", page_size=10)
    return dbc.Container([
        nav_buttons(),
        html.H2(format_sheet_name(sheet_name), className="text-primary"),
        dbc.Button("Download CSV", id=f"download_{sheet_name.lower()}", color='success', className='my-2'),
        dcc.Download(id=f"download_{sheet_name.lower()}_csv"),
        table
    ], className="p-4")

# === Home Page with KPIs and Background Color ===
def layout_home():
    # KPI Cards
    kpi_cards = []
    total_dockets = sheets['Daywise_Report']['docketCount'].sum() if 'Daywise_Report' in sheets else 0
    total_complaints = sheets['Complaint_Breakdown']['count'].sum() if 'Complaint_Breakdown' in sheets else 0
    total_wrong = len(sheets['Wrong_Complaints']) if 'Wrong_Complaints' in sheets else 0
    invalid_recharge_tagging = len(sheets['Invalid_Recharge_Tagging']) if 'Invalid_Recharge_Tagging' in sheets else 0
    reassigned_complaints = len(sheets['Reassigned_Complaints']) if 'Reassigned_Complaints' in sheets else 0
    wrong_dockets = len(sheets['Wrong_Dockets']) if 'Wrong_Dockets' in sheets else 0

    kpis = [
        ("Total Dockets", total_dockets, COLORS['primary']),
        ("Total Complaints", total_complaints, COLORS['accent']),
        ("Wrong Complaints", total_wrong, COLORS['header']),
        ("Invalid Recharge Tagging", invalid_recharge_tagging, COLORS['secondary']),
        ("Reassigned Complaints", reassigned_complaints, COLORS['primary']),
        ("Wrong Dockets", wrong_dockets, COLORS['accent'])
    ]

    for title, value, color in kpis:
        kpi_cards.append(
            dbc.Card([
                dbc.CardBody([
                    html.H6(title, className="text-muted", style={"fontSize": "0.8rem"}),
                    html.H4(f"{value:,}", className="fw-bold", style={"color": color, "fontSize": "1.4rem"})
                ])
            ], className="m-2", style={'width': '8rem', 'textAlign': 'Side'})
        )

    # Date range picker
    min_date = sheets['Daywise_Report']['createdOn'].min() if 'Daywise_Report' in sheets else None
    max_date = sheets['Daywise_Report']['createdOn'].max() if 'Daywise_Report' in sheets else None
    date_picker = dcc.DatePickerRange(
        id='home_date_picker',
        start_date=min_date,
        end_date=max_date,
        min_date_allowed=min_date,
        max_date_allowed=max_date,
        display_format='YYYY-MM-DD',
        style={'marginBottom': '20px'}
    )

    chart_graph = dcc.Graph(id='home_daywise_chart')

    # Sheet navigation cards
    sheet_cards = []
    for sheet in sheets.keys():
        path = f"/{sheet.lower().replace(' ', '_')}"
        sheet_cards.append(
            dbc.Card([
                dbc.CardBody([
                    html.H5(format_sheet_name(sheet), className="card-title"),
                    dbc.Button("View", color="primary", href=path)
                ])
            ], className="m-2", style={'width': '18rem'})
        )

    return dbc.Container([
        html.H1("CRM Analysis Dashboard", className="text-center text-primary my-4"),
        dbc.Row([dbc.Col(card, width="auto") for card in kpi_cards], justify="center"),
        html.H3("Explore Reports", className="text-center text-secondary my-4"),
        dbc.Row([dbc.Col(card, width="auto") for card in sheet_cards], justify="center"),
        dbc.Row(dbc.Col(date_picker, width=6), justify="center"),
        dbc.Row(dbc.Col(chart_graph, width=12), className="my-4")
    ], className="p-4", style={"backgroundColor": COLORS["background"], "minHeight": "100vh"})

# === Callback for Home Page Daywise Chart ===
@app.callback(
    Output('home_daywise_chart', 'figure'),
    Input('home_date_picker', 'start_date'),
    Input('home_date_picker', 'end_date')
)
def update_home_chart(start_date, end_date):
    df = sheets.get('Daywise_Report', pd.DataFrame())
    if df.empty or start_date is None or end_date is None:
        return px.line(title="No Daywise Data")
    df_filtered = df[(df['createdOn'] >= start_date) & (df['createdOn'] <= end_date)]
    daily_sum = df_filtered.groupby('createdOn')['docketCount'].sum().reset_index()
    fig = px.line(daily_sum, x='createdOn', y='docketCount', title="Daywise Dockets Trend", template='plotly_white')
    return fig

# === Page routing and callbacks ===
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content'),
    dcc.Store(id='prev-pathname', data='/')
])

@app.callback(
    Output('prev-pathname', 'data'),
    Input('url', 'pathname'),
    State('prev-pathname', 'data')
)
def update_prev_path(pathname, prev_path):
    return prev_path if pathname == prev_path else pathname

@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname')
)
def display_page(pathname):
    if pathname == "/":
        return layout_home()
    sheet_key = pathname.strip("/").replace("_", " ")
    for sheet in sheets.keys():
        if sheet.lower().replace("_", " ") == sheet_key.lower():
            return layout_generic(sheet, prev_pathname="/")
    return dbc.Container([
        nav_buttons(),
        html.H2("404: Page Not Found", className="text-danger"),
        html.P(f"The requested page {pathname} was not recognised.")
    ], className="p-4")

@app.callback(
    Output('url', 'pathname'),
    Input('btn_back', 'n_clicks'),
    State('prev-pathname', 'data'),
    prevent_initial_call=True
)
def go_back(n_clicks, prev_path):
    return prev_path if prev_path else "/"

# === Auto download callbacks ===
for sheet in sheets.keys():
    @app.callback(
        Output(f"download_{sheet.lower()}_csv", "data"),
        Input(f"download_{sheet.lower()}", "n_clicks"),
        prevent_initial_call=True
    )
    def download_csv(n_clicks, sheet=sheet):
        df = sheets[sheet]
        return dcc.send_data_frame(df.to_csv, f"{sheet}.csv", index=False)

# === Run server ===
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=True)
