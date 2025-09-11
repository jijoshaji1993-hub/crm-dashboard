import pandas as pd
from dash import Dash, dcc, html, Input, Output
import plotly.express as px

# === Load Excel data ===
file_path = "CRM_Analysis_Report.xlsx"
xls = pd.ExcelFile(file_path)

daywise_report = pd.read_excel(xls, sheet_name='Daywise_Report')
wrong_dockets = pd.read_excel(xls, sheet_name='Wrong_Dockets')
reassigned_complaints = pd.read_excel(xls, sheet_name='Reassigned_Complaints')
breakdown = pd.read_excel(xls, sheet_name='Complaint_Breakdown')
repeat_summary = pd.read_excel(xls, sheet_name='Repeat_Call_Report')
wrong_complaints = pd.read_excel(xls, sheet_name='Wrong_Complaints')
invalid_recharge_tagging = pd.read_excel(xls, sheet_name='Invalid_Recharge_Tagging')

# === Setup Dash app ===
app = Dash(__name__)
app.title = "CRM Analysis Dashboard"

app.layout = html.Div([
    html.H1("CRM Analysis Dashboard", style={'textAlign': 'center'}),

    dcc.Tabs(id="tabs", value='tab-daywise', children=[
        dcc.Tab(label='Daywise Summary', value='tab-daywise'),
        dcc.Tab(label='Complaint Breakdown', value='tab-breakdown'),
        dcc.Tab(label='Repeat Calls', value='tab-repeat'),
        dcc.Tab(label='Wrong Dockets', value='tab-wrong-dockets'),
        dcc.Tab(label='Wrong Complaints', value='tab-wrong-complaints'),
        dcc.Tab(label='Invalid Recharge Tagging', value='tab-invalid-recharge')
    ]),

    html.Div(id='tabs-content', style={'padding': '20px'})
])

# === Tab rendering callback ===
@app.callback(
    Output('tabs-content', 'children'),
    Input('tabs', 'value')
)
def render_tab_content(tab):

    if tab == 'tab-daywise':
        daily_sum = daywise_report.groupby('createdOn')['docketCount'].sum().reset_index()
        fig = px.bar(daily_sum, x='createdOn', y='docketCount', title="📅 Daywise Total Dockets",
                     labels={'createdOn': 'Date', 'docketCount': 'Docket Count'})
        fig.update_traces(texttemplate='%{y}', textposition='outside')

        return html.Div([
            dcc.Graph(figure=fig),
            html.H4("Daywise Docket Summary"),
            create_data_table(daily_sum)
        ])

    elif tab == 'tab-breakdown':
        idx_header = breakdown[breakdown['Date'] == "Complaint Breakdown by Category"].index[0]
        next_headers = breakdown[breakdown['Date'].str.contains("Complaint Breakdown by", na=False) & (breakdown.index > idx_header)].index.tolist()
        end_idx = next_headers[0] if next_headers else len(breakdown)
        category_data = breakdown.iloc[idx_header+1:end_idx].dropna(subset=['Category', 'count'])

        fig = px.bar(category_data, x='Category', y='count', title="📊 Complaint Breakdown by Category",
                     labels={'Category': 'Category', 'count': 'Count'})
        fig.update_traces(texttemplate='%{y}', textposition='outside')
        fig.update_layout(xaxis_tickangle=45)

        return html.Div([
            dcc.Graph(figure=fig),
            html.H4("Complaint Breakdown - Category Level"),
            create_data_table(category_data)
        ])

    elif tab == 'tab-repeat':
        fig = px.line(repeat_summary.sort_values('Date'), x='Date',
                      y=['First call complaint', 'Repeat call complaint'],
                      title="📈 First vs Repeat Call Complaints Over Time")
        fig.update_traces(mode='lines+markers')

        return html.Div([
            dcc.Graph(figure=fig),
            html.H4("Repeat Call Report"),
            create_data_table(repeat_summary)
        ])

    elif tab == 'tab-wrong-dockets':
        return html.Div([
            html.H4("❗ Wrong Dockets"),
            create_data_table(wrong_dockets)
        ])

    elif tab == 'tab-wrong-complaints':
        top_agents = wrong_complaints['createdBy'].value_counts().reset_index()
        top_agents.columns = ['Agent', 'Wrong Complaints']
        fig = px.bar(top_agents.head(10), x='Agent', y='Wrong Complaints',
                     title="🚨 Top Contributors - Wrong Complaints")
        fig.update_traces(texttemplate='%{y}', textposition='outside')

        return html.Div([
            dcc.Graph(figure=fig),
            html.H4("Top Agents - Wrong Complaints"),
            create_data_table(top_agents)
        ])

    elif tab == 'tab-invalid-recharge':
        top_agents = invalid_recharge_tagging['createdBy'].value_counts().reset_index()
        top_agents.columns = ['Agent', 'Invalid Taggings']
        fig = px.bar(top_agents.head(10), x='Agent', y='Invalid Taggings',
                     title="⚠️ Top Contributors - Invalid Recharge Tagging")
        fig.update_traces(texttemplate='%{y}', textposition='outside')

        return html.Div([
            dcc.Graph(figure=fig),
            html.H4("Invalid Recharge Tagging Summary"),
            create_data_table(top_agents)
        ])

    else:
        return html.Div([
            html.H3("Tab not found")
        ])


# === Utility: Data table generator ===
def create_data_table(df):
    return html.Table([
        html.Thead(html.Tr([html.Th(col) for col in df.columns])),
        html.Tbody([
            html.Tr([html.Td(df.iloc[i][col]) for col in df.columns])
            for i in range(len(df))
        ])
    ], style={'width': '100%', 'border': '1px solid gray', 'borderCollapse': 'collapse'})


# === Run app ===
if __name__ == '__main__':
    app.run(debug=True, port=8050)
