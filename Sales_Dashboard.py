import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px

# Load dataset
df = pd.read_excel(r"C:\Users\6571kb\OneDrive - BP\Personel\IUM\Data Test - Umar Rosli.xlsx", sheet_name="Raw")

# Add date columns
df['Year'] = pd.to_datetime(df['DocumentDate']).dt.year
df['Quarter'] = pd.to_datetime(df['DocumentDate']).dt.quarter
df['Month'] = pd.to_datetime(df['DocumentDate']).dt.month
df['Date1'] = df['Year'].astype(str) + '-Q' + df['Quarter'].astype(str)
df['Date2'] = df['Year'].astype(str) + '-' + df['Month'].astype(str).str.zfill(2)
df['Year'] = df['Year'].astype(int)

# Initialize Dash app
app = dash.Dash(__name__)

# Dashboard layout
app.layout = html.Div([
    # Main title
    html.H1('IUM Sales Dashboard',style={'font-family': 'Arial, sans-serif', 'font-size': '36px','color': '#007BFF', 'text-align': 'center'}),
    html.Div('Overview of Sales Insights of IUM - a dashboard made by Umar Rosli',style={'text-align': 'center'}),
    html.Br(),
    html.Br(),
    # Dropdown filter
    html.Div([
        html.Label("Select Region:"),
        dcc.Dropdown(
            id="region-filter",
            options=[{"label": str(region), "value": region} for region in sorted(df['Region'].unique())],
            value=sorted(df['Region'].unique()),  # Default to all regions
            clearable=False,
            multi=True, # Allow multiple selection
            style={'width': '300px'}  # Set dropdown width
        )
    ], style={'display': 'inline-block', 'margin': '10px'}),

    html.Br(),

    # KPI Cards
    html.Div([
        html.Div([
            html.H3("Total Sales", style={'text-align': 'center','margin': '0'}),
            html.H4(id="total-sales", style={'text-align': 'center', 'color': '#28a745','margin': '0','font-size': '30px'})
        ], style={'width': '20%', 'display': 'inline-block', 'padding': '10px', 'border': '1px solid #ccc', 'border-radius': '5px', 'box-shadow': '0px 0px 10px rgba(0, 0, 0, 0.1)', 'margin-right': '1%'}),
        
        html.Div([
            html.H3("Total Carton Qty", style={'text-align': 'center','margin': '0'}),
            html.H4(id="total-carton", style={'text-align': 'center', 'color': '#17a2b8','margin': '0','font-size': '30px'})
        ], style={'width': '20%', 'display': 'inline-block', 'padding': '10px', 'border': '1px solid #ccc', 'border-radius': '5px', 'box-shadow': '0px 0px 10px rgba(0, 0, 0, 0.1)', 'margin-right': '1%'}),
        
        html.Div([
            html.H3("# of Customers", style={'text-align': 'center','margin': '0'}),
            html.H4(id="distinct-customers", style={'text-align': 'center', 'color': '#ffc107','margin': '0','font-size': '30px'})
        ], style={'width': '20%', 'display': 'inline-block', 'padding': '10px', 'border': '1px solid #ccc', 'border-radius': '5px', 'box-shadow': '0px 0px 10px rgba(0, 0, 0, 0.1)', 'margin-right': '1%'}),
        
        html.Div([
            html.H3("# of SKUs", style={'text-align': 'center','margin': '0'}),
            html.H4(id="distinct-skus", style={'text-align': 'center', 'color': '#dc3545','margin': '0','font-size': '30px'})
        ], style={'width': '20%', 'display': 'inline-block', 'padding': '10px', 'border': '1px solid #ccc', 'border-radius': '5px', 'box-shadow': '0px 0px 10px rgba(0, 0, 0, 0.1)'})
    ], style={'display': 'flex', 'justify-content': 'center', 'align-items': 'center', 'margin-bottom': '20px', 'flex-wrap': 'wrap'}),


    # Yearly and Quarterly Sales graphs
    html.Div([
            html.H2("Sales Trend"),
        html.Div([
            dcc.Graph(id="yearly-bar-graph")
        ], style={'width': '40%', 'display': 'inline-block'}),
        html.Div([
            dcc.Graph(id="quarterly-bar-graph")
        ], style={'width': '60%', 'display': 'inline-block'})
    ]),

    # Monthly Bar graph
    html.Div([
        dcc.Graph(id="bar-graph")
    ])
])

# App callback
@app.callback(
    Output("total-sales", "children"),
    Output("total-carton", "children"),
    Output("distinct-customers", "children"),
    Output("distinct-skus", "children"),
    Output("yearly-bar-graph", "figure"),
    Output("quarterly-bar-graph", "figure"),
    Output("bar-graph", "figure"),
    Input("region-filter", "value")
)
def update_graph(selected_region):
    # Filter data based on the selected region(s)
    filtered_data = df[df['Region'].isin(selected_region)]

    # KPI metrics
    total_sales = filtered_data['TTL Net Amount'].sum()
    total_carton = filtered_data['CartonQty'].sum()
    distinct_customers = filtered_data['CustomerCode'].nunique()
    distinct_skus = filtered_data['StockCode'].nunique()

    # Format KPI values
    total_sales_str = f"${total_sales / 1_000_000:.2f}M"
    total_carton_str = f"{total_carton / 1_000_000:.2f}M"
    distinct_customers_str = f"{distinct_customers:,}"
    distinct_skus_str = f"{distinct_skus:,}"

    # agg tables of data before creating graph
    ## Monthly sales data
    monthly_data = filtered_data.groupby(['Date2', 'Region'])['TTL Net Amount'].sum().reset_index()

    ## Quarterly sales data
    quarterly_data = filtered_data.groupby(['Date1', 'Region'])['TTL Net Amount'].sum().reset_index()

    ## Yearly sales data
    yearly_data = filtered_data.groupby(['Year', 'Region'])['TTL Net Amount'].sum().reset_index()


    # Calculate totals for annotations
    monthly_totals = monthly_data.groupby('Date2')['TTL Net Amount'].sum().reset_index()
    quarterly_totals = quarterly_data.groupby('Date1')['TTL Net Amount'].sum().reset_index()
    yearly_totals = yearly_data.groupby('Year')['TTL Net Amount'].sum().reset_index()


    # Create monthly bar chart
    monthly_fig = px.bar(
        monthly_data,
        x="Date2",
        y="TTL Net Amount",
        color="Region",
        barmode="stack",
        title="Monthly Sales"
    )
    monthly_fig.update_yaxes(title="Sales in Millions", tickprefix="$", showgrid=True)
    monthly_fig.update_xaxes(title="Month")

    for row in monthly_totals.itertuples():
        monthly_fig.add_annotation(
            x=row.Date2,
            y=row._2,
            text=f"${row._2 / 1_000_000:.2f}M",
            showarrow=False,
            font=dict(size=9, color="black"),
            yshift=10
        )

    # Create quarterly sales bar chart
    quarterly_fig = px.bar(
        quarterly_data,
        x="Date1",
        y="TTL Net Amount",
        color="Region",
        barmode="stack",
        title="Quarterly Sales"
    )
    quarterly_fig.update_yaxes(title="Sales in Millions", tickprefix="$", showgrid=True)
    quarterly_fig.update_xaxes(title="Quarter")

    for row in quarterly_totals.itertuples():
        quarterly_fig.add_annotation(
            x=row.Date1,
            y=row._2,
            text=f"${row._2 / 1_000_000:.2f}M",
            showarrow=False,
            font=dict(size=11, color="black"),
            yshift=10
        )

    # Create yearly sales horizontal bar chart
    yearly_fig = px.bar(
        yearly_data,
        x="TTL Net Amount",
        y="Year",
        color="Region",
        orientation='h',
        barmode="stack",
        title="Yearly Sales"
    )
    yearly_fig.update_xaxes(title="Sales in Millions", tickprefix="$", showgrid=True)
    yearly_fig.update_yaxes(title="Year",tickformat="d") 

    for row in yearly_totals.itertuples():
        yearly_fig.add_annotation(
            x=row._2,
            y=row.Year,
            text=f"${row._2 / 1_000_000:.2f}M",
            showarrow=False,
            font=dict(size=12, color="black"),
            xshift=10
        )

    return total_sales_str, total_carton_str, distinct_customers_str, distinct_skus_str, yearly_fig, quarterly_fig, monthly_fig

# Run app
if __name__ == "__main__":
    app.run_server(debug=True)
