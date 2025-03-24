# Importar las bibliotecas necesarias
import pandas as pd
from dash import Dash, html, dash_table, dcc, Input, Output, State
import plotly.express as px
import dash_bootstrap_components as dbc
import logging
from datetime import datetime

# Configurar logging para poder ver errores en Render
logging.basicConfig(level=logging.INFO)

# Crear la aplicación Dash con un tema de Bootstrap
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Importante: Exponer el servidor WSGI para Gunicorn
server = app.server

# =============================================
# Carga y preparación de datos
# =============================================

# Cargar los archivos Excel desde las rutas locales
file_name_organizations = "detail-organizations-2025-03-24.xlsx"
file_name_subscriptions = "detail-subscription-2025-03-24.xlsx"
logging.info(f"Archivos cargados: {file_name_organizations}, {file_name_subscriptions}")

# Cargar los datos en DataFrames
try:
    df_organizations = pd.read_excel(file_name_organizations, sheet_name="Organizations")
    df_subscriptions = pd.read_excel(file_name_subscriptions)
    logging.info("DataFrames cargados correctamente")
except Exception as e:
    logging.error(f"Error al cargar archivos: {str(e)}")
    df_organizations = pd.DataFrame()
    df_subscriptions = pd.DataFrame()

# Normalizar los valores de 'status' en el primer archivo
df_organizations['status'] = df_organizations['status'].str.strip().str.capitalize()

# Filtrar los datos del segundo archivo
df_subscriptions_filtrado = df_subscriptions[(df_subscriptions['status'] == 'active') & (df_subscriptions['company'].isna())]

# Calcular empresas activas, pendientes y suspendidas
empresas_activas = df_organizations[df_organizations['status'] == 'Active'].shape[0]
empresas_pendientes = df_organizations[df_organizations['status'] == 'Pending'].shape[0]
empresas_suspendidas = df_organizations[df_organizations['status'] == 'Suspended'].shape[0]

# Análisis por status (active/pending/suspended)
status_summary = df_organizations['status'].value_counts().reset_index()
status_summary.columns = ['Status', 'Cantidad']

# Análisis por KAM (owner)
kam_status_summary = df_organizations.groupby(['owner', 'status']).size().unstack()
kam_status_summary = kam_status_summary.fillna(0)
kam_status_summary = kam_status_summary.reindex(columns=['Active', 'Pending', 'Suspended'], fill_value=0)
kam_status_summary['Total'] = kam_status_summary.sum(axis=1)
kam_status_summary = kam_status_summary.reset_index()

# Crear el resumen por owner y país
try:
    resumen_owner_pais = pd.DataFrame()
    owners = df_organizations['owner'].dropna().unique()
    
    for owner in owners:
        df_owner = df_organizations[df_organizations['owner'] == owner]
        countries = df_owner['country'].dropna().unique()
        
        for country in countries:
            df_filtered = df_owner[df_owner['country'] == country]
            
            active = df_filtered[df_filtered['status'] == 'Active'].shape[0]
            pending = df_filtered[df_filtered['status'] == 'Pending'].shape[0]
            suspended = df_filtered[df_filtered['status'] == 'Suspended'].shape[0]
            
            total = active + pending + suspended
            avance = round((active / total) * 100, 2) if total > 0 else 0
            
            nueva_fila = pd.DataFrame({
                'owner': [owner],
                'country': [country],
                'Active': [active],
                'Pending': [pending],
                'Suspended': [suspended],
                'Total': [total],
                'Avance': [avance]
            })
            
            resumen_owner_pais = pd.concat([resumen_owner_pais, nueva_fila], ignore_index=True)
    
    resumen_owner_pais['Active'] = resumen_owner_pais['Active'].astype(int)
    resumen_owner_pais['Pending'] = resumen_owner_pais['Pending'].astype(int)
    resumen_owner_pais['Suspended'] = resumen_owner_pais['Suspended'].astype(int)
    resumen_owner_pais['Total'] = resumen_owner_pais['Total'].astype(int)
    resumen_owner_pais['Avance'] = resumen_owner_pais['Avance'].astype(float)
    
    logging.info("Resumen por Owner y País creado correctamente")
    
except Exception as e:
    logging.error(f"Error al crear resumen por owner y país: {str(e)}")
    resumen_owner_pais = pd.DataFrame({
        'owner': ['Error'], 'country': ['Error'], 'Active': [0], 'Pending': [0],
        'Suspended': [0], 'Total': [0], 'Avance': [0.0]
    })

# =============================================
# Preparación de datos para el Marketplace
# =============================================

def prepare_marketplace_data(df):
    """Prepara los datos para el dashboard del Marketplace"""
    try:
        df['Date Creation Order'] = pd.to_datetime(df['Date Creation Order'], format='%d-%m-%Y')
        df['TCV Item'] = (
            df['TCV Item']
            .astype(str)
            .str.replace('[^\d.,-]', '', regex=True)
            .str.replace(',', '.')
            .astype(float)
        )
        df['Order Type'] = df['Order created by'].apply(
            lambda x: 'Orion Hub' if isinstance(x, str) and x.endswith('@orion.global') else 'Market'
        )
        return df
    except Exception as e:
        logging.error(f"Error al preparar datos del Marketplace: {str(e)}")
        return pd.DataFrame()

# Cargar y procesar datos del Marketplace
try:
    df_orders = pd.read_excel("detail-order-2025-01-01-to-2025-03-24.xlsx")  
    df_orders = prepare_marketplace_data(df_orders.copy())
    
    start_date = datetime(2025, 1, 1)
    end_date = datetime(2025, 3, 24)
    market_orders = df_orders[
        (df_orders['Date Creation Order'] >= start_date) & 
        (df_orders['Date Creation Order'] <= end_date) & 
        (df_orders['Order Type'] == 'Market')
    ].copy()

    # Calcular métricas del Marketplace
    metrics_marketplace = {
        'unique_orders': market_orders['Order id'].nunique(),
        'unique_companies': market_orders['Organization'].nunique(),
        'total_amount': market_orders['TCV Item'].sum(),
        'amount_by_company': market_orders.groupby('Organization')['TCV Item'].sum().reset_index()
    }
    
except Exception as e:
    logging.error(f"Error en datos del Marketplace: {str(e)}")
    market_orders = pd.DataFrame()
    metrics_marketplace = {
        'unique_orders': 0,
        'unique_companies': 0,
        'total_amount': 0,
        'amount_by_company': pd.DataFrame(columns=['Organization', 'TCV Item'])
    }

# =============================================
# Definición de los layouts para cada pestaña
# =============================================

def create_dashboard_empresas():
    """Dashboard de Empresas"""
    return dbc.Container([
        dbc.Row(dbc.Col(html.H1("Dashboard de Empresas", className="text-center my-4"))),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col(dcc.Graph(
                            figure=px.pie(status_summary, values='Cantidad', names='Status', 
                                        title="Distribución de Empresas por Estado")
                        ), width=6),
                        dbc.Col([
                            html.H2(f"Empresas Activas: {empresas_activas}", className="text-center"),
                            html.H2(f"Empresas Pendientes: {empresas_pendientes}", className="text-center"),
                            html.H2(f"Empresas Suspendidas: {empresas_suspendidas}", className="text-center")
                        ], width=6, className="d-flex flex-column justify-content-center")
                    ])
                ])
            ], className="mb-4"), width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("Empresas sin Segmento", className="text-center"),
                html.H3(f"Número de Empresas sin Segmento Asignado: {df_organizations['segment'].isna().sum()}", className="text-center"),
                
                dbc.Card([
                    dbc.CardBody([
                        dash_table.DataTable(
                            columns=[{"name": i, "id": i} for i in df_organizations[df_organizations['segment'].isna()][['owner', 'name']].columns],
                            data=df_organizations[df_organizations['segment'].isna()][['owner', 'name']].to_dict('records'),
                            style_table={'height': '300px', 'overflowY': 'auto'},
                            style_cell={'textAlign': 'left', 'padding': '10px'},
                            page_size=10
                        ),
                        html.Br(),
                        dbc.Button("Descargar Empresas sin Segmento", id="btn-descargar-empresas", color="primary"),
                        dcc.Download(id="descargar-empresas")
                    ])
                ], className="mb-4"),
                
                dbc.Card([
                    dbc.CardBody(dcc.Graph(
                        figure=px.bar(
                            df_organizations[df_organizations['segment'].isna()]['owner'].value_counts().reset_index().rename(columns={'index': 'owner', 0: 'count'}),
                            x='owner', y='count',
                            title="Empresas sin Segmento por Propietario",
                            labels={'owner': 'Propietario', 'count': 'Cantidad'}
                        )
                    ))
                ], className="mb-4")
            ], width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("Resumen por Owner y País", className="text-center"),
                dbc.Card([
                    dbc.CardBody([
                        dash_table.DataTable(
                            columns=[{"name": i, "id": i} for i in resumen_owner_pais.columns],
                            data=resumen_owner_pais.to_dict('records'),
                            style_table={'height': '300px', 'overflowY': 'auto'},
                            style_cell={'textAlign': 'left', 'padding': '10px'},
                            page_size=10
                        ),
                        html.Br(),
                        dbc.Button("Descargar Resumen", id="btn-descargar-resumen", color="primary"),
                        dcc.Download(id="descargar-resumen")
                    ])
                ], className="mb-4")
            ], width=12)
        ])
    ], fluid=True)

def create_dashboard_subscripciones():
    """Dashboard de Subscripciones"""
    return dbc.Container([
        dbc.Row(dbc.Col(html.H1("Dashboard de Subscripciones", className="text-center my-4"))),
        
        dbc.Row([
            dbc.Col([
                html.H2("Subscripciones Activas sin Compañía", className="text-center"),
                html.H3(f"Registros Filtrados: {len(df_subscriptions_filtrado)}", className="text-center"),
                
                dbc.Tabs([
    # Pestaña 1: Tabla de subscripciones
    dbc.Tab([
        dbc.Card([
            dbc.CardBody([
                dash_table.DataTable(
                    columns=[{"name": i, "id": i} for i in df_subscriptions_filtrado.columns],
                    data=df_subscriptions_filtrado.to_dict('records'),
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={'textAlign': 'left', 'padding': '10px'},
                    page_size=10
                ),
                html.Br(),
                dbc.Button("Descargar Subscripciones Filtradas", id="btn-descargar-subscripciones", color="primary"),
                dcc.Download(id="descargar-subscripciones")
            ])
        ], className="mb-4")
    ], label="Tabla de Subscripciones"),

    # Pestaña 2: Distribución por dominio
    dbc.Tab([
        dbc.Card([
            dbc.CardBody(
                dcc.Graph(
                    figure=px.bar(
                        df_subscriptions_filtrado['console_domain'].value_counts().reset_index().rename(columns={'index': 'console_domain', 0: 'count'}),
                        x='console_domain', y='count',
                        title="Distribución por dominio",
                        labels={'console_domain': 'dominio', 'count': 'Cantidad'}
                    )
                )
            )
        ])
    ], label="Distribución por Dominio"),

    # Pestaña 3: Distribución por producto
    dbc.Tab([
        dbc.Card([
            dbc.CardBody(
                dcc.Graph(
                    figure=px.pie(df_subscriptions_filtrado, names='product', title="Distribución por Producto")
                )
            )
        ])
    ], label="Distribución por Producto")
])
            ], width=12)
        ])
    ], fluid=True)

def create_marketplace_dashboard():
    """Dashboard del Marketplace"""
    return dbc.Container([
        dbc.Row(dbc.Col(html.H1("Análisis del Marketplace", className="text-center my-4"))),
        
        dbc.Row(dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Período Analizado", className="card-title"),
                html.P(f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", className="card-text")
            ])
        ], className="mb-4"))),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Órdenes Únicas", className="card-title"),
                    html.H2(f"{metrics_marketplace['unique_orders']}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4),
            
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Empresas Únicas", className="card-title"),
                    html.H2(f"{metrics_marketplace['unique_companies']}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4),
            
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Monto Total", className="card-title"),
                    html.H2(f"${metrics_marketplace['total_amount']:,.2f}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4)
        ], className="mb-4"),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardHeader(html.H5("Distribución por Empresa")),
                dbc.CardBody(dcc.Graph(
                    figure=px.bar(
                        metrics_marketplace['amount_by_company'],
                        x='Organization',
                        y='TCV Item',
                        labels={'Organization': 'Empresa', 'TCV Item': 'Monto (USD)'},
                        text_auto='.2f'
                    ).update_layout(showlegend=False)
                ))
            ], className="shadow-sm"), md=6),
            
            dbc.Col(dbc.Card([
                dbc.CardHeader(html.H5("Top Empresas")),
                dbc.CardBody(dash_table.DataTable(
                    columns=[
                        {"name": "Empresa", "id": "Organization"},
                        {"name": "Monto Total", "id": "TCV Item", "type": "numeric", "format": {"specifier": "$,.2f"}}
                    ],
                    data=metrics_marketplace['amount_by_company'].to_dict('records'),
                    style_table={'overflowX': 'auto'},
                    style_cell={'textAlign': 'left'},
                    style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'}
                ))
            ], className="shadow-sm"), md=6)
        ], className="mb-4"),
        
        dbc.Row(dbc.Col(dbc.Card([
            dbc.CardHeader(html.H5("Detalle de Transacciones")),
            dbc.CardBody(dash_table.DataTable(
                columns=[
                    {"name": "Orden ID", "id": "Order id"},
                    {"name": "Empresa", "id": "Organization"},
                    {"name": "Producto", "id": "Product"},
                    {"name": "Monto", "id": "TCV Item", "type": "numeric", "format": {"specifier": "$,.2f"}},
                    {"name": "Fecha", "id": "Date Creation Order"},
                    {"name": "Estado", "id": "Order status"}
                ],
                data=market_orders.to_dict('records'),
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left'},
                style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'}
            ))
        ], className="shadow-sm")))
    ], fluid=True)

# =============================================
# Layout principal con pestañas
# =============================================

app.layout = dbc.Container([
    dbc.Row(dbc.Col(html.H1("Panel de Control Integral", className="text-center my-4"))),
    
    dbc.Tabs([
        dbc.Tab(create_dashboard_empresas(), label="Empresas"),
        dbc.Tab(create_dashboard_subscripciones(), label="Subscripciones"),
        dbc.Tab(create_marketplace_dashboard(), label="Marketplace"),
    ])
], fluid=True)

# =============================================
# Callbacks
# =============================================

@app.callback(
    Output("descargar-empresas", "data"),
    Input("btn-descargar-empresas", "n_clicks"),
    prevent_initial_call=True
)
def descargar_empresas(n_clicks):
    return dcc.send_data_frame(df_organizations[df_organizations['segment'].isna()][['owner', 'name']].to_excel, 
                             "empresas_sin_segmento.xlsx", index=False)

@app.callback(
    Output("descargar-subscripciones", "data"),
    Input("btn-descargar-subscripciones", "n_clicks"),
    prevent_initial_call=True
)
def descargar_subscripciones(n_clicks):
    return dcc.send_data_frame(df_subscriptions_filtrado.to_excel, 
                             "subscripciones_filtradas.xlsx", index=False)

@app.callback(
    Output("descargar-resumen", "data"),
    Input("btn-descargar-resumen", "n_clicks"),
    prevent_initial_call=True
)
def descargar_resumen(n_clicks):
    return dcc.send_data_frame(resumen_owner_pais.to_excel, 
                             "resumen_owner_pais.xlsx", index=False)

# Ejecutar la aplicación
if __name__ == "__main__":
    app.run(debug=False, port=8089)
