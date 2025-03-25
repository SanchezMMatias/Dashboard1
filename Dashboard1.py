import dash
from dash import html, dash_table, dcc
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import logging
from datetime import datetime

# Configurar logging
logging.basicConfig(level=logging.INFO)

# Simulación de base de datos de usuarios
USUARIOS = {
    'admin': 'password123',
    'usuario1': 'clave456'
}

# Inicializar la aplicación Dash
app = dash.Dash(__name__, 
                external_stylesheets=[dbc.themes.BOOTSTRAP], 
                suppress_callback_exceptions=True)

# Importante: Exponer el servidor WSGI para Gunicorn
server = app.server


# =============================================
# Layout de inicio de sesión
# =============================================

def login_layout():
    """Layout de inicio de sesión"""
    return dbc.Container([
        dbc.Row([
            dbc.Col([
                html.H2("Iniciar Sesión", className="text-center mt-4"),
                dbc.Card([
                    dbc.CardBody([
                        dbc.Form([
                            dbc.Input(id="input-username", type="text", placeholder="Usuario", className="mb-2"),
                            dbc.Input(id="input-password", type="password", placeholder="Contraseña", className="mb-2"),
                            dbc.Button("Iniciar Sesión", color="primary", id="login-button", className="w-100"),
                            html.Div(id="login-message", className="mt-3 text-center")
                        ])
                    ])
                ])
            ], width={"size": 6, "offset": 3})
        ])
    ], fluid=True)

# =============================================
# Funciones para cargar y preparar datos
# =============================================

def load_and_prepare_data():
    """Carga y prepara los datos para el dashboard"""
    try:
        # Cargar los archivos Excel (ajusta estas rutas)
        file_name_organizations = "detail-organizations-2025-03-24.xlsx"
        file_name_subscriptions = "detail-subscription-2025-03-24.xlsx"
        
        df_organizations = pd.read_excel(file_name_organizations, sheet_name="Organizations")
        df_subscriptions = pd.read_excel(file_name_subscriptions)
        
        # Procesamiento de datos
        df_organizations['status'] = df_organizations['status'].str.strip().str.capitalize()
        df_subscriptions_filtrado = df_subscriptions[(df_subscriptions['status'] == 'active') & (df_subscriptions['company'].isna())]
        
        # Cálculos
        empresas_activas = df_organizations[df_organizations['status'] == 'Active'].shape[0]
        empresas_pendientes = df_organizations[df_organizations['status'] == 'Pending'].shape[0]
        empresas_suspendidas = df_organizations[df_organizations['status'] == 'Suspended'].shape[0]
        
        status_summary = df_organizations['status'].value_counts().reset_index()
        status_summary.columns = ['Status', 'Cantidad']
        
        # Resumen por owner y país
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
        
        # Marketplace data
        df_orders = pd.read_excel("detail-order-2025-01-01-to-2025-03-24.xlsx")  
        df_orders['Date Creation Order'] = pd.to_datetime(df_orders['Date Creation Order'], format='%d-%m-%Y')
        df_orders['TCV Item'] = (
            df_orders['TCV Item']
            .astype(str)
            .str.replace('[^\d.,-]', '', regex=True)
            .str.replace(',', '.')
            .astype(float)
        )
        df_orders['Order Type'] = df_orders['Order created by'].apply(
            lambda x: 'Orion Hub' if isinstance(x, str) and x.endswith('@orion.global') else 'Market'
        )
        
        start_date = datetime(2025, 1, 1)
        end_date = datetime(2025, 3, 24)
        market_orders = df_orders[
            (df_orders['Date Creation Order'] >= start_date) & 
            (df_orders['Date Creation Order'] <= end_date) & 
            (df_orders['Order Type'] == 'Market')
        ].copy()

        metrics_marketplace = {
            'unique_orders': market_orders['Order id'].nunique(),
            'unique_companies': market_orders['Organization'].nunique(),
            'total_amount': market_orders['TCV Item'].sum(),
            'amount_by_company': market_orders.groupby('Organization')['TCV Item'].sum().reset_index()
        }
        
        return {
            'df_organizations': df_organizations,
            'df_subscriptions_filtrado': df_subscriptions_filtrado,
            'empresas_activas': empresas_activas,
            'empresas_pendientes': empresas_pendientes,
            'empresas_suspendidas': empresas_suspendidas,
            'status_summary': status_summary,
            'resumen_owner_pais': resumen_owner_pais,
            'market_orders': market_orders,
            'metrics_marketplace': metrics_marketplace,
            'start_date': start_date,
            'end_date': end_date
        }
        
    except Exception as e:
        logging.error(f"Error al cargar datos: {str(e)}")
        return None

# =============================================
# Layout del dashboard principal
# =============================================

def create_main_dashboard(data):
    """Crea el dashboard principal con todas las pestañas"""
    # Dashboard de Empresas
    empresas_tab = dbc.Container([
        dbc.Row(dbc.Col(html.H1("Dashboard de Empresas", className="text-center my-4"))),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col(dcc.Graph(
                            figure=px.pie(data['status_summary'], values='Cantidad', names='Status', 
                                        title="Distribución de Empresas por Estado")
                        ), width=6),
                        dbc.Col([
                            html.H2(f"Empresas Activas: {data['empresas_activas']}", className="text-center"),
                            html.H2(f"Empresas Pendientes: {data['empresas_pendientes']}", className="text-center"),
                            html.H2(f"Empresas Suspendidas: {data['empresas_suspendidas']}", className="text-center")
                        ], width=6, className="d-flex flex-column justify-content-center")
                    ])
                ])
            ], className="mb-4"), width=12)
        ]),

        dbc.Row([
            dbc.Col([
                html.H2("Empresas sin Segmento", className="text-center"),
                html.H3(f"Número de Empresas sin Segmento Asignado: {data['df_organizations']['segment'].isna().sum()}", className="text-center"),
                
                dbc.Card([
                    dbc.CardBody([
                        dash_table.DataTable(
                            columns=[{"name": i, "id": i} for i in data['df_organizations'][data['df_organizations']['segment'].isna()][['owner', 'name']].columns],
                            data=data['df_organizations'][data['df_organizations']['segment'].isna()][['owner', 'name']].to_dict('records'),
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
                            data['df_organizations'][data['df_organizations']['segment'].isna()]['owner'].value_counts().reset_index().rename(columns={'index': 'owner', 0: 'count'}),
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
                            columns=[{"name": i, "id": i} for i in data['resumen_owner_pais'].columns],
                            data=data['resumen_owner_pais'].to_dict('records'),
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

    # Dashboard de Subscripciones
    subscripciones_tab = dbc.Container([
        dbc.Row(dbc.Col(html.H1("Dashboard de Subscripciones", className="text-center my-4"))),
        
        dbc.Row([
            dbc.Col([
                html.H2("Subscripciones Activas sin Compañía", className="text-center"),
                html.H3(f"Registros Filtrados: {len(data['df_subscriptions_filtrado'])}", className="text-center"),
                
                dbc.Tabs([
                    dbc.Tab([
                        dbc.Card([
                            dbc.CardBody([
                                dash_table.DataTable(
                                    columns=[{"name": i, "id": i} for i in data['df_subscriptions_filtrado'].columns],
                                    data=data['df_subscriptions_filtrado'].to_dict('records'),
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

                    dbc.Tab([
                        dbc.Card([
                            dbc.CardBody(
                                dcc.Graph(
                                    figure=px.bar(
                                        data['df_subscriptions_filtrado']['console_domain'].value_counts().reset_index().rename(columns={'index': 'console_domain', 0: 'count'}),
                                        x='console_domain', y='count',
                                        title="Distribución por dominio",
                                        labels={'console_domain': 'dominio', 'count': 'Cantidad'}
                                    )
                                )
                            )
                        ])
                    ], label="Distribución por Dominio"),

                    dbc.Tab([
                        dbc.Card([
                            dbc.CardBody(
                                dcc.Graph(
                                    figure=px.pie(data['df_subscriptions_filtrado'], names='product', title="Distribución por Producto")
                                )
                            )
                        ])
                    ], label="Distribución por Producto")
                ])
            ], width=12)
        ])
    ], fluid=True)

    # Dashboard del Marketplace
    marketplace_tab = dbc.Container([
        dbc.Row(dbc.Col(html.H1("Análisis del Marketplace", className="text-center my-4"))),
        
        dbc.Row(dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Período Analizado", className="card-title"),
                html.P(f"{data['start_date'].strftime('%d/%m/%Y')} - {data['end_date'].strftime('%d/%m/%Y')}", className="card-text")
            ])
        ], className="mb-4"))),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Órdenes Únicas", className="card-title"),
                    html.H2(f"{data['metrics_marketplace']['unique_orders']}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4),
            
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Empresas Únicas", className="card-title"),
                    html.H2(f"{data['metrics_marketplace']['unique_companies']}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4),
            
            dbc.Col(dbc.Card([
                dbc.CardBody([
                    html.H5("Monto Total", className="card-title"),
                    html.H2(f"${data['metrics_marketplace']['total_amount']:,.2f}", className="card-text text-center")
                ])
            ], className="shadow-sm"), md=4)
        ], className="mb-4"),
        
        dbc.Row([
            dbc.Col(dbc.Card([
                dbc.CardHeader(html.H5("Distribución por Empresa")),
                dbc.CardBody(dcc.Graph(
                    figure=px.bar(
                        data['metrics_marketplace']['amount_by_company'],
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
                    data=data['metrics_marketplace']['amount_by_company'].to_dict('records'),
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
                data=data['market_orders'].to_dict('records'),
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left'},
                style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'}
            ))
        ], className="shadow-sm")))
    ], fluid=True)

    # Layout principal con pestañas
    return dbc.Container([
        dbc.Row([
            dbc.Col(html.H1("Panel de Control Integral", className="text-center my-4")),
            dbc.Col(dbc.Button("Cerrar Sesión", id="logout-button", color="danger", className="mt-3 float-end"), width=2)
        ]),
        
        dbc.Tabs([
            dbc.Tab(empresas_tab, label="Empresas"),
            dbc.Tab(subscripciones_tab, label="Subscripciones"),
            dbc.Tab(marketplace_tab, label="Marketplace"),
        ])
    ], fluid=True)

# =============================================
# Diseño inicial de la aplicación
# =============================================

app.layout = html.Div([
    dcc.Store(id='login-state', storage_type='session'),
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content')
])

# =============================================
# Callbacks
# =============================================

# Callback para verificar sesión y cargar contenido inicial
@app.callback(
    Output('page-content', 'children'),
    [Input('url', 'pathname'),
     Input('login-state', 'data')]
)
def display_page(pathname, login_state):
    # Si hay sesión activa, mostrar dashboard, sino mostrar login
    if login_state and 'username' in login_state:
        # Usuario ya autenticado, cargar dashboard
        data = load_and_prepare_data()
        if data:
            return create_main_dashboard(data)
        else:
            # Error al cargar datos
            return html.Div([
                html.H3("Error al cargar los datos"),
                dbc.Button("Volver al login", id="logout-button", color="primary")
            ])
    else:
        # No hay sesión, mostrar login
        return login_layout()

@app.callback(
    [Output('login-state', 'data', allow_duplicate=True),
     Output('login-message', 'children')],
    [Input('login-button', 'n_clicks')],
    [State('input-username', 'value'),
     State('input-password', 'value')],
    prevent_initial_call=True
)
def login_process(n_clicks, username, password):
    if not n_clicks:
        return dash.no_update, dash.no_update
    
    if not username or not password:
        return None, dbc.Alert("Por favor, ingrese usuario y contraseña", color="warning")
    
    if username in USUARIOS and USUARIOS[username] == password:
        # Autenticación exitosa
        return {'username': username}, dbc.Alert("Inicio de sesión exitoso", color="success")
    else:
        # Credenciales incorrectas
        return None, dbc.Alert("Credenciales incorrectas", color="danger")

@app.callback(
    Output('login-state', 'data', allow_duplicate=True),
    Input('logout-button', 'n_clicks'),
    prevent_initial_call=True
)
def logout_process(n_clicks):
    # Limpiar los datos de la sesión
    return None

# Callbacks para descargas
@app.callback(
    Output("descargar-empresas", "data"),
    Input("btn-descargar-empresas", "n_clicks"),
    prevent_initial_call=True
)
def descargar_empresas(n_clicks):
    data = load_and_prepare_data()
    return dcc.send_data_frame(data['df_organizations'][data['df_organizations']['segment'].isna()][['owner', 'name']].to_excel, 
                             "empresas_sin_segmento.xlsx", index=False)

@app.callback(
    Output("descargar-subscripciones", "data"),
    Input("btn-descargar-subscripciones", "n_clicks"),
    prevent_initial_call=True
)
def descargar_subscripciones(n_clicks):
    data = load_and_prepare_data()
    return dcc.send_data_frame(data['df_subscriptions_filtrado'].to_excel, 
                             "subscripciones_filtradas.xlsx", index=False)

@app.callback(
    Output("descargar-resumen", "data"),
    Input("btn-descargar-resumen", "n_clicks"),
    prevent_initial_call=True
)
def descargar_resumen(n_clicks):
    data = load_and_prepare_data()
    return dcc.send_data_frame(data['resumen_owner_pais'].to_excel, 
                             "resumen_owner_pais.xlsx", index=False)

# =============================================
# Ejecutar la aplicación
# =============================================

if __name__ == "__main__":
    app.run(debug=True)
