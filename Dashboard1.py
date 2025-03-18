# Asegúrate de que este sea el nombre de tu archivo: Dashboard1.py

import pandas as pd
import plotly.express as px
from dash import Dash, html, dcc, dash_table, Input, Output
import dash_bootstrap_components as dbc

# Aquí debes cargar tus archivos Excel
file_name_organizations = "detail-organizations-2025-03-18.xlsx"
file_name_subscriptions = "detail-subscription-2025-03-18.xlsx"

# Cargar los datos en DataFrames
df_organizations = pd.read_excel(file_name_organizations, sheet_name="Organizations")
df_subscriptions = pd.read_excel(file_name_subscriptions)

# Normalizar los valores de 'status' en el primer archivo
df_organizations['status'] = df_organizations['status'].str.strip().str.capitalize()

# Filtrar los datos del segundo archivo
df_subscriptions_filtrado = df_subscriptions[(df_subscriptions['status'] == 'active') & (df_subscriptions['company'].isna())]

# Calcular empresas activas y pendientes
empresas_activas = df_organizations[df_organizations['status'] == 'Active'].shape[0]
empresas_pendientes = df_organizations[df_organizations['status'] == 'Pending'].shape[0]

# Análisis por status (active/pending)
status_summary = df_organizations['status'].value_counts().reset_index()
status_summary.columns = ['Status', 'Cantidad']

# Análisis por KAM (owner)
kam_status_summary = df_organizations.groupby(['owner', 'status']).size().unstack()
kam_status_summary = kam_status_summary.fillna(0)  # Reemplazar NaN con 0
kam_status_summary = kam_status_summary.reindex(columns=['Active', 'Pending'], fill_value=0)  # Asegurar columnas
kam_status_summary['Total'] = kam_status_summary.sum(axis=1)  # Agregar columna Total
kam_status_summary = kam_status_summary.reset_index()  # Resetear índice

# Crear la aplicación
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Definir el layout del dashboard
app.layout = dbc.Container([
    dbc.Row(dbc.Col(html.H1("Dashboard de Empresas y Subscripciones", className="text-center my-4"))),

    # Sección 1: Empresas por estado (Active/Pending)
    dbc.Row([
        dbc.Col([
            html.H2("Empresas por Estado", className="text-center"),
            
            # Tarjeta con el gráfico circular y los números de empresas activas y pendientes
            dbc.Card([
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col(
                            dcc.Graph(
                                id='grafico-status-pie',
                                figure=px.pie(status_summary, values='Cantidad', names='Status', title="Distribución de Empresas por Estado")
                            ), width=6
                        ),
                        dbc.Col([
                            html.H2(f"Empresas Activas: {empresas_activas}", className="text-center"),
                            html.H2(f"Empresas Pendientes: {empresas_pendientes}", className="text-center")
                        ], width=6, className="d-flex flex-column justify-content-center")
                    ])
                ])
            ], className="mb-4")
        ], width=12)
    ]),

    # Sección 2: Empresas sin segmento
    dbc.Row([
        dbc.Col([
            html.H2("Empresas sin Segmento", className="text-center"),
            html.H3(f"Número de Empresas sin Segmento Asignado: {df_organizations['segment'].isna().sum()}", className="text-center"),
            
            # Tarjeta con la tabla de empresas sin segmento (con paginación)
            dbc.Card([
                dbc.CardBody([
                    dash_table.DataTable(
                        id='tabla-empresas',
                        columns=[{"name": i, "id": i} for i in df_organizations[df_organizations['segment'].isna()][['owner', 'name']].columns],
                        data=df_organizations[df_organizations['segment'].isna()][['owner', 'name']].to_dict('records'),
                        style_table={'height': '300px', 'overflowY': 'auto'},
                        style_cell={'textAlign': 'left', 'padding': '10px'},
                        page_size=10  # Mostrar 10 filas por página
                    ),
                    html.Br(),
                    dbc.Button("Descargar Empresas sin Segmento", id="btn-descargar-empresas", color="primary"),
                    dcc.Download(id="descargar-empresas")
                ])
            ], className="mb-4"),
            
            # Tarjeta con el gráfico de empresas sin segmento por propietario
            dbc.Card([
                dbc.CardBody([
                    dcc.Graph(
                        id='grafico-empresas',
                        figure=px.bar(
                            df_organizations[df_organizations['segment'].isna()]['owner'].value_counts().reset_index().rename(columns={'index': 'owner', 0: 'count'}),
                            x='owner',  # Usar 'owner' para los valores únicos
                            y='count',  # Usar 'count' para las frecuencias
                            title="Empresas sin Segmento por Propietario",
                            labels={'owner': 'Propietario', 'count': 'Cantidad'}
                        )
                    )
                ])
            ])
        ], width=12)
    ]),

    # Sección 3: Subscripciones activas sin compañía (en pestañas)
    dbc.Row([
        dbc.Col([
            html.H2("Subscripciones Activas sin Compañía", className="text-center"),
            html.H3(f"Registros Filtrados: {len(df_subscriptions_filtrado)}", className="text-center"),
            
            # Pestañas para organizar la información
            dbc.Tabs([
                # Pestaña 1: Tabla de subscripciones filtradas (con paginación)
                dbc.Tab([
                    dbc.Card([
                        dbc.CardBody([
                            dash_table.DataTable(
                                id='tabla-subscripciones',
                                columns=[{"name": i, "id": i} for i in df_subscriptions_filtrado.columns],
                                data=df_subscriptions_filtrado.to_dict('records'),
                                style_table={'height': '300px', 'overflowY': 'auto'},
                                style_cell={'textAlign': 'left', 'padding': '10px'},
                                page_size=10  # Mostrar 10 filas por página
                            ),
                            html.Br(),
                            dbc.Button("Descargar Subscripciones Filtradas", id="btn-descargar-subscripciones", color="primary"),
                            dcc.Download(id="descargar-subscripciones")
                        ])
                    ], className="mb-4")
                ], label="Tabla de Subscripciones"),

                # Pestaña 2: Gráfico de distribución por `console_domain`
                dbc.Tab([
                    dbc.Card([
                        dbc.CardBody([
                            dcc.Graph(
                                id='grafico-console-domain',
                                figure=px.bar(
                                    df_subscriptions_filtrado['console_domain'].value_counts().reset_index().rename(columns={'index': 'console_domain', 0: 'count'}),
                                    x='console_domain', y='count',
                                    title="Distribución por dominio",
                                    labels={'console_domain': 'dominio', 'count': 'Cantidad'}
                                )
                            )
                        ])
                    ])
                ], label="Distribución por Dominio"),

                # Pestaña 3: Gráfico de distribución por `product`
                dbc.Tab([
                    dbc.Card([
                        dbc.CardBody([
                            dcc.Graph(
                                id='grafico-product',
                                figure=px.pie(df_subscriptions_filtrado, names='product', title="Distribución por Producto")
                            )
                        ])
                    ])
                ], label="Distribución por Producto")
            ])
        ], width=12)
    ])
], fluid=True)

# Callback para la descarga de datos de empresas sin segmento
@app.callback(
    Output("descargar-empresas", "data"),
    Input("btn-descargar-empresas", "n_clicks"),
    prevent_initial_call=True
)
def descargar_empresas(n_clicks):
    return dcc.send_data_frame(df_organizations[df_organizations['segment'].isna()][['owner', 'name']].to_excel, "empresas_sin_segmento.xlsx", index=False)

# Callback para la descarga de datos de subscripciones filtradas
@app.callback(
    Output("descargar-subscripciones", "data"),
    Input("btn-descargar-subscripciones", "n_clicks"),
    prevent_initial_call=True
)
def descargar_subscripciones(n_clicks):
    return dcc.send_data_frame(df_subscriptions_filtrado.to_excel, "subscripciones_filtradas.xlsx", index=False)

# Ejecutar la aplicación
if __name__ == "__main__":
    app.run_server(debug=False)

#----End Of Cell----

