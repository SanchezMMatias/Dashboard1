# Importar las bibliotecas necesarias
import pandas as pd
from dash import Dash, html, dash_table, dcc, Input, Output, State
import plotly.express as px
import dash_bootstrap_components as dbc
import logging

# Configurar logging para poder ver errores en Render
logging.basicConfig(level=logging.INFO)

# Cargar los archivos Excel desde las rutas locales
file_name_organizations = "detail-organizations-2025-03-19.xlsx"  # Ruta del primer archivo
file_name_subscriptions = "detail-subscription-2025-03-19.xlsx"  # Ruta del segundo archivo
logging.info(f"Archivos cargados: {file_name_organizations}, {file_name_subscriptions}")

# Cargar los datos en DataFrames
try:
    df_organizations = pd.read_excel(file_name_organizations, sheet_name="Organizations")
    df_subscriptions = pd.read_excel(file_name_subscriptions)
    logging.info("DataFrames cargados correctamente")
except Exception as e:
    logging.error(f"Error al cargar archivos: {str(e)}")
    # Creamos DataFrames vacíos para evitar errores
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
kam_status_summary = kam_status_summary.fillna(0)  # Reemplazar NaN con 0
kam_status_summary = kam_status_summary.reindex(columns=['Active', 'Pending', 'Suspended'], fill_value=0)  # Asegurar columnas
kam_status_summary['Total'] = kam_status_summary.sum(axis=1)  # Agregar columna Total
kam_status_summary = kam_status_summary.reset_index()  # Resetear índice

# CORRECCIÓN: Crear el resumen por owner y país de manera simplificada
try:
    # Inicializar DataFrame vacío
    resumen_owner_pais = pd.DataFrame()
    
    # Obtener valores únicos de owner y country
    owners = df_organizations['owner'].dropna().unique()
    
    # Para cada combinación owner-país, calcular los valores
    for owner in owners:
        # Filtrar por owner
        df_owner = df_organizations[df_organizations['owner'] == owner]
        countries = df_owner['country'].dropna().unique()
        
        for country in countries:
            # Filtrar por país
            df_filtered = df_owner[df_owner['country'] == country]
            
            # Contar por status
            active = df_filtered[df_filtered['status'] == 'Active'].shape[0]
            pending = df_filtered[df_filtered['status'] == 'Pending'].shape[0]
            suspended = df_filtered[df_filtered['status'] == 'Suspended'].shape[0]
            
            # Calcular total y avance
            total = active + pending + suspended
            avance = round((active / total) * 100, 2) if total > 0 else 0
            
            # Crear nueva fila
            nueva_fila = pd.DataFrame({
                'owner': [owner],
                'country': [country],
                'Active': [active],
                'Pending': [pending],
                'Suspended': [suspended],
                'Total': [total],
                'Avance': [avance]
            })
            
            # Añadir al DataFrame principal
            resumen_owner_pais = pd.concat([resumen_owner_pais, nueva_fila], ignore_index=True)
    
    # Asegurar tipos de datos correctos
    resumen_owner_pais['Active'] = resumen_owner_pais['Active'].astype(int)
    resumen_owner_pais['Pending'] = resumen_owner_pais['Pending'].astype(int)
    resumen_owner_pais['Suspended'] = resumen_owner_pais['Suspended'].astype(int)
    resumen_owner_pais['Total'] = resumen_owner_pais['Total'].astype(int)
    resumen_owner_pais['Avance'] = resumen_owner_pais['Avance'].astype(float)
    
    logging.info("Resumen por Owner y País creado correctamente")
    logging.info(f"Filas en resumen: {len(resumen_owner_pais)}")
    
except Exception as e:
    logging.error(f"Error al crear resumen por owner y país: {str(e)}")
    # Crear un DataFrame de ejemplo para evitar errores
    resumen_owner_pais = pd.DataFrame({
        'owner': ['Error'],
        'country': ['Error'],
        'Active': [0],
        'Pending': [0],
        'Suspended': [0],
        'Total': [0],
        'Avance': [0.0]
    })

# Crear la aplicación Dash con un tema de Bootstrap
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Definir el layout del dashboard unificado
app.layout = dbc.Container([
    # Título principal
    dbc.Row(
        dbc.Col(html.H1("Dashboard de Empresas y Subscripciones", className="text-center my-4"))
    ),

    # Sección 1: Empresas por estado (Active/Pending/Suspended)
    dbc.Row([
        dbc.Col([
            # Tarjeta con el gráfico circular y los números de empresas activas, pendientes y suspendidas
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
                            html.H2(f"Empresas Pendientes: {empresas_pendientes}", className="text-center"),
                            html.H2(f"Empresas Suspendidas: {empresas_suspendidas}", className="text-center")
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
            ], className="mb-4")
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

                # Pestaña 2: Gráfico de distribución por console_domain
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

                # Pestaña 3: Gráfico de distribución por product
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
    ]),

    # Sección 4: Resumen por owner y país
    dbc.Row([
        dbc.Col([
            html.H2("Resumen por Owner y País", className="text-center"),
            
            # CORRECCIÓN: Implementar manejo de errores para la tabla resumen
            dbc.Card([
                dbc.CardBody([
                    html.Div(id='container-tabla-resumen', children=[
                        # La tabla se crea dinámicamente para manejar errores potenciales
                        dash_table.DataTable(
                            id='tabla-resumen-owner-pais',
                            columns=[{"name": i, "id": i} for i in resumen_owner_pais.columns],
                            data=resumen_owner_pais.to_dict('records'),
                            style_table={'height': '300px', 'overflowY': 'auto'},
                            style_cell={'textAlign': 'left', 'padding': '10px'},
                            page_size=10  # Mostrar 10 filas por página
                        ),
                    ]),
                    html.Br(),
                    dbc.Button("Descargar Resumen", id="btn-descargar-resumen", color="primary"),
                    dcc.Download(id="descargar-resumen")
                ])
            ], className="mb-4")
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

# Callback para la descarga de datos del resumen por owner y país
@app.callback(
    Output("descargar-resumen", "data"),
    Input("btn-descargar-resumen", "n_clicks"),
    prevent_initial_call=True
)
def descargar_resumen(n_clicks):
    return dcc.send_data_frame(resumen_owner_pais.to_excel, "resumen_owner_pais.xlsx", index=False)

# Ejecutar la aplicación
if __name__ == "__main__":
    app.run(debug=False)
