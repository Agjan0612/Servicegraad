import dash
import pandas as pd
from dash import Dash, html, dash_table, dcc, callback, Output, Input
from dash.exceptions import PreventUpdate
import plotly.express as px
import dash_bootstrap_components as dbc
import numpy as np
from dash_bootstrap_templates import load_figure_template
import openpyxl as pxl
import gunicorn
import dash_ag_grid as dag


pd.options.display.width= None
pd.options.display.max_columns= None
pd.set_option('display.max_rows', 3000)
pd.set_option('display.max_columns', 3000)

# dataframes maken en kolommen inlezen
kol_recept = pd.read_excel('Kolommen receptverwerking.xlsx')
kol_ass = pd.read_excel('Kolommen assortiment.xlsx')
columns_recept = kol_recept.columns
columns_assort = kol_ass.columns
recept = pd.read_csv('recept.txt')
assortiment = pd.read_csv('assortiment.txt')
recept.columns = columns_recept
assortiment.columns = columns_assort

# kolommen toevoegen aan receptverwerking

recept['ddDatumRecept'] = pd.to_datetime(recept['ddDatumRecept'])
recept['maand'] = recept['ddDatumRecept'].dt.month

#kolommen toevoegen aan assortiment

assortiment['voorraadmaximum'] = pd.to_numeric(assortiment['voorraadmaximum'], errors='coerce')
assortiment['voorraadmaximum'] = assortiment['voorraadmaximum'].replace(np.nan,0)
assortiment['voorraadmaximum'] = assortiment['voorraadmaximum'].astype(int)




# Maak een dataframe voor het dashboard
dashboard_data = recept[['ddDatumRecept','maand', 'ReceptHerkomst', 'cf',
       'ndReceptnummer', 'ndATKODE', 'sdEtiketNaam', 'ndAantal', 'Uitgifte', 'ndVoorraadTotaal']]



# maak een dataframe voor de taartdiagram

# filters definieren --> alleen ladekast verstrekkingen EU en TU






# Bouw de app

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

app.layout = dbc.Container([
    dbc.Row([html.H1('Servicegraad Dashboard')]),
    dbc.Row([html.H5('Upload op github het: assortimentsbestand, receptverwerkingsbestand')]),
    dbc.Row([dcc.RangeSlider(id='periode', min=1, max=12, step=1, marks={1:'jan', 2:'febr', 3:'mrt', 4:'apr', 5:'mei', 6:'jun', 7:'jul', 8:'aug', 9:'sept', 10:'okt', 11:'nov', 12:'dec'}, value=[1, 12])]),
    dbc.Row([dcc.Graph(id='servicegraad')]),
    dbc.Row([html.H4('Defecturen tabel (EU/TU leveringen apotheek)')]),
    dbc.Row([html.Div(id='tabel')]),
    dbc.Row([dbc.Col([], width=5), dbc.Col([dbc.Button(id='knop',children='defecturen xlsx', color="success", className="me-1")]), dbc.Col([], width=5)]),
    dbc.Row([]),
    dbc.Row([dcc.Download(id='download')]),

])

# Service-graad taartdiagram
@callback(
    Output('servicegraad', 'figure'),
    Input('periode', 'value')
)
def service_graad(periode):
    # maak een dataframe voor de taartdiagram

    # filters definieren --> alleen ladekast verstrekkingen EU en TU

    # filters
    cf_nee = (dashboard_data['cf'] == 'N')
    geen_VU = (dashboard_data['Uitgifte'] != 'VU')
    geen_ONB = (dashboard_data['Uitgifte'] != 'ONB')
    geen_dienst = (dashboard_data['ReceptHerkomst'] != 'DIENST')
    geen_distributie = (dashboard_data['ReceptHerkomst'] != 'D')
    geen_zorgregel = (dashboard_data['ReceptHerkomst'] != 'Z')
    service_periode_1 = (dashboard_data['maand'] >= periode[0])  # range slider
    service_periode_2 = (dashboard_data['maand'] <= periode[1])  # range slider

    service_data = dashboard_data.loc[
        cf_nee & geen_VU & geen_ONB & geen_dienst & geen_distributie & geen_zorgregel & service_periode_1 & service_periode_2]

    # defecturen kolom maken
    service_data['vrd na aanschrijven'] = service_data['ndVoorraadTotaal'] - service_data['ndAantal']
    voorwaarde = [service_data['vrd na aanschrijven'] < 0, service_data['vrd na aanschrijven'] >= 0]
    categorie = ['defectuur', 'voorraad toereikend']
    service_data['defectuur'] = np.select(voorwaarde, categorie, default='defectuur')

    service_data_1 = service_data.groupby(by=['defectuur'])['defectuur'].count().to_frame('aantal').reset_index()
    service_graad = px.pie(service_data_1, names='defectuur',values='aantal', title='SERVICE GRAAD OVER MEETPERIODE (EU/TU leveringen)')
    return service_graad

@callback(
    Output('tabel', 'children'),
    Input('periode', 'value')
)
def tabel(periode):
    # Tabel maken (tabel met defecturen, waarbij je de min/max, locatie en voorraad ziet

    # filters definieren
    # filters
    cf_nee = (dashboard_data['cf'] == 'N')
    geen_VU = (dashboard_data['Uitgifte'] != 'VU')
    geen_ONB = (dashboard_data['Uitgifte'] != 'ONB')
    geen_dienst = (dashboard_data['ReceptHerkomst'] != 'DIENST')
    geen_distributie = (dashboard_data['ReceptHerkomst'] != 'D')
    geen_zorgregel = (dashboard_data['ReceptHerkomst'] != 'Z')
    tabel_periode_1 = (dashboard_data['maand'] >= periode[0])  # range slider
    tabel_periode_2 = (dashboard_data['maand'] <= periode[1])  # range slider

    tabel_data = dashboard_data.loc[
        cf_nee & geen_VU & geen_ONB & geen_dienst & geen_distributie & geen_zorgregel & tabel_periode_1 & tabel_periode_2]
    # defecturen kolom maken
    tabel_data['vrd na aanschrijven'] = tabel_data['ndVoorraadTotaal'] - tabel_data['ndAantal']
    voorw = [tabel_data['vrd na aanschrijven'] < 0, tabel_data['vrd na aanschrijven'] >= 0]
    cat = ['defectuur', 'voorraad toereikend']
    tabel_data['defectuur'] = np.select(voorw, cat, default='defectuur')

    # filter zodat je alleen defecturen ziet
    alleen_defecturen = (tabel_data['defectuur'] == 'defectuur')
    tabel_data_1 = tabel_data.loc[alleen_defecturen]

    # tel de defecturen per product en sorteer ze aflopend
    tabel_data_2 = tabel_data_1.groupby(by=['ndATKODE', 'sdEtiketNaam'])['ndATKODE'].count().to_frame(
        'aantal defecturen').reset_index()
    tabel_data_3 = tabel_data_2.sort_values(by=['aantal defecturen'], ascending=False)

    # merge met het assortiment

    tabel_defect_assort = tabel_data_3.merge(assortiment[['atckode', 'zinummer', 'artikelnaam',
                                                          'voorraadminimum', 'voorraadmaximum', 'locatie1',
                                                          'voorraadtotaal']],
                                             how='inner',
                                             left_on='ndATKODE',
                                             right_on='zinummer').drop(columns=['sdEtiketNaam', 'ndATKODE'])

    tabel_defect_assort = tabel_defect_assort[['zinummer', 'atckode', 'artikelnaam', 'aantal defecturen',
                                               'voorraadminimum', 'voorraadmaximum', 'locatie1', 'voorraadtotaal']]
    tabel_defect_assort.columns = ['ZI', 'ATC', 'ARTIKEL', 'DEFETUREN AANTAL',
                                    'MINVRD', 'MAXVRD', 'LOC', 'VOORRAAD']

    grid_defect = dag.AgGrid(
        rowData=tabel_defect_assort.to_dict('records'),
        columnDefs=[{'field':i}for i in tabel_defect_assort.columns],
        defaultColDef={"filter": "agTextColumnFilter"},
        dashGridOptions={"enableCellTextSelection": True, "pagination": True,'paginationPageSize':100},
        columnSize="sizeToFit",
        style={"height": 600}


    )
    return grid_defect


# DOWNLOAD KNOP
@callback(
    Output('download', 'data'),
    Output('knop', 'n_clicks'),
    Input('periode', 'value'),
    Input('knop', 'n_clicks')
)
def download(periode, n_clicks):
    # filters definieren
    # filters
    cf_nee = (dashboard_data['cf'] == 'N')
    geen_VU = (dashboard_data['Uitgifte'] != 'VU')
    geen_ONB = (dashboard_data['Uitgifte'] != 'ONB')
    geen_dienst = (dashboard_data['ReceptHerkomst'] != 'DIENST')
    geen_distributie = (dashboard_data['ReceptHerkomst'] != 'D')
    geen_zorgregel = (dashboard_data['ReceptHerkomst'] != 'Z')
    tabel_periode_1 = (dashboard_data['maand'] >= periode[0])  # range slider
    tabel_periode_2 = (dashboard_data['maand'] <= periode[1])  # range slider

    tabel_data = dashboard_data.loc[
        cf_nee & geen_VU & geen_ONB & geen_dienst & geen_distributie & geen_zorgregel & tabel_periode_1 & tabel_periode_2]
    # defecturen kolom maken
    tabel_data['vrd na aanschrijven'] = tabel_data['ndVoorraadTotaal'] - tabel_data['ndAantal']
    voorw = [tabel_data['vrd na aanschrijven'] < 0, tabel_data['vrd na aanschrijven'] >= 0]
    cat = ['defectuur', 'voorraad toereikend']
    tabel_data['defectuur'] = np.select(voorw, cat, default='defectuur')

    # filter zodat je alleen defecturen ziet
    alleen_defecturen = (tabel_data['defectuur'] == 'defectuur')
    tabel_data_1 = tabel_data.loc[alleen_defecturen]

    # tel de defecturen per product en sorteer ze aflopend
    tabel_data_2 = tabel_data_1.groupby(by=['ndATKODE', 'sdEtiketNaam'])['ndATKODE'].count().to_frame(
        'aantal defecturen').reset_index()
    tabel_data_3 = tabel_data_2.sort_values(by=['aantal defecturen'], ascending=False)

    # merge met het assortiment

    excel_defect_assort = tabel_data_3.merge(assortiment[['atckode', 'zinummer', 'artikelnaam',
                                                          'voorraadminimum', 'voorraadmaximum', 'locatie1',
                                                          'voorraadtotaal']],
                                             how='inner',
                                             left_on='ndATKODE',
                                             right_on='zinummer').drop(columns=['sdEtiketNaam', 'ndATKODE'])

    excel_defect_assort = excel_defect_assort[['zinummer', 'atckode', 'artikelnaam', 'aantal defecturen',
                                               'voorraadminimum', 'voorraadmaximum', 'locatie1', 'voorraadtotaal']]

    excel_defect_assort.columns = ['ZI', 'ATC', 'ARTIKEL', 'DEFETUREN AANTAL',
                                    'MINVRD', 'MAXVRD', 'LOC', 'VOORRAAD']

    if not n_clicks:
        raise PreventUpdate

    excel = dcc.send_data_frame(excel_defect_assort.to_excel, "defecturen.xlsx", index=False)
    n_clicks = 0
    return excel, n_clicks


if __name__ == '__main__':
    app.run(debug=True)