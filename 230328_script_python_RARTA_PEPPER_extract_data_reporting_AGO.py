#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 27 16:56:12 2023

@author: clementine.garoche
script to extract data from excel run files, for GRTA reporting."
"""
# Importer les librairies
from pathlib import Path
import os
import pandas as pd
import bokeh
from bokeh.models import ColumnDataSource
from bokeh.models.widgets import DataTable, TableColumn, Div
from bokeh.layouts import grid
from bokeh.io import export_png

#changer le dossier dans lequel on travaille # replace 'ago' with 'antago' or 'interf'accordingly
os.chdir('/Users/clementine/Documents/TRAVAIL/PEPPER/RAR/Manip_Inserm/PHASE_1/ago')

#boucle pour appliquer le script dans tous les fichiers xlsx contenus dans le dossier # same as above
mydir = Path('/Users/clementine/Documents/TRAVAIL/PEPPER/RAR/Manip_Inserm/PHASE_1/ago')
for file in mydir.glob('*.xlsm'):
   
# Importer les donnees depuis Excel dans un dataframe pandas

# FOR THE TOP CHEMICAL

#Importer la date du run, la molecule TOP, le mode, le numero du run
    df = pd.read_excel(file,sheet_name = 'Luminescence')
    date = df.iat[2,4]
    molecule_top = df.iat[13,4]
    mode = df.iat[13,6]
    run_top = df.iat[13,7]
    basalmean = df.iat[17,4]
    basalsd = df.iat[17,5]
    IFmean = df.iat[18,4]
    IFsd = df.iat[18,5]
    Ref80mean = '-'
    Ref80sd = '-'
    Coexp= '_'

    data_top = [[date, molecule_top, mode, run_top, basalmean, basalsd, IFmean, IFsd, Ref80mean, Ref80sd, Coexp]]
    df_id_top = pd.DataFrame(data_top, columns = ["Date", "Molecule", "Mode", "Run number", "Basal (%mean)", "Basal (%sd)", "Induction factor (mean)", "Induction factor (sd)", "", "", ""])
   
   
    #Importer les valeurs des conc et de RLU%
    df_values_top = pd.read_excel(file,
                           sheet_name = 'Luminescence', index_col = False, skiprows = 29,  nrows= 7, usecols = 'D,L:O, P, Q, S')
    df_values_top.set_axis(["Conc(M)", "norm1", "norm2", "norm3", "norm4", "mean%", "sd%", "CV%"],
                        axis=1,inplace=True)
   
   
    #Importer viability
    df_viab_top = pd.read_excel(file,
                           sheet_name = 'Viability', index_col = False, skiprows = 29,  nrows= 7, usecols = 'D,L:O, P, Q, S')
    df_viab_top.set_axis(["Conc(M)", "norm1", "norm2", "norm3", "norm4", "mean%", "sd%", "CV%"],
                        axis=1,inplace=True)
   
    # Convert dataframe to column data source
    source_id_top = ColumnDataSource(df_id_top)
    source_values_top = ColumnDataSource(df_values_top)
    source_viab_top = ColumnDataSource(df_viab_top)
   
    # Create tables
    percentformat = bokeh.models.NumberFormatter(format='0.0%')
    scientificformat = bokeh.models.ScientificFormatter(precision=4)
    numberformat = bokeh.models.NumberFormatter(format='0.0')
   
   
    table_columns_id =[TableColumn(field='Date'),
                       TableColumn(field='Molecule'),
                       TableColumn(field='Mode'),
                       TableColumn(field='Run number', width = 40),
                       TableColumn(field='Basal (%mean)', formatter = percentformat, width = 40),
                       TableColumn(field='Basal (%sd)', formatter = percentformat, width = 40),
                       TableColumn(field='Induction factor (mean)', formatter = numberformat, width = 60),
                       TableColumn(field='Induction factor (sd)', formatter = numberformat, width = 60),
                       ]
   
    table_columns_values = [TableColumn(field='Conc(M)', formatter = scientificformat),
    TableColumn(field='norm1', formatter = percentformat),
             TableColumn(field='norm2', formatter = percentformat),
             TableColumn(field='norm3', formatter = percentformat),
             TableColumn(field='norm4', formatter = percentformat),
             TableColumn(field='mean%', formatter = percentformat),
             TableColumn(field='sd%', formatter = percentformat),
             TableColumn(field='CV%', formatter = percentformat)
             ]
   
    table_columns_viab = [TableColumn(field='Conc(M)', formatter = scientificformat),
    TableColumn(field='norm1', formatter = percentformat),
             TableColumn(field='norm2', formatter = percentformat),
             TableColumn(field='norm3', formatter = percentformat),
             TableColumn(field='norm4', formatter = percentformat),
             TableColumn(field='mean%', formatter = percentformat),
             TableColumn(field='sd%', formatter = percentformat),
             TableColumn(field='CV%', formatter = percentformat)
             ]
   
    p_id_top = DataTable(source=source_id_top, columns=table_columns_id, height = 50, width = 1300)
    p_values_top = DataTable(source=source_values_top, columns=table_columns_values , height = 200)
    p_viab_top = DataTable(source=source_viab_top, columns=table_columns_values, height = 200)
   
    #Add titres
    Lum_title = Div(text="""<b>Luminescence</b>""")
    Viab_title = Div(text="""<b>Viability</b>""")
   
    # Arranger les tables
    Grid_top = grid([[p_id_top], [Lum_title, Viab_title], [p_values_top, p_viab_top]])
       
    #export grid
    date_str = str(date)
    molecule_top_str = str(molecule_top)
    run_top_str = str(run_top)
    names_top = [molecule_top_str, mode, run_top_str, date_str]
    namefile_top = '_'.join(names_top) + '.png'
   
    export_png(Grid_top, filename=namefile_top)
   
    #same for BOTTOM MOLECULE, in the same loop
    molecule_bot = df.iat[14,4]
    run_bot = df.iat[14,7]
   
    if molecule_bot == "no compound" or isinstance(molecule_bot, float):
        continue
   
    data_bot = [[date, molecule_bot, mode, run_bot, basalmean, basalsd, IFmean, IFsd, Ref80mean, Ref80sd, Coexp]]
    df_id_bot = pd.DataFrame(data_bot, columns = ["Date", "Molecule", "Mode", "Run number", "Basal (%mean)", "Basal (%sd)", "Induction factor (mean)", "Induction factor (sd)", "", "", ""])
       
       
    #Importer les valeurs des conc et de RLU%
    df_values_bot = pd.read_excel(file,
        sheet_name = 'Luminescence', index_col = False, skiprows = 41,  nrows= 7, usecols = 'D,L:O, P, Q, S')
    df_values_bot.set_axis(["Conc(M)", "norm1", "norm2", "norm3", "norm4", "mean%", "sd%", "CV%"],
        axis=1,inplace=True)
       
       
    #Importer viability
    df_viab_bot = pd.read_excel(file,
        sheet_name = 'Viability', index_col = False, skiprows = 41,  nrows= 7, usecols = 'D,L:O, P, Q, S')
    df_viab_bot.set_axis(["Conc(M)", "norm1", "norm2", "norm3", "norm4", "mean%", "sd%", "CV%"],
        axis=1,inplace=True)
       
    # Convert dataframe to column data source
    source_id_bot = ColumnDataSource(df_id_bot)
    source_values_bot = ColumnDataSource(df_values_bot)
    source_viab_bot = ColumnDataSource(df_viab_bot)
       
    # Create tables
    p_id_bot = DataTable(source=source_id_bot, columns=table_columns_id, height = 50, width = 1300)
    p_values_bot = DataTable(source=source_values_bot, columns=table_columns_values, height = 200)
    p_viab_bot = DataTable(source=source_viab_bot, columns=table_columns_values, height = 200)
       
    # Arranger les tables
    Grid_bot = grid([[p_id_bot], [Lum_title, Viab_title], [p_values_bot, p_viab_bot]])
           
    #export grid
    molecule_bot_str = str(molecule_bot)
    run_bot_str = str(run_bot)
    names_bot = [molecule_bot_str, mode, run_bot_str, date_str]
    namefile_bot = '_'.join(names_bot) + '.png'
       
    export_png(Grid_bot, filename=namefile_bot)    
   

   
   
   
    #conda install selenium geckodriver firefox -c conda-forge
    
    #improvements: pas bokeh car pas adapté pour statique, trop lent. Deprecated format openpyxl ? deprecated d'ajouter axis pandas.

