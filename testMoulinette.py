#Importations
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import re
from fonctions import split_section_analytique, df_journaux, ajouter_type_piece5, completer_date_echeance,merge_custom, merge_Tiers2, merge_Tiers3, merge_journaux, merge_IFRS
from fonctions import retriev_code_PNL_from_df, renommer_colonnes, preparation2, convertir_sage100_en_x31, qualifier_et_controler3, pipeline, nettoyer_dataframe, transformer_ecritures_stock

#Feuille de mapping journal pour chaque société
jLCR = "mapping_LCR"
jMIN = "mapping_MINO"
jLOG = "mapping_LOG"
jSCI = "mapping_LAVION"
jGC = "mapping_GRP"

#Les sociétés:
LCR = 'LCR'
MIN = 'MIN'
LOG = 'LOG'
SCI = 'SCI'
GC = 'GC'

#Chemin des écritures pour chaque mois
cheminJanvier = "datas\data_ecritures\ecritures janvier 2025.xlsx"
cheminFevrier = "datas\data_ecritures\ecritures fevrier 2025.xlsx"
cheminMars = "datas\data_ecritures\ecritures mars 2025.xlsx"

# Données fixes de mapping
df_tiers = pd.read_excel("datas\MappingTousLesTiers.xlsx", sheet_name="ALL++")
df_IFRS = pd.read_excel("datas\Mapping IFRS LCR.xlsx")
pnl = pd.read_excel("datas\Mapping_PNL_Vrai.xlsb")

#Lancement des tests : remplir ces variable et lancer le code
chemin = cheminJanvier
journal = df_journaux(jGC)
societe = GC
mois = 'Janvier'
version = 'F1'

data, dataAN, dfX3, df_filtered, df_tiers_filtered, df_excluded, df_tiers_filtered, df_qualite = pipeline(chemin, df_tiers, journal, df_IFRS, pnl, societe, mois, version)


