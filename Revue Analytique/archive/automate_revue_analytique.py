import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import sys
import os
import re
import datetime
import subprocess
import warnings
import traceback
from unidecode import unidecode
from pathlib import Path
from chardet import detect
import threading
import json # Added for configuration management

from openpyxl import Workbook
from openpyxl.chart import (LineChart, PieChart, BarChart3D, Reference)
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.text import Paragraph, ParagraphProperties, RegularTextRun,CharacterProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from assets.automate_reporting import insert_images_from_excel_to_word,config_elements_fixes,config_elements_variables


warnings.filterwarnings("ignore")

# --- Configuration Management ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller runtime temp directory
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

CONFIG_FILE = resource_path('assets/configuration.json')

LISTE_ENCODING = ["utf-8","UTF-16","MacRoman","SHIFT_JIS","CP932"]
JOURNAL_CODES_ACHATS = ["AC", "HA", "A10","60","ACH"]
JOURNAL_CODES_VENTES = ["VE", "VT", "VTE","70","VEN"]
FOURNISSEURS_CODE = ["401"]
CLIENTS_CODE = ["411"]
TAUX_CMC = ['0%','2.1%','5.5%','8.5%','10%','20%']

# Default configuration values
DEFAULT_CONFIG = {
    "liste_encoding":       LISTE_ENCODING,
    "fournisseurs_code":    [{"value": code, "selected": True} for code in FOURNISSEURS_CODE],
    "clients_code":         [{"value": code, "selected": True} for code in CLIENTS_CODE],
    "journal_codes_achats": [{"value": code, "selected": True} for code in JOURNAL_CODES_ACHATS],
    "journal_codes_ventes": [{"value": code, "selected": True} for code in JOURNAL_CODES_VENTES],
    'taux_cmc' :            TAUX_CMC
}

def load_config():
    """Loads configuration from configuration or returns default if not found."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                return loaded_config
        except json.JSONDecodeError:
            messagebox.showerror("Erreur de configuration", "Fichier configuration corrompu. Utilisation des paramètres par défaut.")
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config):
    """Saves the given configuration to configuration."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
    except IOError as e:
        messagebox.showerror("Erreur de sauvegarde", f"Impossible de sauvegarder la configuration: {e}")

# --- End Configuration Management ---

# Fonction de personnalisation du rounding
def custom_round(value):
    """
    Rounds a given value to the nearest integer, but rounds up if the decimal part is 0.5 or greater.
    """
    if value - np.floor(value) >= 0.5:
        return int(np.ceil(value))
    else:
        return int(np.floor(value))

def calculate_custom_cumsum(df, columns=None, add_total=False):    
    """
    Calculates a custom cumulative sum for the given DataFrame, which is a cumulative sum but rounds up
    if the decimal part is 0.5 or greater. The function also takes care of NaNs by ignoring them in the sum.
    An optional total row can be added at the end of the resulting DataFrame.

    :param df: The DataFrame to process.
    :param columns: The columns to process. If None, all columns except the first one are processed.
    :param add_total: If True, adds a total row at the end of the DataFrame with the cumulative sum of the last row.
    :return: A tuple containing the processed DataFrame and the last cumulative sum value, or None if multiple columns were processed.
    """
    result_df = df.copy()
    
    if columns is None:
        columns = df.columns[1:]
        
    last_values = {}
    
    for column in columns:
        values = result_df[column].to_list()
        result = []
        is_nan_beginning = True
        last_cumulative_value = 0

        for value in values:
            if pd.isna(value):
                result.append(0 if is_nan_beginning else custom_round(last_cumulative_value))
            else:
                is_nan_beginning = False
                last_cumulative_value += value
                result.append(custom_round(last_cumulative_value))

        result_df[column] = result
        last_values[column] = last_cumulative_value

    if add_total:
        total_row = ['Total'] + [last_values.get(col, 0) for col in columns]
        result_df.loc[len(result_df)] = total_row

    last_value = last_values[list(last_values.keys())[0]] if len(last_values) == 1 else None

    return result_df, last_value

def reorder_months(df):
    """
    Reorders the months in a given DataFrame to start from the first month found in the oldest year.

    The function takes a DataFrame as input and returns a list of 12 months in the correct order.
    If no valid year is found, the original order is returned.

    :param df: The DataFrame to process
    :return: A list of 12 months in the correct order
    """
    mois_en_ordre = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    # Get all unique years in ascending order
    years = sorted(df['Année'].unique())
    
    if len(years) == 0:
        return mois_en_ordre   # Return original order if no data found
    
    # Get the first month for the oldest year
    oldest_year = years[0]
    year_data = df[df['Année'] == oldest_year]
    # Convert Mois to categorical with the correct order
    year_data['Mois'] = pd.Categorical(year_data['Mois'], 
                                     categories=mois_en_ordre,
                                     ordered=True)
    
    # Now min() will respect our month order
    first_month = year_data['Mois'].min()
    
    if pd.isna(first_month):
        return mois_en_ordre  # Return original order if no valid month found
    
    # Find the index of the first month in the original list
    start_idx = mois_en_ordre.index(first_month)
    
    # Create new list starting from first_month
    reordered_months = mois_en_ordre[start_idx:] + mois_en_ordre[:start_idx]
    
    return reordered_months

# Chiffre d'affaires
def chiffres_affaires(df): 
    """
    Calcule les chiffres d'affaires mensuels et cumulés
    
    Parameters
    ----------
    df : pandas.DataFrame
        Le DataFrame contenant les données de la balance comptable
        
    Returns
    -------
    tuple of pandas.DataFrame
        Un tuple contenant deux DataFrames : le premier contient les chiffres d'affaires mensuels, le second contient les chiffres d'affaires cumulés
    """
    
    # Add this line at the start of the function
    mois_en_ordre = reorder_months(df)
    
    # Filtrer uniquement les comptes de chiffre d'affaires
    ca = df[df['CompteNum'].astype(str).str[:3].astype(int).between(700, 709)]
    # Calculer les montants nets (Débit - Crédit)
    ca['Montant'] = ca["Credit"] - ca["Debit"]

    # Appliquer l'ordre des mois
    ca['Mois'] = pd.Categorical(ca['Mois'], categories=mois_en_ordre, ordered=True)

    if len(ca['Année'].unique()) == 1:
        # Grouper par année et mois et calculer le total
        ca_mensuel = ca.groupby(['Année','Mois'])['Montant'].sum().reset_index()
        ca_mensuel = ca_mensuel.pivot(index="Mois", columns="Année", values="Montant")
        ca_mensuel= ca_mensuel.map(custom_round)
        ca_mensuel.columns.name = None
        ca_mensuel = ca_mensuel.reset_index()
    else:
        ca_mensuel = ca.groupby(['Mois'])['Montant'].sum().reset_index()
        ca_mensuel['Montant']= ca_mensuel['Montant'].map(custom_round)
        ca_mensuel.rename(columns={'Montant': '-'.join([str(annee) for annee in ca['Année'].unique()])}, inplace=True)

    ca_mensuel_cumule,total_cumule = calculate_custom_cumsum(ca_mensuel)
    ca_mensuel['Mois'] = pd.Categorical(ca_mensuel['Mois'], categories=mois_en_ordre, ordered=True) 

    colonne_annee = list(ca_mensuel.columns)[1]    
    ca_mensuel.loc[len(ca_mensuel)] = ['Total', ca_mensuel[colonne_annee].sum()]

    if len(df["Mois"].unique()) == 12:
        ca_mensuel[f"% {colonne_annee}"] = ca_mensuel.apply(lambda row: f'{round((row[colonne_annee]/ca_mensuel.loc[len(ca_mensuel)-1,colonne_annee])*100)}%',axis = 1)
        ca_mensuel.at[len(ca_mensuel)-1, f"% {colonne_annee}"] = "100%"
    ca_mensuel.rename(columns={colonne_annee:str(colonne_annee)}, inplace=True)
    
    if len(df["Mois"].unique()) == 12:
        ca_mensuel_cumule[f"% {colonne_annee}"] = ca_mensuel_cumule.apply(lambda row: f'{round((row[colonne_annee]/total_cumule)*100)}%',axis = 1)
    ca_mensuel_cumule.rename(columns={colonne_annee:str(colonne_annee)}, inplace=True)
    ca_mensuel_cumule.rename(columns={colonne_annee:str(colonne_annee)}, inplace=True)

    ca_mensuel_cumule['Mois'] = pd.Categorical(ca_mensuel_cumule['Mois'], categories=mois_en_ordre, ordered=True)

    return ca_mensuel ,ca_mensuel_cumule

# Achats et Sous-Traitance
def achats_et_sous_traitances(df):

    """
    Calculates monthly and cumulative totals for purchases and subcontracting.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing accounting data.

    Returns
    -------
    tuple of pandas.DataFrame
        A tuple containing two DataFrames:
        - The first DataFrame contains monthly totals for purchases and subcontracting.
        - The second DataFrame contains cumulative totals for purchases and subcontracting.
    """

    achats = df[df['CompteNum'].astype(str).str[:3].astype(int).between(601, 609)& ~df['CompteNum'].astype(str).str[:3].isin(['603','604','606'])]
    sous_traitance = df[df['CompteNum'].astype(str).str[:3].isin(['604','611'])]
    mois_en_ordre = reorder_months(df)
    
    def generer_table(data,type):
        data['Mois'] = pd.Categorical(data['Mois'], categories=mois_en_ordre, ordered=True)
        data[type] = data["Debit"] - data["Credit"]   
        data = data.groupby(['Mois'])[type].sum().reset_index()        

        data[type] = data[type].map(custom_round)

        data.loc[len(data)] = ['Total', data[type].sum()]

        return data
    
    table_achats = generer_table(achats,"Achats")
    table_sous_traitance = generer_table(sous_traitance,"Sous-traitance")

    table_achat_sous_traitance = pd.merge(table_achats,table_sous_traitance,on='Mois')
    table_achat_sous_traitance["Total"] = table_achat_sous_traitance["Achats"] + table_achat_sous_traitance["Sous-traitance"]

    table_achat_sous_traitance_cumule,_ = calculate_custom_cumsum(table_achat_sous_traitance[:-1])
    table_achat_sous_traitance_cumule = table_achat_sous_traitance_cumule[["Mois","Total"]]

    return table_achat_sous_traitance, table_achat_sous_traitance_cumule

# Investissement
def investissement(df,annee):
    """
    Generate a DataFrame with the total investment for each account between 20 and 26.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing accounting data.
    annee : int
        The year for which to calculate the total investment.

    Returns
    -------
    tuple of pandas.DataFrame
        A tuple containing two DataFrames:
        - The first DataFrame contains the total investment for each account between 20 and 26.
        - The second DataFrame contains the total investment for the given year.
    """
    
    data = df[df['CompteNum'].astype(str).str[:2].astype(int).between(20, 26)]
    data_an = df[df['CompteNum'].astype(str).str[:2].astype(int).between(20, 26)].loc[df['JournalCode'].str.strip().isin(['AA','AN'])]

    colonne_annee = str(annee)

    # Calculer les montants nets (Débit - Crédit)
    data['Solde'] =  data["Debit"] - data["Credit"]
    data = data.groupby(['CompteNum','CompteLib'])['Solde'].sum().reset_index()
    data['Solde']= data['Solde'].map(custom_round)
    data.rename(columns={'Solde': colonne_annee}, inplace=True)
    
    data_an['Solde'] =  data_an["Debit"] - data_an["Credit"]
    data_an = data_an.groupby(['CompteNum','CompteLib'])['Solde'].sum().reset_index()
    data_an['Solde']= data_an['Solde'].map(custom_round)
    data_an.rename(columns={'Solde': colonne_annee}, inplace=True)

    data[colonne_annee] = data[colonne_annee] - data_an[colonne_annee]
    total_investissement = pd.DataFrame([{"Année":colonne_annee,"Total":data[colonne_annee][:-1].sum()}])
    
    return data ,total_investissement      

# Trésorerie
def tresorerie(df):
    """
    Génère un tableau de la trésorerie mensuelle 

    Parameters
    ----------
    df : pandas.DataFrame
        Le DataFrame contenant les données de la balance comptable
                    
    Returns
    -------
    pandas.DataFrame
        Un DataFrame contenant la trésorerie mensuelle brute, nette des emprunts, et nette des emprunts et des comptes courants
    """

    # Add this line at the start of the function
    mois_en_ordre = reorder_months(df)

    comptes_bancaires = df[df['CompteNum'].astype(str).str.startswith('5')]
    emprunts = df[df['CompteNum'].astype(str).str.startswith('16')]
    compte_courants = df[df['CompteNum'].astype(str).str.startswith('45')]

    def generer_table_tresorerie(data, type):
        # Retirer les lignes avec des dates invalides
        data = data.dropna(subset=["EcritureDate"])
        
        # Calcul du solde
        data["Solde"] = data["Debit"] - data["Credit"]    

        # Grouper par Mois et CompteLib en gardant tous les mois
        tresorerie_mensuelle = data.groupby(["CompteLib", "Mois"], observed=False)["Solde"].sum().reset_index()

        # Créer un DataFrame contenant tous les mois et comptes possibles
        all_combinations = pd.MultiIndex.from_product([data["CompteLib"].unique(), mois_en_ordre], names=["CompteLib", "Mois"])

        tresorerie_mensuelle = tresorerie_mensuelle.set_index(["CompteLib", "Mois"]).reindex(all_combinations).reset_index()

        # Pivot pour aligner les mois
        tresorerie_mensuelle_pivot = tresorerie_mensuelle.pivot(index="Mois", columns="CompteLib", values="Solde").reindex(mois_en_ordre).reset_index()
        tresorerie_mensuelle_pivot.columns.name = ""

        tresorerie_mensuelle_pivot,_ = calculate_custom_cumsum(tresorerie_mensuelle_pivot)
        

        # Calcul du total
        total_label = {
            "comptes bancaires": "Trésorerie Brute",
            "emprunts": "Total des emprunts",
        }.get(type, "Total des C/C")

        tresorerie_mensuelle_pivot[total_label] = tresorerie_mensuelle_pivot.iloc[:, 1:].sum(axis=1)



        return tresorerie_mensuelle_pivot
    
    
    tableau_comptes_bancaires = generer_table_tresorerie(
        comptes_bancaires,
        'comptes bancaires',
    )
    tableau_emprunts = generer_table_tresorerie(
        emprunts,
        'emprunts',
    )
    tableau_compte_courants = generer_table_tresorerie(
        compte_courants,
        'comptes courants',
    )
    

    tableau_emprunts["Trésorerie nette des emprunts"] = tableau_emprunts["Total des emprunts"] + tableau_comptes_bancaires["Trésorerie Brute"]

    tableau_compte_courants["Trésorerie nette des emprunts et CC"] = tableau_compte_courants["Total des C/C"] + tableau_emprunts["Trésorerie nette des emprunts"]

    
    tableau_tresorie = tableau_comptes_bancaires.merge(tableau_emprunts,on="Mois", how='outer').merge(tableau_compte_courants,on="Mois", how='outer')
    tableau_tresorie["Mois"] = pd.Categorical(
        tableau_tresorie["Mois"],
        categories=mois_en_ordre,
        ordered=True,
    )
    tableau_tresorie = tableau_tresorie.dropna(subset=['Mois'])
    tableau_tresorie = tableau_tresorie.sort_values("Mois").reset_index(drop=True)

    tableau_tresorie.loc[:, tableau_tresorie.columns!="Mois"] = tableau_tresorie.loc[:, tableau_tresorie.columns!="Mois"].fillna(0)
    numeric_columns = tableau_tresorie.select_dtypes(include=["number"]).columns
    tableau_tresorie[numeric_columns] = tableau_tresorie[numeric_columns].map(custom_round)
  
    tableau_tresorie["Trésorerie nette des emprunts"] = tableau_tresorie["Total des emprunts"] + tableau_tresorie["Trésorerie Brute"]
    tableau_tresorie["Trésorerie nette des emprunts et CC"] = tableau_tresorie["Total des C/C"] + tableau_tresorie["Trésorerie nette des emprunts"]


        
    return tableau_tresorie

# Salaires
def frais_personnels(salaires,charges_sociales):
    """
    Renvoie un tableau avec l'évolution des frais de personnels
    (salaires et charges sociales) sur l'année.

    Parameters
    ----------
    salaires : pd.DataFrame
        Dataframe contenant la colonne 'Mois' et les colonnes 'Debit' et
        'Credit'
    charges_sociales : pd.DataFrame
        Dataframe contenant la colonne 'Mois' et les colonnes 'Debit' et
        'Credit'

    Returns
    -------
    table_frais_personnels : pd.DataFrame
        Tableau avec l'évolution des frais de personnels
    table_frais_personnels_cumules : pd.DataFrame
        Tableau avec l'évolution cumulative des frais de personnels
    """
    # Add this line at the start of the function
    mois_en_ordre = reorder_months(salaires)  # Using salaires df as it contains the necessary date info
    
    def generer_table_frais_personnels(data,type):

        data['Mois'] = pd.Categorical(data['Mois'], categories=mois_en_ordre, ordered=True)

        data[type] = data["Debit"] - data["Credit"]    

        frais_personnels = data.groupby(['Mois'])[type].sum().reset_index()
        frais_personnels[type] = frais_personnels[type].map(custom_round)
        frais_personnels.loc[len(frais_personnels)] = ['Total', frais_personnels[type].sum()]

        return frais_personnels

    table_salaire = generer_table_frais_personnels(salaires,
                                                   "Salaires")
    
    table_charges_sociales = generer_table_frais_personnels(charges_sociales,
                                                            "Charges sociales")
    
    table_frais_personnels = pd.merge(table_salaire,table_charges_sociales,on='Mois')

    table_frais_personnels["Total Salaires et Charges Sociales"] = table_frais_personnels["Salaires"] + table_frais_personnels["Charges sociales"]

    table_frais_personnels_cumules,_ = calculate_custom_cumsum(table_frais_personnels[["Mois", "Total Salaires et Charges Sociales"]][:-1])



    return table_frais_personnels,table_frais_personnels_cumules

# Top Fournisseurs/Clients
def top_dix(df, type, fournisseurs_codes=[], clients_codes=[], journal_code_achats = [], journal_code_ventes = []):
    """
    Génère un tableau des 10 principaux fournisseurs ou clients pour une année donnée.

    Parameters
    ----------
    df : pandas.DataFrame
        Le DataFrame contenant les données de la balance comptable.
    annee : int
        L'année pour laquelle le top 10 doit être calculé.
    type : Literal["Fournisseurs", "Clients"]
        Indique si le classement doit être fait pour les "Fournisseurs" ou les "Clients".
    show_title : bool, optional
        Si True, affiche le titre du tableau, par défaut False.

    Returns
    -------
    pd.DataFrame
        Un DataFrame contenant les 10 premiers fournisseurs ou clients, avec un résumé
        des autres et un total. Chaque ligne inclut le numéro et le libellé du compte auxiliaire,
        le solde et le pourcentage du total.
    """

    compte_num = fournisseurs_codes if type =="Fournisseurs" else clients_codes

    liste_journal_code = journal_code_achats + journal_code_ventes

    def solde(data,type):
        return data["Credit"] - data["Debit"] if type =="Clients" else data["Debit"] - data["Credit"]

    def classement_top_10(data,compte_num_column,compte_lib_column):
            
        classement = data.groupby([compte_num_column,compte_lib_column])['Solde'].sum().reset_index().sort_values(by='Solde',ascending = False)        
        classement["Solde"] = classement["Solde"].map(custom_round)
        classement_top_10 = classement.head(10)
        def contrepartie(row):
            contrepartie_data = data.loc[data[compte_lib_column] == row[compte_lib_column]]
            ecriture_num_ids = list(contrepartie_data["EcritureNum"].value_counts().keys())
            new_df = df.loc[(df["EcritureNum"].isin(ecriture_num_ids)) & (df['JournalCode'].str.strip().isin(liste_journal_code))]
            data_without_compte_num = new_df.loc[~new_df[compte_num_column].astype(str).str[:3].isin(compte_num + ['44'])]
            data_without_compte_num["Solde"] = solde(data_without_compte_num,type)
            grouped_df = data_without_compte_num.groupby("CompteNum")["Solde"].sum().to_frame().reset_index()
            return grouped_df[grouped_df["Solde"] == grouped_df["Solde"].max()].reset_index()["CompteNum"][0]
        
        classement_top_10["Contrepartie"] = classement_top_10.apply(lambda row: contrepartie(row),axis=1)
        classement_top_10_tail = pd.DataFrame([{compte_num_column:"",
                                                compte_lib_column:f'Autres {type}',
                                                "Solde":classement.iloc[10:]["Solde"].sum(),
                                                "Contrepartie":""
                                                },
                                               {compte_num_column:"Total",
                                                compte_lib_column:'',
                                                "Solde":classement["Solde"].sum(),
                                                "Contrepartie":""
                                                }
                                            ])
        classement_top_10_final = pd.concat([classement_top_10,classement_top_10_tail],axis = 0)
        classement_top_10_final["Pourcentage"] = classement_top_10_final.apply(lambda row: round((row["Solde"]/classement_top_10_final["Solde"].iloc[-1])*100),axis = 1)
        classement_top_10_final.set_index(compte_num_column, inplace=True)
        classement_top_10_final = classement_top_10_final.reset_index()
        classement_top_10_final = classement_top_10_final[[compte_num_column,compte_lib_column,"Solde","Pourcentage","Contrepartie"]]

        return classement_top_10_final,compte_lib_column

    
    data = df[(df["JournalCode"].str.strip().isin(liste_journal_code))].loc[df['CompteNum'].astype(str).str[:3].isin(compte_num)]

    if data.shape[0] != 0:

        if type =="Fournisseurs":
            data["Solde"] = data["Credit"] - data["Debit"]
        else:
            data["Solde"] = data["Debit"] - data["Credit"]

        try:

            classement_top_10_final,compte_lib_column = classement_top_10(data,
                                                                          compte_num_column = "CompAuxNum",
                                                                          compte_lib_column = "CompAuxLib"
                                                                          )

        except ValueError:

            classement_top_10_final,compte_lib_column = classement_top_10(data,
                                                                          compte_num_column = "CompteNum",
                                                                          compte_lib_column = "CompteLib"
                                                                          )

        return classement_top_10_final, compte_lib_column
    
    else:
        return pd.DataFrame([{"CompteNum":"","CompteLib":'',"Solde":''}]),None    

def top_dix_hors_taxes(df, type, fournisseurs_codes=[], clients_codes=[], journal_code_achats = [], journal_code_ventes = []):
    
    """
    Returns the top 10 'Fournisseurs' or 'Clients' by solde HT, excluding the ones with TVA (44*), with the corresponding 'Contrepartie' and 'Solde HT' values.
    The 'Solde HT' value is calculated by subtracting the sum of the 'Solde' values of the tail entries from the total 'Solde' value of the top 10 entries.
    The 'Contrepartie' value is the account number with the highest 'Solde' value among the tail entries.
    The 'Pourcentage HT' value is the percentage of the total 'Solde HT' value.
    The 'CompteNum' and 'CompteLib' columns are the account number and label of the top 10 entries.
    If the dataframe is empty, returns an empty dataframe with the same columns.

    Parameters:
    df (pd.DataFrame): The dataframe containing the data.
    type (str): 'Fournisseurs' or 'Clients'.
    fournisseurs_codes (list): The list of account numbers to consider for the 'Fournisseurs' type.
    clients_codes (list): The list of account numbers to consider for the 'Clients' type.
    journal_code_achats (list): The list of journal codes to consider for the 'Fournisseurs' type.
    journal_code_ventes (list): The list of journal codes to consider for the 'Clients' type.

    Returns:
    pd.DataFrame: The dataframe with the top 10 entries, with the 'CompteNum', 'CompteLib', 'Solde', 'Solde HT', 'Pourcentage HT' and 'Contrepartie' columns.
    str: The column name of the account label column.
    """
    compte_num,liste_journal_code = (fournisseurs_codes,journal_code_achats) if type =="Fournisseurs" else (clients_codes,journal_code_ventes)
    compte_num_columns = ["CompAuxNum","CompteNum"]
    compte_lib_columns = ["CompAuxLib","CompteLib"]

    def next_value_after(target,list):
        """
        Returns the next value after the first occurrence of `target` in the list `lst`.
        Returns None if the target is not found or if it is the last element.
        """
        try:
            idx = list.index(target)
            if idx < len(list) - 1:
                return list[idx + 1]
        except ValueError:
            pass  # target not in list
        return None
    
    def solde(data,type):
        return data["Credit"] - data["Debit"] if type =="Fournisseurs" else data["Debit"] - data["Credit"]
    
    def classement_top_10(data,type,compte_num ,compte_num_column,compte_lib_column):
        
        def has_tva_top_10(df,type,row,compte_num,compte_num_column= compte_num_column,compte_lib_column=compte_lib_column):
            
            data = df.loc[df[compte_lib_column] == row[compte_lib_column]]
            ecriture_num_ids = list(data["EcritureNum"].value_counts().keys())
            new_df = df.loc[(df["EcritureNum"].isin(ecriture_num_ids)) & (df['JournalCode'].str.strip().isin(liste_journal_code)) ]
            data_without_compte_num = new_df.loc[~new_df[compte_num_column].astype(str).str[:3].isin(compte_num + ['44'])]
            if new_df.equals(data_without_compte_num):
                data_without_compte_num = new_df.loc[~new_df[next_value_after(compte_num_column,compte_num_columns)].astype(str).str[:3].isin(compte_num + ['44'])]

            data_without_compte_num["Solde"] = - solde(data_without_compte_num,type)
            grouped_df = data_without_compte_num.groupby("CompteNum")["Solde"].sum().to_frame().reset_index()
            # Find the row with the max value
            contrepartie = grouped_df[grouped_df["Solde"] == grouped_df["Solde"].max()].reset_index()["CompteNum"][0] 
            montant_ht = custom_round(data_without_compte_num["Solde"].sum())
            
            # Si les deux montants sont différents ca veut dire qu'il y'a une TVA appliquée
            return pd.Series({"Solde HT":montant_ht,
                              "Contrepartie":contrepartie})
        

        def has_tva_tail(df, tail_data, type, compte_num, compte_num_column=compte_num_column, compte_lib_column=compte_lib_column):
            # Get all EcritureLib values for the tail entries
            tail_ecritures = []
            for lib in tail_data[compte_lib_column].unique():
                comp_data = df.loc[df[compte_lib_column] == lib]
                tail_ecritures.extend(list(comp_data["EcritureNum"].value_counts().keys()))
            
            # Filter the dataframe based on these EcritureLib values
            new_df = df.loc[(df["EcritureNum"].isin(tail_ecritures)) & (df['JournalCode'].str.strip().isin(liste_journal_code))]
            data_without_compte_num = new_df.loc[~new_df[compte_num_column].astype(str).str[:3].isin(compte_num + ['44'])]
            if new_df.equals(data_without_compte_num):
                data_without_compte_num = new_df.loc[~new_df[next_value_after(compte_num_column,compte_num_columns)].astype(str).str[:3].isin(compte_num + ['44'])]
            
            data_without_compte_num["Solde"] = -solde(data_without_compte_num, type)
            montant_ht = custom_round(data_without_compte_num["Solde"].sum())
            
            return montant_ht

        classement = data.groupby([compte_num_column,compte_lib_column])['Solde'].sum().reset_index().sort_values(by='Solde',ascending = False)
                
        classement["Solde"] = classement["Solde"].map(custom_round)        

        classement_top_10 = classement.head(10)
        classement_top_10[["Solde HT","Contrepartie"]] = classement_top_10.apply(lambda row : has_tva_top_10(df,type,row,compte_num),axis = 1)
        
        tail_data = classement.iloc[10:]
        solde_ht_autres = has_tva_tail(df,tail_data,type=type,compte_num=compte_num)
        
        classement_top_10_tail = pd.DataFrame([{compte_num_column: "",compte_lib_column:f'Autres {type}',"Solde":classement.iloc[10:]["Solde"].sum(),"Solde HT":solde_ht_autres,"Contrepartie":""},
                                               {compte_num_column:"Total",compte_lib_column:'',"Solde":classement["Solde"].sum(),"Solde HT": classement_top_10["Solde HT"].sum() +solde_ht_autres ,"Contrepartie":""}])
        classement_top_10_final = pd.concat([classement_top_10,classement_top_10_tail],axis = 0)
        classement_top_10_final["Pourcentage HT"] = classement_top_10_final.apply(lambda row: round((row["Solde HT"]/classement_top_10_final["Solde HT"].iloc[-1])*100),axis = 1)
        classement_top_10_final.set_index(compte_num_column, inplace=True)
        classement_top_10_final = classement_top_10_final.reset_index()
        classement_top_10_final = classement_top_10_final[[compte_num_column,compte_lib_column,"Solde","Solde HT","Pourcentage HT","Contrepartie"]]

        return classement_top_10_final,compte_lib_column
    
    data = df[(df["JournalCode"].astype(str).str.strip().isin(liste_journal_code))].loc[df['CompteNum'].astype(str).str[:3].isin(compte_num)]
        
    if data.empty:
        return pd.DataFrame([{"CompteNum": "", "CompteLib": '', "Solde": ''}]), None
        
    data["Solde"] = solde(data,type)

    try:
        classement_top_10_final,compte_lib_column = classement_top_10(data,type,compte_num,compte_num_column = compte_num_columns[0],compte_lib_column = compte_lib_columns[0])
    except ValueError:
        classement_top_10_final,compte_lib_column = classement_top_10(data,type,compte_num,compte_num_column = compte_num_columns[1],compte_lib_column = compte_lib_columns[1])

    return classement_top_10_final, compte_lib_column
            
# Crédit Moyen Client/Fournisseurs
def generer_table_soldes(df,type,fournisseurs_codes=[], clients_codes=[],table_solde_data = None,index_column = "Mois"):
    """
    Génère un tableau des soldes mensuels pour les fournisseurs ou les clients
    pour une année donnée.

    Parameters
    ----------
    df : pandas.DataFrame
        Le DataFrame contenant les données de la balance comptable.
    type : Literal["Fournisseurs", "Clients"]
        Indique si le calcul des soldes doit être fait pour les "Fournisseurs" ou les "Clients".

    Returns
    -------
    pandas.DataFrame
        Un DataFrame avec les soldes mensuels cumulés pour chaque mois de l'année,
        avec une ligne 'Total' à la fin.
    """

    mois_en_ordre = reorder_months(df)

    # On définit le numéro du compte pour les fournisseurs ou les clients
    compte_num = fournisseurs_codes if type =="Fournisseurs" else clients_codes
   
    data = table_solde_data.loc[table_solde_data['CompteNum'].astype(str).str[:3].isin(compte_num)] if table_solde_data is not None else df.loc[df['CompteNum'].astype(str).str[:3].isin(compte_num)]

    # On trie les données par mois
    data["Mois"] = pd.Categorical(data[index_column], categories=mois_en_ordre, ordered=True)
    
    data["Soldes"] = data["Credit"] - data["Debit"] if type =="Fournisseurs" else data["Debit"] - data["Credit"]

    # On calcule les soldes mensuels cumulés
    calcul_soldes = data.groupby([index_column])["Soldes"].sum().reset_index()
    calcul_soldes["Soldes"] = calcul_soldes["Soldes"].map(custom_round)

    # Pour toutes les colonnes sauf la première:
    calcul_soldes, total_solde = calculate_custom_cumsum(calcul_soldes,add_total=True)

    return calcul_soldes, total_solde

def generer_table_jours(tableau_source, type, previous_year_data=None):
    """
    Génère un tableau montrant l'évolution du crédit moyen clients/fournisseurs
    en prenant en compte les données de l'année précédente.

    Args:
        tableau_source: DataFrame contenant les données de la balance comptable de l'année courante
        type: ["Clients", "Fournisseurs"]
        previous_year_data: DataFrame contenant les données de l'année précédente (obligatoire)

    Returns:
        data: Un DataFrame contenant le crédit moyen mensuel
    """
    colonne_ttc = "Total TTC" if type == "Clients" else "Achats TTC"
    jours = []
    data = tableau_source.copy()
    
    # Si les données de l'année précédente sont fournies, on les copie et on retire la dernière ligne
    if previous_year_data is not None:
        prev_data = previous_year_data.copy()
        prev_data = prev_data[[colonne_ttc]].iloc[:-1]
    else:
        prev_data = None
    
    for i in range(len(data)):
        solde = data['Soldes'].iloc[i]
        achat = data[colonne_ttc].iloc[i]
        
        if solde > achat:
            somme_achats_precedents = achat
            k = i
            
            while k > 0 and somme_achats_precedents < solde:
                k -= 1
                somme_achats_precedents += data[colonne_ttc].iloc[k]
            
            if somme_achats_precedents >= solde:
                somme_achats_precedents -= data[colonne_ttc].iloc[k]
                jours_valeur = ((solde - somme_achats_precedents) / data[colonne_ttc].iloc[k]) * 30 + 30 * (i - k)
            
            # Si on atteint la première ligne et que le solde n'est pas couvert, on passe à l'année précédente
            elif prev_data is not None:
                k_prev = len(prev_data) - 1
                while k_prev >= 0 and somme_achats_precedents < solde:
                    somme_achats_precedents += prev_data[colonne_ttc].iloc[k_prev]
                    k_prev -= 1
                
                if somme_achats_precedents >= solde:
                    somme_achats_precedents -= prev_data[colonne_ttc].iloc[k_prev + 1]
                    jours_valeur = ((solde - somme_achats_precedents) / prev_data[colonne_ttc].iloc[k_prev + 1]) * 30 + 30 * (len(prev_data) - k_prev - 1 + i)
                else:
                    jours_valeur = 0
            else:
                jours_valeur = 0
        else:
            jours_valeur = (solde / achat) * 30 if achat != 0 else 0
        
        jours.append(jours_valeur)
    
    # On ajoute les jours à la colonne 'Jours' du DataFrame
    data['Jours'] = jours
    # On sélectionne les colonnes 'Mois' et 'Jours'
    data = data[["Mois", "Jours"]]
    # On remplace les valeurs manquantes par 0
    data = data.fillna(0)
    # On arrondit les valeurs de la colonne 'Jours'
    data['Jours'] = data['Jours'].map(custom_round)
    
    return data

def credit_moyen_client(df, annee, previous_year_data=None, account_config=None,fournisseurs_codes=[], clients_codes=[]):
    """
    Génère un tableau montrant l'évolution du crédit moyen clients
    pour chaque mois de l'année spécifiée, avec une ligne 'Total' à la fin.
    Les colonnes sont regroupées par taux de TVA.
    
    Args:
        df: DataFrame contenant les données de la balance comptable
        annee: L'année pour laquelle générer le tableau
        previous_year_data: Optional data from previous year for calculations
        account_config: Configuration of accounts to include and their VAT rates,
                        expected to be the full config structure with years as keys

    Returns:
        table_cmc: Un DataFrame contenant le crédit moyen clients mensuel
    """
    mois_en_ordre = reorder_months(df)
    
    # If account_config is provided and has data for the specified year, use it
    if account_config and annee in account_config and account_config[annee]:
        # Extract the configuration for the specified year
        year_config = account_config[annee]
        
        # Create an empty DataFrame with months to ensure all months are represented
        base_df = pd.DataFrame({'Mois': mois_en_ordre})
        base_df['Mois'] = pd.Categorical(base_df['Mois'], categories=mois_en_ordre, ordered=True)
        
        # Create DataFrames for each VAT rate
        vat_rates_data = {}
        total_ttc_df = base_df.copy()
        total_ttc_df['Total TTC'] = 0.0  # Initialize with zeros
        
        # First pass: identify unique VAT rates in the config
        unique_vat_rates = set()
        for compte_lib, config in year_config.items():
            if config['selected']:  # Only process selected accounts
                unique_vat_rates.add(config['tva'])
        
        # Initialize DataFrames for each VAT rate
        for rate in unique_vat_rates:
            rate_str = f"{int(rate * 100) if rate*100 == int(rate*100) else rate*100}%"
            column_name = f"Ventes {rate_str}"
            vat_rates_data[rate] = base_df.copy()
            vat_rates_data[rate][column_name] = 0.0
        
        # Process each account and add to the appropriate VAT rate DataFrame
        for compte_lib, config in year_config.items():
            if config['selected']:  # Only process selected accounts
                compte_num = config['compte_num']
                tva_rate = config['tva']
                
                # Filter data for this account
                account_data = df[df['CompteNum'].astype(str) == str(compte_num)]
                
                if not account_data.empty:
                    # Calculate for this account
                    account_data['Mois'] = pd.Categorical(account_data['Mois'], categories=mois_en_ordre, ordered=True)
                    account_data['Value'] = account_data["Credit"] - account_data["Debit"]
                    account_monthly = account_data.groupby(['Mois'])['Value'].sum().reset_index()
                    
                    # Ensure all months are present
                    account_monthly = pd.merge(base_df, account_monthly, on='Mois', how='left')
                    account_monthly['Value'] = account_monthly['Value'].fillna(0).map(custom_round)
                    
                    # Define the VAT rate column name
                    rate_str = f"{int(tva_rate * 100) if tva_rate*100 == int(tva_rate * 100) else tva_rate * 100}%"
                    column_name = f"Ventes {rate_str}"
                    
                    # Add to the corresponding VAT rate DataFrame
                    vat_rates_data[tva_rate]['Mois'] = pd.Categorical(vat_rates_data[tva_rate]['Mois'], 
                                                                      categories=mois_en_ordre, ordered=True)
                    vat_rates_data[tva_rate][column_name] += account_monthly['Value']
                    
                    # Calculate and add TTC amount
                    ttc_value = account_monthly['Value'] * (1 + tva_rate)
                    ttc_value = ttc_value.map(custom_round)
                    
                    # Add to total TTC
                    total_ttc_df['Total TTC'] += ttc_value
        
        # Combine all VAT rate DataFrames
        table_cmc = base_df.copy()
        
        # Add each VAT rate column
        vat_columns = []
        for rate, rate_df in vat_rates_data.items():
            rate_str = f"{int(rate * 100) if rate*100 == int(rate*100) else rate*100}%"
            column_name = f"Ventes {rate_str}"
            vat_columns.append(column_name)
            table_cmc = pd.merge(table_cmc, rate_df[['Mois', column_name]], on='Mois', how='left')
            table_cmc[column_name] = table_cmc[column_name].map(custom_round)
        
        # Add Total TTC column
        table_cmc = pd.merge(table_cmc, total_ttc_df[['Mois', 'Total TTC']], on='Mois', how='left')
        table_cmc['Total TTC'] = table_cmc['Total TTC'].map(custom_round)
        
        # Get table_soldes for clients
        table_soldes,total_soldes = generer_table_soldes(df, 'Clients',fournisseurs_codes=fournisseurs_codes, clients_codes=clients_codes)
        table_cmc = table_cmc.merge(table_soldes, on='Mois', how='left')
                
        # Generate and merge tables_jours
        tables_jours = generer_table_jours(table_cmc, 'Clients', previous_year_data=previous_year_data)
        tables_jours = tables_jours[tables_jours['Mois'] != 'Total']  # Skip the total row
        table_cmc = table_cmc.merge(tables_jours, on='Mois', how='left')
        
        # Ensure all numeric columns have zeros instead of NaN
        for col in table_cmc.columns:
            if col != 'Mois':
                table_cmc[col] = table_cmc[col].fillna(0)
                if col not in ['Soldes', 'Jours']:  # Skip rounding for specific columns if needed
                    table_cmc[col] = table_cmc[col].map(custom_round)
        
        # Add total row
        total_row = {'Mois': 'Total'}
        for col in table_cmc.columns:
            if col != 'Mois':
                # Skip calculation for columns that shouldn't be summed
                if col not in ['Jours']:
                    if col == 'Soldes':
                        total_row[col] = total_soldes
                    else:
                        total_row[col] = table_cmc[col].sum()
                else:
                    total_row[col] = None  # or some other appropriate value
        
        # Append total row properly
        table_cmc = pd.concat([table_cmc, pd.DataFrame([total_row])], ignore_index=True)
        
        # Ensure proper column order
        desired_columns = ['Mois'] + vat_columns + ['Total TTC', 'Soldes', 'Jours']
        # Filter to only include columns that exist
        desired_columns = [col for col in desired_columns if col in table_cmc.columns]
        table_cmc = table_cmc[desired_columns]
        
        return table_cmc
        
    # If no custom config or no accounts selected, use the original method
    ventes = df[df['CompteNum'].astype(str).str[:3].astype(int).between(700, 709) & 
                ~df['CompteNum'].astype(str).str.startswith('706')]
    prestations = df[df['CompteNum'].astype(str).str.startswith('706')]

    def generer_cmc(data, type):
        data['Mois'] = pd.Categorical(data['Mois'], categories=mois_en_ordre, ordered=True)
        data[type] = data["Credit"] - data["Debit"]    
        cmc = data.groupby(['Mois'])[type].sum().reset_index()
        cmc[type] = cmc[type].map(custom_round)
        cmc.loc[len(cmc)] = ['Total', cmc[type].sum()]
        
        return cmc

    table_ventes = generer_cmc(ventes, "Ventes de marchandises")
    table_prestations = generer_cmc(prestations,"Prestation de services")
    
    table_soldes,_ = generer_table_soldes(df, 'Clients',fournisseurs_codes=fournisseurs_codes, clients_codes=clients_codes)    

    table_cmc = pd.merge(table_prestations, table_ventes, on='Mois')

    table_cmc["Total TTC"] = ((table_cmc["Prestation de services"] + 
                                  table_cmc["Ventes de marchandises"]) * 1.2).map(custom_round)
    table_cmc = table_cmc.merge(table_soldes, on='Mois', how='left')

    tables_jours = generer_table_jours(table_cmc, 'Clients')
    table_cmc = table_cmc.merge(tables_jours, on='Mois', how='left')
    
        
    return table_cmc

def credit_moyen_client_top_cinq(df, client, annee, previous_year_data=None, account_config=None,fournisseurs_codes=[], clients_codes=[]):
    """
    Génère un tableau montrant l'évolution du crédit moyen clients
    pour chaque mois de l'année spécifiée pour un client appartenant au top 5, avec une ligne 'Total' à la fin.
    Les colonnes sont regroupées par taux de TVA.
    
    Args:
        df: DataFrame contenant les données de la balance comptable
        client : Le nom du client
        annee: L'année pour laquelle générer le tableau
        previous_year_data: Optional data from previous year for calculations
        account_config: Configuration of accounts to include and their VAT rates,
                        expected to be the full config structure with years as keys

    Returns:
        table_cmc: Un DataFrame contenant le crédit moyen clients mensuel
    """
    client_data = df.loc[df["CompAuxLib"] == client]
    ecriture_lib_ids = list(client_data["EcritureNum"].value_counts().keys())
    new_df = df.loc[df["EcritureNum"].isin(ecriture_lib_ids)]
    table_solde_data = new_df
    title = f'Crédit moyen client {client} {annee}'

    mois_en_ordre = reorder_months(df)
    
    # If account_config is provided and has data for the specified year, use it
    if account_config and annee in account_config and account_config[annee]:
        # Extract the configuration for the specified year
        year_config = account_config[annee]
        
        # Create an empty DataFrame with months to ensure all months are represented
        base_df = pd.DataFrame({'Mois': mois_en_ordre})
        base_df['Mois'] = pd.Categorical(base_df['Mois'], categories=mois_en_ordre, ordered=True)
        
        # Create DataFrames for each VAT rate
        vat_rates_data = {}
        total_ttc_df = base_df.copy()
        total_ttc_df['Total TTC'] = 0.0  # Initialize with zeros
        
        # First pass: identify unique VAT rates in the config
        unique_vat_rates = set()
        for compte_lib, config in year_config.items():
            if config['selected']:  # Only process selected accounts
                unique_vat_rates.add(config['tva'])
        
        # Initialize DataFrames for each VAT rate
        for rate in unique_vat_rates:
            rate_str = f"{int(rate * 100) if rate*100 == int(rate*100) else rate*100}%"
            column_name = f"Ventes {rate_str}"
            vat_rates_data[rate] = base_df.copy()
            vat_rates_data[rate][column_name] = 0.0
        
        # Process each account and add to the appropriate VAT rate DataFrame
        for compte_lib, config in year_config.items():
            if config['selected']:  # Only process selected accounts
                compte_num = config['compte_num']
                tva_rate = config['tva']
                
                # Filter data for this account
                account_data = new_df[new_df['CompteNum'].astype(str) == str(compte_num)]
                
                if not account_data.empty:
                    # Calculate for this account
                    account_data['Mois'] = pd.Categorical(account_data['Mois'], categories=mois_en_ordre, ordered=True)
                    account_data['Value'] = account_data["Credit"] - account_data["Debit"]
                    account_monthly = account_data.groupby(['Mois'])['Value'].sum().reset_index()
                    
                    # Ensure all months are present
                    account_monthly = pd.merge(base_df, account_monthly, on='Mois', how='left')
                    account_monthly['Value'] = account_monthly['Value'].fillna(0).map(custom_round)
                    
                    # Define the VAT rate column name
                    rate_str = f"{int(tva_rate * 100) if tva_rate*100 == int(tva_rate * 100) else tva_rate * 100}%"
                    column_name = f"Ventes {rate_str}"
                    
                    # Add to the corresponding VAT rate DataFrame
                    vat_rates_data[tva_rate]['Mois'] = pd.Categorical(vat_rates_data[tva_rate]['Mois'], 
                                                                      categories=mois_en_ordre, ordered=True)
                    vat_rates_data[tva_rate][column_name] += account_monthly['Value']
                    
                    # Calculate and add TTC amount
                    ttc_value = account_monthly['Value'] * (1 + tva_rate)
                    ttc_value = ttc_value.map(custom_round)
                    
                    # Add to total TTC
                    total_ttc_df['Total TTC'] += ttc_value
        
        # Combine all VAT rate DataFrames
        table_cmc = base_df.copy()
        
        # Add each VAT rate column
        vat_columns = []
        for rate, rate_df in vat_rates_data.items():
            rate_str = f"{int(rate * 100) if rate*100 == int(rate*100) else rate*100}%"
            column_name = f"Ventes {rate_str}"
            vat_columns.append(column_name)
            table_cmc = pd.merge(table_cmc, rate_df[['Mois', column_name]], on='Mois', how='left')
            table_cmc[column_name] = table_cmc[column_name].map(custom_round)
        
        # Add Total TTC column
        table_cmc = pd.merge(table_cmc, total_ttc_df[['Mois', 'Total TTC']], on='Mois', how='left')
        table_cmc['Total TTC'] = table_cmc['Total TTC'].map(custom_round)
        
        # Get table_soldes for clients
        table_soldes,total_soldes = generer_table_soldes(df, 
                                                         'Clients',
                                                         fournisseurs_codes=fournisseurs_codes, 
                                                         clients_codes=clients_codes,
                                                         table_solde_data=table_solde_data)
        table_cmc = table_cmc.merge(table_soldes, on='Mois', how='left')
                
        # Generate and merge tables_jours
        tables_jours = generer_table_jours(table_cmc, 'Clients', previous_year_data=previous_year_data)
        tables_jours = tables_jours[tables_jours['Mois'] != 'Total']  # Skip the total row
        table_cmc = table_cmc.merge(tables_jours, on='Mois', how='left')
        
        # Ensure all numeric columns have zeros instead of NaN
        for col in table_cmc.columns:
            if col != 'Mois':
                table_cmc[col] = table_cmc[col].fillna(0)
                if col not in ['Soldes', 'Jours']:  # Skip rounding for specific columns if needed
                    table_cmc[col] = table_cmc[col].map(custom_round)
        
        # Add total row
        total_row = {'Mois': 'Total'}
        for col in table_cmc.columns:
            if col != 'Mois':
                # Skip calculation for columns that shouldn't be summed
                if col not in ['Jours']:
                    if col == 'Soldes':
                        total_row[col] = total_soldes
                    else:
                        total_row[col] = table_cmc[col].sum()
                else:
                    total_row[col] = None  # or some other appropriate value
        
        # Append total row properly
        table_cmc = pd.concat([table_cmc, pd.DataFrame([total_row])], ignore_index=True)
        
        # Ensure proper column order
        desired_columns = ['Mois'] + vat_columns + ['Total TTC', 'Soldes', 'Jours']
        # Filter to only include columns that exist
        desired_columns = [col for col in desired_columns if col in table_cmc.columns]
        table_cmc = table_cmc[desired_columns]
        
        return (title,table_cmc)
        
    # If no custom config or no accounts selected, use the original method
    ventes = df[df['CompteNum'].astype(str).str[:3].astype(int).between(700, 709) & 
                ~df['CompteNum'].astype(str).str.startswith('706')]
    prestations = df[df['CompteNum'].astype(str).str.startswith('706')]

    def generer_cmc(data, type):
        data['Mois'] = pd.Categorical(data['Mois'], categories=mois_en_ordre, ordered=True)
        data[type] = data["Credit"] - data["Debit"]    
        cmc = data.groupby(['Mois'])[type].sum().reset_index()
        cmc[type] = cmc[type].map(custom_round)
        cmc.loc[len(cmc)] = ['Total', cmc[type].sum()]
        
        return cmc

    table_ventes = generer_cmc(ventes, "Ventes de marchandises")
    table_prestations = generer_cmc(prestations,"Prestation de services")
    
    table_soldes,_ = generer_table_soldes(df, 'Clients',fournisseurs_codes=fournisseurs_codes, clients_codes=clients_codes)    

    table_cmc = pd.merge(table_prestations, table_ventes, on='Mois')

    table_cmc["Total TTC"] = ((table_cmc["Prestation de services"] + 
                                  table_cmc["Ventes de marchandises"]) * 1.2).map(custom_round)
    table_cmc = table_cmc.merge(table_soldes, on='Mois', how='left')

    tables_jours = generer_table_jours(table_cmc, 'Clients')
    table_cmc = table_cmc.merge(tables_jours, on='Mois', how='left')
    
        
    return (title,table_cmc)

def credit_moyen_fournisseur(df,previous_year_data=None,fournisseurs_codes=[], clients_codes=[],journal_code_achats=[]):    
    """
    Génère un tableau pivot montrant l'évolution du crédit moyen fournisseurs pour chaque mois de l'année spécifiée, avec une ligne 'Total' à la fin.
    
    Args:
        df: DataFrame contenant les données de la balance comptable

    Returns:
        table_cmf: Un DataFrame contenant le crédit moyen fournisseurs mensuel
    """
    # Ajouter cette ligne au début de la fonction
    mois_en_ordre = reorder_months(df)

    achats_ttc = df[(df["JournalCode"].str.strip().isin(journal_code_achats))]

    achats_ttc['Mois'] = pd.Categorical(achats_ttc['Mois'], categories=mois_en_ordre, ordered=True)

    achats_ttc['Achats TTC'] = achats_ttc["Debit"]    

    tables_achats = achats_ttc.groupby(['Mois'])['Achats TTC'].sum().reset_index()        

    tables_achats['Achats TTC'] = tables_achats['Achats TTC'].map(custom_round)


    table_soldes,total_solde = generer_table_soldes(df,
                                        type = 'Fournisseurs',fournisseurs_codes=fournisseurs_codes, clients_codes=clients_codes)
    
    table_cmf = pd.merge(tables_achats,table_soldes,on='Mois')

    tables_jours = generer_table_jours(table_cmf,'Fournisseurs',previous_year_data=previous_year_data)
    table_cmf = table_cmf.merge(tables_jours,on = 'Mois')
    # Add the Total row back at the end
    total_row = {'Mois': 'Total'}
    for col in table_cmf.columns:
        if col!='Mois':
            if col =='Achats TTC':
                total_row[col] = table_cmf[col].sum()
            elif col == 'Soldes':
                total_row[col] = total_solde
            else:
                total_row[col] = None  # or some other appropriate value
    
    table_cmf = pd.concat([table_cmf, pd.DataFrame([total_row])], ignore_index=True)

    return table_cmf

def credit_moyen_fournisseur_top_cinq(df,annee,fournisseur,previous_year_data=None,fournisseurs_codes=[], clients_codes=[],journal_code_achats=[]):    
    """
    Génère un tableau pivot montrant l'évolution du crédit moyen fournisseurs pour chaque mois de l'année spécifiée, avec une ligne 'Total' à la fin.
    
    Args:
        df: DataFrame contenant les données de la balance comptable

    Returns:
        table_cmf: Un DataFrame contenant le crédit moyen fournisseurs mensuel
    """
    # Ajouter cette ligne au début de la fonction
    mois_en_ordre = reorder_months(df)
    fournisseur_data = df.loc[df["CompAuxLib"] == fournisseur]
    piece_ref_ids = list(fournisseur_data["EcritureNum"].value_counts().keys())
    new_df = df.loc[df["EcritureNum"].isin(piece_ref_ids)]
    table_solde_data = fournisseur_data
    title = f'Crédit moyen fournisseur {fournisseur} {annee}'
    
    achats_ttc = new_df[(new_df["JournalCode"].str.strip().isin(journal_code_achats))]

    achats_ttc['Mois'] = pd.Categorical(achats_ttc['Mois'], categories=mois_en_ordre, ordered=True)

    achats_ttc['Achats TTC'] = achats_ttc["Debit"]    

    tables_achats = achats_ttc.groupby(['Mois'])['Achats TTC'].sum().reset_index()        

    tables_achats['Achats TTC'] = tables_achats['Achats TTC'].map(custom_round)


    table_soldes,total_solde = generer_table_soldes(df,
                                                    'Fournisseurs',
                                                    table_solde_data=table_solde_data,fournisseurs_codes=fournisseurs_codes, clients_codes=clients_codes)
    
    table_cmf = pd.merge(tables_achats,table_soldes,on='Mois')

    tables_jours = generer_table_jours(table_cmf,'Fournisseurs',previous_year_data=previous_year_data)
    table_cmf = table_cmf.merge(tables_jours,on = 'Mois')
    # Add the Total row back at the end
    total_row = {'Mois': 'Total'}
    for col in table_cmf.columns:
        if col!='Mois':
            if col =='Achats TTC':
                total_row[col] = table_cmf[col].sum()
            elif col == 'Soldes':
                total_row[col] = total_solde
            else:
                total_row[col] = None  # or some other appropriate value
    
    table_cmf = pd.concat([table_cmf, pd.DataFrame([total_row])], ignore_index=True)

    return title,table_cmf

# Fonctions de mise en forme
def adjust_column_width(ws):
    """
    Dynamically adjust column widths based on cell content.
    
    Args:
        ws: The worksheet to adjust
    
    Adjusts column width starting from the second column and second row.
    Stops when more than 3 empty columns and more than 11 empty rows are encountered.
    """

    # Track empty columns and rows
    empty_columns_count = 0
    empty_rows_count = 0
    
    # Start from the second column (index 1) and second row (index 1)
    for col_idx, column_cells in enumerate(ws.iter_cols(min_row=2, min_col=2), start=2):
        column_empty = True
        column_max_width = 10  # Default minimum width
        
        row_index = 2  # Start at second row
        for cell in column_cells:
            # Skip cells that are completely empty
            if cell.value is None or str(cell.value).strip() == '':
                empty_rows_count += 1
                row_index += 1
                continue
            
            # Reset empty rows count when a non-empty cell is found
            empty_rows_count = 0
            column_empty = False
            
            # Calculate width based on cell content
            cell_content = str(cell.value)
            cell_width = max(len(cell_content), 10)  # Minimum width of 10
            
            # Adjust for bold or other formatting
            if cell.font and cell.font.bold:
                cell_width += 2
            
            column_max_width = max(column_max_width, cell_width)
            row_index += 1
        
        # Track empty columns
        if column_empty:
            empty_columns_count += 1
        else:
            empty_columns_count = 0
            # Store column width
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = column_max_width
        
        # Stop if we've encountered too many empty columns or rows
        if empty_columns_count > 3 and empty_rows_count > 11:
            break
    
def color_between_columns(ws, sheet_name,df_width,start_column):
    """
    Colore les cellules entre deux colonnes identifiées par leurs noms d'en-tête.
    
    Args:
        ws: La feuille de calcul
        sheet_name: Le nom de la feuille de calcul
        df_width: La largeur du DataFrame
        start_column: La colonne de départ du DataFrame
    """

    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    darker_blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    pink_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    lighter_green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    darker_pink_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    purple_fill = PatternFill(start_color="E5E5FF", end_color="E5E5FF", fill_type="solid")

    row_to_color_attributes = {
        "Trésorerie": [["Mois", "Trésorerie Brute", blue_fill,False],
                       ["Trésorerie Brute", "Trésorerie nette des emprunts", pink_fill,False],
                       ["Trésorerie nette des emprunts", "Trésorerie nette des emprunts et CC", green_fill,False]
                       ],
        "Crédit Moyen CLT": [["Total", purple_fill,True],
                             ["Moyenne", purple_fill,True]],
        "Crédit Moyen CLT TOP 5": [["Total", purple_fill,True],
                             ["Moyenne", purple_fill,True]
                             ],        
        "Crédit Moyen FR": [["Total", purple_fill,True],
                            ["Moyenne", purple_fill,True]
                            ],
        "Crédit Moyen FR TOP 5": [["Total", purple_fill,True],
                             ["Moyenne", purple_fill,True]
                             ],
        "Salaires": [["Total", darker_pink_fill,True]
                     ],
        "CA": [["Total", darker_blue_fill,True]
               ],
        "TOP CLT": [["Total", lighter_green_fill,True]
                    ],
        "TOP FR": [["Total", lighter_green_fill,True]
                   ]
    }

    if sheet_name in list(row_to_color_attributes.keys()):
        value_to_find = row_to_color_attributes[sheet_name][0][0]
        
        # Trouver toutes les lignes d'en-tête
        rows_to_colour = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
            for cell in row:
                if cell.value == value_to_find:
                    rows_to_colour.append(row_idx)

        # Pour chaque ligne d'en-tête trouvée
        for row_idx in rows_to_colour:        
            for row_to_color in row_to_color_attributes[sheet_name]:
                first_col = None
                second_col = None
                if sheet_name == "Trésorerie":
                    start_col = row_to_color[0]
                    end_col = row_to_color[1]
                    fill_color = row_to_color[2]
                    include_start_col_and_end_col = row_to_color[3]
                else:
                    start_col = row_to_color[0]
                    end_col = str(ws.cell(row=row_idx, column=start_column + df_width - 1).value).strip()
                    if end_col == "None" or end_col == "":
                        end_col = str(ws.cell(row=row_idx, column=start_column + df_width - 2).value).strip()
                        
                    fill_color = row_to_color[1]
                    include_start_col_and_end_col = row_to_color[2]
                
                # Obtenir la ligne contenant l'en-tête
                header_row = ws[row_idx]
                
                # Trouver les colonnes dans cette table spécifique
                for cell in header_row:
                    cell_value = str(cell.value).strip() if cell.value else ""
                    if cell_value == start_col:
                        first_col = cell.column
                    if cell_value == end_col:
                        second_col = cell.column
                
                # Si nous avons trouvé les deux colonnes pour cette table
                if first_col and second_col and first_col < second_col:
                    first_cell = ws.cell(row=row_idx, column=first_col)
                    first_cell_fill_color = first_cell.fill.start_color.rgb
                    if first_cell_fill_color != fill_color.start_color.rgb:
                        if include_start_col_and_end_col:                        
                            for col in range(first_col, second_col + 1):
                                cell = ws.cell(row=row_idx, column=col)
                                cell.fill = fill_color
                        else:                      
                            for col in range(first_col + 1, second_col):
                                cell = ws.cell(row=row_idx, column=col)
                                cell.fill = fill_color

def color_columns(ws, sheet_name, current_row, last_row):
    """
    Colore les colonnes d'un DataFrame en fonction du nom de la feuille et de la largeur du DataFrame.

    Args:
        ws: La feuille de calcul
        sheet_name: Le nom de la feuille de calcul
        current_row: La ligne de départ du DataFrame
        last_row: La ligne de fin du DataFrame
    """
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    purple_fill = PatternFill(start_color="E5E5FF", end_color="E5E5FF", fill_type="solid")
    darker_blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    
    column_to_color_attributes = {"CA":[["%"],darker_blue_fill],
        "Trésorerie": [["Trésorerie Brute", "Trésorerie nette des emprunts", "Trésorerie nette des emprunts et CC"], gray_fill],
        "Crédit Moyen CLT": [["Jours"], purple_fill],
        "Crédit Moyen CLT TOP 5": [["Jours"], purple_fill],
        "Crédit Moyen FR": [["Jours"], purple_fill],
        "Crédit Moyen FR TOP 5": [["Jours"], purple_fill],
        "TOP CLT": [["Pourcentage"], green_fill],
        "TOP FR": [["Pourcentage"], green_fill]
    }  

    if sheet_name in list(column_to_color_attributes.keys()):
        header_row = ws[current_row]  # Ligne d'en-tête de la table actuelle
        columns_to_color = column_to_color_attributes[sheet_name][0]
        fill_color = column_to_color_attributes[sheet_name][1]
        for cell in header_row:
            if cell.value in columns_to_color:
                first_cell = ws.cell(row=current_row, column=cell.column)
                first_cell_fill_color = first_cell.fill.start_color.rgb
                if first_cell_fill_color != fill_color.start_color.rgb:
                    # Colorier de l'en-tête à la dernière ligne de la table actuelle
                    for row in range(current_row, last_row + 1):
                        target_cell = ws.cell(row=row, column=cell.column)
                        target_cell.fill = fill_color
            else:
                if sheet_name == "CA":
                    if cell.value and "%" in cell.value:
                        first_cell = ws.cell(row=current_row, column=cell.column)
                        first_cell_fill_color = first_cell.fill.start_color.rgb
                        if first_cell_fill_color != fill_color.start_color.rgb:
                            # Colorier de l'en-tête à la dernière ligne de la table actuelle
                            for row in range(current_row, last_row + 1):
                                target_cell = ws.cell(row=row, column=cell.column)
                                target_cell.fill = fill_color

def color_title_cell(ws, title, sheet_name):
    """
    Colore la cellule de titre en fonction du nom de la feuille et ajuste sa largeur de colonne.
    
    Args:
        ws: Feuille Excel
        title: Contenu de la cellule à trouver
        sheet_name: Nom de la feuille pour déterminer la couleur
    """
    # Définir les couleurs pour différentes feuilles
    title_colors = {
        "CA": "BDD7EE",  # Blue
        "Trésorerie": "FFE699",  # Orange
        "Crédit Moyen CLT": "E5E5FF",
        "Crédit Moyen CLT TOP 5": "E5E5FF",  # Purple
        "Crédit Moyen FR": "E5E5FF",  # Purple
        "Crédit Moyen FR TOP 5": "E5E5FF",
        "TOP CLT": "C6E0B4",  # Green
        "TOP FR": "C6E0B4",  # Green
        "Salaires": "F8CBAD",  # DarkPink
        "Investissement" : "DDEBF7" # Light Blue
    }
    
    # Trouver la cellule de titre
    if sheet_name in list(title_colors.keys()):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == title:
                    # Obtenir la couleur pour la feuille actuelle (par défaut à blanc si la feuille n'est pas dans le dictionnaire)
                    color = title_colors.get(sheet_name, "FFFFFF")
                    
                    # Appliquer la couleur
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    cell.font = Font(bold=True, color="000000")  # Texte noir en gras
                    
                    # Ajuster la largeur de colonne en fonction de la longueur du titre
                    col_letter = get_column_letter(cell.column)
                    ws.column_dimensions[col_letter].width = len(str(title)) + 2
                    
                    return

# Fonctions de création de graphiques
def plot_data(ws, df, df_start_row, start_column, x_axis_col, value_cols, chart_type="line", percentage_col=None, remove_total=False, title=""):
    """
    Crée un graphique dans une feuille Excel.

    Args:
        ws (Worksheet): La feuille Excel.
        df (pd.DataFrame): Le DataFrame.
        df_start_row (int): La ligne de départ du DataFrame.
        start_column (int): La colonne de départ.
        x_axis_col (str): La colonne de l'axe des abscisses. 
        value_cols (list): Les colonnes de valeurs.
        chart_type (str): Le type de graphique: ["line", "pie", "barchart"]
        percentage_col (str): La colonne de la valeur en pourcentage.
        remove_total (bool): Si True, le total est supprimé.
        title (str): Le titre du graphique.
    
    Returns:
        chart: Le graphique.
    """
    if "Mois" in df.columns and "Total" in df["Mois"].values:
        remove_total = True
        
    if remove_total:
        df = df.iloc[:-1].copy()

    # Increased chart dimensions for better spacing
    chart_height = 10
    chart_width = 20
    
    if chart_type == "line":
        
        chart = LineChart()
        
        # Définir le titre dans le graphique avec le formatage approprié
        chart.title = title
        paragraph = Paragraph(
            pPr=ParagraphProperties(),
            r=[RegularTextRun(rPr=CharacterProperties(sz=1400, b=True), t=title)]
        )
        chart.title.tx.rich.paragraphs = [paragraph]        
        # Apply layout to prevent title from overlapping with chart
        chart.title.layout = Layout(
            manualLayout=ManualLayout(
                x=0.01,  # Position from left edge (1% from left)
                y=0.01,  # Position from top edge (1% from top)
                h=0.1    # Height of title area (10% of chart height)
            )
        )
        
        chart.height = chart_height
        chart.width = chart_width
        chart.style = 12
        
        # Configurer les axes avec plus d'espace
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.numFmt = '0'
        chart.x_axis.title = None
        chart.y_axis.title = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = None
        
        # Position the legend at the bottom with proper spacing
        chart.legend.position = "b"
        
        # Configure plot area to leave space for title and legend
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.15,    # 15% from left
                y=0.15,   # 15% from top (allowing space for title)
                h=0.7,    # Height of plot area (70% of chart)
                w=0.7,    # Width of plot area (70% of chart)
                xMode="edge",
                yMode="edge",
            )
        )

        # Obtenir les indices de colonne corrects
        mois_col_idx = start_column + list(df.columns).index(x_axis_col)        
        min_col = start_column + list(df.columns).index(value_cols[0])
        max_col = start_column + list(df.columns).index(value_cols[-1])

        data = Reference(
            ws,
            min_col=min_col,
            max_col=max_col,
            min_row=df_start_row,
            max_row=df_start_row + len(df)
        )

        chart.add_data(data, titles_from_data=True)

        # Définir les catégories de l'axe des x (mois) après avoir ajouté les séries
        categories = Reference(ws, 
                            min_col=mois_col_idx, 
                            min_row=df_start_row + 1,  # Passer l'en-tête pour les catégories
                            max_row=df_start_row + len(df))
        

        # If there's only one series, remove the legend
        if min_col == max_col:  
            chart.legend = None
        elif len(value_cols) == 2 and '%' in "".join([str(value_col) for value_col in value_cols]):
            chart.legend = None

        # Définir les couleurs personnalisées pour chaque série
        colors = ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', 'FF0000', '00FF00', '0000FF', 'FFA500', '800080']
        
        # Remove percentage series (in reverse to not affect indexes)
        for i in range(len(chart.series)-1, -1, -1):
            series = chart.series[i]
            # Check if series.tx.v exists and contains percentage indicators
            if i < len(value_cols) and ("Pourcentage" in str(value_cols[i]) or "%" in str(value_cols[i])):
                del chart.series[i]

        # Style remaining series
        for i, series in enumerate(chart.series):
            series.graphicalProperties.line.solidFill = colors[i % len(colors)]
            series.graphicalProperties.line.width = 28350
            series.smooth = False

        chart.set_categories(categories)
        
        return chart

    elif chart_type == "barchart":
        
        chart = BarChart3D()

        chart.type = "bar"
        # Définir le titre dans le graphique avec le formatage approprié
        chart.title = title
        paragraph = Paragraph(
            pPr=ParagraphProperties(),
            r=[RegularTextRun(rPr=CharacterProperties(sz=1400, b=True), t=title)]
        )
        chart.title.tx.rich.paragraphs = [paragraph]
        
        # Apply layout to prevent title from overlapping with chart
        chart.title.layout = Layout(
            manualLayout=ManualLayout(
                x=0.01,  # Position from left edge (1% from left)
                y=0.01,  # Position from top edge (1% from top)
                h=0.1    # Height of title area (10% of chart height)
            )
        )
        
        chart.height = chart_height
        chart.width = chart_width
        chart.style = 12
        
        # Configurer les axes
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.numFmt = '0'
        chart.x_axis.title = None
        chart.y_axis.title = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        
        # Position the legend to the right with proper spacing
        chart.legend.position = "b"
        # Configure plot area to leave space for title and legend
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.15,    # 15% from left
                y=0.15,   # 15% from top (allowing space for title)
                h=0.7,    # Height of plot area (70% of chart)
                w=0.7,    # Width of plot area (70% of chart)
                xMode="edge",
                yMode="edge",
            )
        )

        # Obtenir les indices de colonne corrects
        cat_col_idx = start_column + list(df.columns).index(x_axis_col)        
        min_col = start_column + list(df.columns).index(value_cols[0])
        max_col = start_column + list(df.columns).index(value_cols[-1])

        data = Reference(
            ws,
            min_col=min_col,
            max_col=max_col,
            min_row=df_start_row,
            max_row=df_start_row + len(df)
        )

        chart.add_data(data, titles_from_data=True)

        # Définir les catégories de l'axe des x après avoir ajouté les séries
        categories = Reference(ws, 
                            min_col=cat_col_idx, 
                            min_row=df_start_row + 1,  # Passer l'en-tête pour les catégories
                            max_row=df_start_row + len(df))
        chart.set_categories(categories)

        colors = ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '70AD47', '4472C4', 'C00000']
        for i, series in enumerate(chart.series):
            series.graphicalProperties.solidFill = colors[i % len(colors)]        
        # Explicitly set series names from column headers
        if min_col == max_col:  
            chart.legend = None

        return chart
    
    elif chart_type == "pie":

        if not percentage_col or percentage_col not in df.columns:
            raise ValueError("A valid percentage column is required for pie charts.")

        chart = PieChart()
        
        # Définir le titre avec le formatage approprié et la position
        chart.title = title
        paragraph = Paragraph(
            pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1400, b=True)),
            r=[RegularTextRun(t=title)]
        )
        chart.title.tx.rich.paragraphs = [paragraph]
        chart.title.overlay = False  # Ensure title doesn't overlap with chart
        
        # Position the title properly
        chart.title.layout = Layout(
            manualLayout=ManualLayout(
                x=0.01,  # Position from left edge (1% from left)
                y=0.01,  # Position from top edge (1% from top)
                h=0.1    # Height of title area
            )
        )
        
        chart.height = chart_height
        chart.width = chart_width

        # Configure data labels
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showCatName = False
        chart.dataLabels.showVal = False
        chart.dataLabels.showSerName = False
        
        # Configure plot area to leave space for title and legend
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.15,    # 15% from left
                y=0.15,   # 15% from top (allowing space for title)
                h=0.7,    # Height of plot area (70% of chart)
                w=0.7,    # Width of plot area (70% of chart)
                xMode="edge",
                yMode="edge",
            )
        )

        # Définition des références Excel
        data_col_idx = start_column + df.columns.get_loc(percentage_col)
        label_col_idx = start_column + df.columns.get_loc(x_axis_col)

        data_ref = Reference(ws, min_col=data_col_idx, min_row=df_start_row + 1, max_row=df_start_row + len(df))
        label_ref = Reference(ws, min_col=label_col_idx, min_row=df_start_row + 1, max_row=df_start_row + len(df))

        chart.add_data(data_ref, titles_from_data=False)
        chart.set_categories(label_ref)

        return chart

    else:
        raise ValueError("Unsupported chart type. Use 'line' or 'pie'.")

# Fonction de création de DataFrames et de graphiques
def add_df_to_sheet(wb, sheet_name, *data_frames_with_titles, start_column=2, plot=False, x_axis="Mois", chart_type="line", percentage_col=None, remove_total=True,spacing = 2):
    """
    Adds DataFrames as Excel tables and their corresponding charts to an Excel sheet.
    Includes error handling and safeguards for file opening.
    """
    # Create or get worksheet
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]
        # Clear existing tables to avoid conflicts
        tables_to_delete = list(ws.tables.keys())  # Create a list of table names
        for table_name in tables_to_delete:
            del ws.tables[table_name]  # Delete tables from the dictionary

    current_row = 2
    
    CHART_ROWS = 20  # Chart height in rows
    SPACING = spacing  # Number of empty rows between sections
    used_table_names = set()  # Track used table names

    for title, df in data_frames_with_titles:
        # Skip empty DataFrames
        if df.empty or df.iloc[0].isna().all():
            continue

        # Add and color title
        ws.cell(row=current_row, column=start_column, value=title)
        color_title_cell(ws, title, sheet_name)
        current_row += 1
        df_start_row = current_row

        # Write DataFrame headers
        for col_idx, column in enumerate(df.columns, start=start_column):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = str(column)  # Convert to string to avoid type issues
            cell.font = Font(bold=True)

        ignore_sheetnames = ["TOP CLT","TOP FR"]
            
        # Write DataFrame data
        for row_idx, row in enumerate(df.values, start=current_row + 1):
            for col_idx, value in enumerate(row, start=start_column):
                cell = ws.cell(row=row_idx, column=col_idx)
                # Handle different types of zero values and None
                if sheet_name not in ignore_sheetnames :
                    if value in (0, "0", 0.0, None, ""):
                        cell.value = ""
                    else:
                        cell.value = value
                else : 
                    cell.value = value

        # Create unique table name
        base_table_name = re.sub(r'[^a-zA-Z0-9_]', '_', unidecode(title))  # Remove special characters
        table_name = base_table_name
        counter = 1
        while table_name in used_table_names:
            table_name = f"{base_table_name}_{counter}"
            counter += 1
        used_table_names.add(table_name)

        # Create Excel table with error handling
        try:
            tab = Table(
                displayName=table_name,
                ref=f"{get_column_letter(start_column)}{df_start_row}:"
                    f"{get_column_letter(start_column + len(df.columns) - 1)}{df_start_row + len(df)}",
                tableStyleInfo=TableStyleInfo(
                    name=None,  # Use a built-in style for better compatibility
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
            )
            ws.add_table(tab)
        except Exception as e:
            print(f"Warning: Could not create table for {title}: {str(e)}")

        last_row = current_row + len(df)
        
        # Apply conditional coloring
        color_between_columns(ws, sheet_name, len(df.columns), start_column)    
        color_columns(ws, sheet_name, current_row, last_row)
        
        current_row += len(df)
        value_cols = [col for col in df.columns if col != x_axis]
        if sheet_name == "Investissement":
            value_cols = [col for col in df.columns[2:]]

        if list(df) == ['Année', 'Total']:
            plot=False
            
        if plot:
            try:
                # Generate and add chart
                chart = plot_data(
                    ws=ws,
                    df=df,
                    df_start_row=df_start_row,
                    start_column=start_column,
                    x_axis_col=x_axis,
                    value_cols=value_cols,
                    chart_type=chart_type,
                    percentage_col=percentage_col,
                    remove_total=remove_total,
                    title=title
                )

                # Calculate safe chart position
                max_data_column = start_column + len(df.columns)
                chart_col = max_data_column + 2
                chart_cell = f"{get_column_letter(chart_col)}{df_start_row}"
                
                # Ensure chart doesn't overlap with data
                ws.add_chart(chart, chart_cell)

                # Update current_row considering chart height
                chart_end_row = df_start_row + CHART_ROWS
                current_row = max(current_row, chart_end_row)
            except Exception as e:
                print(f"Warning: Could not create chart for {title}: {str(e)}")

        # Add spacing after each section
        current_row += SPACING

    # Adjust column widths at the end
    try:
        adjust_column_width(ws)
    except Exception as e:
        print(f"Warning: Could not adjust column widths: {str(e)}")


# --- Settings Window Class ---
class SettingsWindow(tk.Toplevel):
    def __init__(self, master, current_config, update_main_app_callback,open_settings_window_callback,has_imported_config = False):
        super().__init__(master)
        self.title("Paramètres par défaut")
        self.geometry("450x600") # Made window smaller
        self.parametres = current_config.copy() # Work with a copy until saved
        self.update_main_app_callback = update_main_app_callback
        self.open_settings_window_callback = open_settings_window_callback
        self.treeviews = {} # To store Treeview widgets for each list
        self.list_entry_widgets = {} # To store Entry widgets for adding new items

        self.create_widgets()
        self.load_settings_into_ui()
        self._set_unsaved_changes_state(has_imported_config)

    def create_widgets(self):
        # Create a canvas and scrollbar for the main content
        self.canvas = tk.Canvas(self, borderwidth=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create a frame inside the canvas to hold all the settings
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas_frame_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", tags="self.scrollable_frame")

        # Configure canvas scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel) # For Windows/Mac scroll
        self.canvas.bind_all("<Button-4>", self._on_mousewheel) # For Linux scroll
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

        # Fournisseurs and Clients codes (updated to match list format logic)
        param_frame = ttk.LabelFrame(self.scrollable_frame, padding="10")
        param_frame.pack(fill="x", pady=5)

        self.default_button = ttk.Button(param_frame, text="Retourner aux valeurs par défaut", command=self.return_to_default)
        self.default_button.pack(side="right", expand=True, padx=5)

        self._create_encoding_section(self.scrollable_frame, "Liste des encodages", "liste_encoding", is_int=False)
        
        # Journal Codes sections
        self._create_list_section(self.scrollable_frame, "Comptes Fournisseurs",
                                  "fournisseurs_code", is_int=False)
        self._create_list_section(self.scrollable_frame, "Comptes Clients",
                                  "clients_code", is_int=False)        
        self._create_list_section(self.scrollable_frame, "Codes Journal Achats",
                                  "journal_codes_achats", is_int=False)
        self._create_list_section(self.scrollable_frame, "Codes Journal Ventes",
                                  "journal_codes_ventes", is_int=False)

        # Action buttons
        button_frame = ttk.Frame(self.scrollable_frame, padding="10")
        button_frame.pack(fill="x", pady=10)

        self.save_button = ttk.Button(button_frame, text="Sauvegarder", command=self.save_settings,state="disabled")
        self.save_button.pack(side="left", expand=True, padx=5)
        
        is_new_default_config = tk.BooleanVar(value=False)        
        # Checkbox to set as default configuration
        # This checkbox will allow the user to set the current configuration as the default
        new_default_config_cb = ttk.Checkbutton(
            button_frame,
            text="Définir comme configuration\npar défaut",
            variable=is_new_default_config
            ,command=lambda: self._enable_save_button(is_new_default_config.get())
        )
        new_default_config_cb.pack(side="left", padx=5)
        self.save_button.config(command=lambda: self.save_settings(is_new_default_config.get()))

        self.cancel_button = ttk.Button(button_frame, text="Annuler", command=self.destroy,state="disabled")
        self.cancel_button.pack(side="right", expand=True, padx=5)

        import_export_frame = ttk.Frame(self.scrollable_frame, padding="10")
        import_export_frame.pack(fill="x", pady=10)

        # Ajoutez ces lignes pour les boutons Import/Export :
        self.import_button = ttk.Button(import_export_frame, text="Importer Configuration", command=self.import_config)
        self.import_button.pack(side="left",expand=True, padx=5)

        self.export_button = ttk.Button(import_export_frame, text="Exporter Configuration", command=self.export_config)
        self.export_button.pack(side="right", expand=True,padx=5)

    def show_info_popup(self):
        if self.info_popup is not None:
            return

        self.info_popup = tk.Toplevel(self)
        self.info_popup.overrideredirect(True)
        self.info_popup.configure(bg="#ffffff", padx=10, pady=5)
        self.info_popup.wm_attributes("-topmost", 1)

        # Position the popup near the icon
        x = self.info_icon.winfo_rootx() + 10
        y = self.info_icon.winfo_rooty() + 20
        self.info_popup.geometry(f"+{x}+{y}")

        label = ttk.Label(
            self.info_popup,
            text="Par exemple, entrer '2' signifie que les données des 2 mois avant et 2 mois après seront incluses.\nSi le mois de clôture est Décembre, les périodes seront donc Novembre-Décembre et Janvier-Février.",
            background="#ffffff",
            justify="left",
            font=("Segoe UI", 9),
            wraplength=200
        )
        label.pack()

    def close_info_popup(self, event=None):
        if self.info_popup is not None:
            self.info_popup.destroy()
            self.info_popup = None

    def on_canvas_configure(self, event):
        """Resizes the scrollable frame to match the canvas width."""
        self.canvas.itemconfig(self.canvas_frame_id, width=event.width)

    def _enable_save_button(self,new_default_config_checked=False):
        """Enables the save button and cancel button when changes are made."""
        if new_default_config_checked and not self.save_button['state'] == "normal":
            self.save_button.config(state="normal")

    def _on_mousewheel(self, event):
        """Handles mouse wheel scrolling."""
        if sys.platform.startswith('win'):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif sys.platform == 'darwin':
            self.canvas.yview_scroll(int(-1 * event.delta), "units")
        else: # Linux
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

    def _create_encoding_section(self, parent_frame, title, config_key, is_int=False):
        """Helper to create a section for list parameters with a single 'Value' column."""
        list_frame = ttk.LabelFrame(parent_frame, text=title, padding="10")
        list_frame.pack(fill="x", pady=5)

        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill="both", expand=True, pady=(0, 5))

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        
        # Treeview with ONLY 'Value' column
        self.treeviews[config_key] = ttk.Treeview(
            tree_frame,
            columns=("Value",),     # Tuple with single column
            show="headings",
            yscrollcommand=scrollbar.set,
            height=3
        )
        scrollbar.config(command=self.treeviews[config_key].yview)

        # Configure the single 'Value' column
        self.treeviews[config_key].heading("Value", text="Valeur")
        self.treeviews[config_key].column("Value", stretch=tk.YES)

        self.treeviews[config_key].pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Input and buttons
        input_frame = ttk.Frame(list_frame)
        input_frame.pack(fill="x", pady=5)

        new_entry = ttk.Entry(input_frame, width=30)
        new_entry.pack(side="left", expand=True, padx=5)
        self.list_entry_widgets[config_key] = new_entry

        def _add_item_to_list(self, config_key, entry_widget, is_int=False):
            new_value_str = entry_widget.get().strip()
            if not new_value_str:
                return

            try:
                new_value = int(new_value_str) if is_int else str(new_value_str)
            except ValueError:
                messagebox.showerror("Erreur de saisie", "Veuillez entrer un nombre entier valide pour ce champ.")
                return

            # Check for duplicates
            if new_value in self.parametres[config_key]:
                messagebox.showinfo("Duplicata", "Cette valeur existe déjà dans la liste.")
                return

            self.parametres[config_key].append(new_value)
            self.treeviews[config_key].insert("", "end", values=(new_value,))
            entry_widget.delete(0, tk.END)

            self._set_unsaved_changes_state(True)


        def _remove_item_from_list(self, config_key):
            selected_tree_items = self.treeviews[config_key].selection()
            if not selected_tree_items:
                messagebox.showinfo("Sélection manquante", "Veuillez sélectionner un élément à supprimer.")
                return

            values_to_remove = []
            for item_id in selected_tree_items:
                # Get the 'value' from the Treeview item (single column now)
                value_to_remove = self.treeviews[config_key].item(item_id, "values")[0]
                values_to_remove.append(value_to_remove)
                self.treeviews[config_key].delete(item_id)  # Remove from Treeview immediately

            # Update the actual config list (remove by value)
            self.parametres[config_key] = [v for v in self.parametres[config_key] if v not in values_to_remove]
            self._set_unsaved_changes_state(True)

        add_button = ttk.Button(
            input_frame,
            text="Ajouter",
            command=lambda: _add_item_to_list(config_key, new_entry)
        )
        add_button.pack(side="left", padx=5)

        remove_button = ttk.Button(
            input_frame,
            text="Supprimer",
            command=lambda: _remove_item_from_list(config_key)
        )
        remove_button.pack(side="left", padx=5)

    def _create_list_section(self, parent_frame, title, config_key, is_int=False):
        """Helper to create a section for list parameters with a 'Selected' column."""
        list_frame = ttk.LabelFrame(parent_frame, text=title, padding="10")
        list_frame.pack(fill="x", pady=5)

        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill="both", expand=True, pady=(0,5))

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        self.treeviews[config_key] = ttk.Treeview(tree_frame, columns=("Selected", "Value"), show="headings",
                                         yscrollcommand=scrollbar.set, height=3) # Set height to 5 rows
        scrollbar.config(command=self.treeviews[config_key].yview)

        self.treeviews[config_key].heading("Selected", text="Actif")
        self.treeviews[config_key].column("Selected", stretch=tk.NO, minwidth=0, width=50, anchor="center") # Fixed width for "Selected"
        self.treeviews[config_key].heading("Value", text="Valeur")
        self.treeviews[config_key].column("Value", stretch=tk.YES)

        self.treeviews[config_key].pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind click event on the "Selected" column to toggle state
        self.treeviews[config_key].bind("<Button-1>", lambda event, key=config_key: self._on_selected_column_click(event, key))

        input_frame = ttk.Frame(list_frame)
        input_frame.pack(fill="x", pady=5)

        new_entry = ttk.Entry(input_frame, width=30)
        new_entry.pack(side="left", expand=True, padx=5)
        self.list_entry_widgets[config_key] = new_entry

        add_button = ttk.Button(input_frame, text="Ajouter",
                                command=lambda: self._add_item_to_list(config_key, new_entry, is_int))
        add_button.pack(side="left", padx=5)

        remove_button = ttk.Button(input_frame, text="Supprimer",
                                   command=lambda: self._remove_item_from_list(config_key))
        remove_button.pack(side="left", padx=5)

    def _on_selected_column_click(self, event, config_key):
        """Handles clicks on the 'Selected' column to toggle state."""
        col = self.treeviews[config_key].identify_column(event.x)
        if col == "#1": # Column #1 is "Selected" (0-indexed for internal use, but #1 for display)
            item_id = self.treeviews[config_key].identify_row(event.y)
            if item_id:
                # Get the value from the item's 'Value' column to identify the config entry
                item_value = self.treeviews[config_key].item(item_id, "values")[1] # Get value from second column

                # Find the corresponding dictionary in self.parametres and toggle 'selected'
                for item_dict in self.parametres[config_key]:
                    if item_dict["value"] == item_value:
                        item_dict["selected"] = not item_dict["selected"]
                        # Update the Treeview display for the 'Selected' column
                        self.treeviews[config_key].item(item_id, values=(self._get_selected_display(item_dict["selected"]), item_value))
                        self._set_unsaved_changes_state(True)
                        break

    def _get_selected_display(self, selected):
        """Returns a string representation for the 'Selected' status (e.g., '✓' or '✗')."""
        return "✓" if selected else "✗"

    def load_settings_into_ui(self):
        """Populates the UI with current configuration values, including selected state."""
        for key, treeview in self.treeviews.items():
            # Clear existing items
            for item in treeview.get_children():
                treeview.delete(item)
            # Insert current config items with 'Selected' status
            for item in self.parametres.get(key, []):
                if type(item) == dict:  # If item is a dict with 'value' and 'selected'
                    treeview.insert("", "end", values=(self._get_selected_display(item["selected"]), item["value"]))
                else:  # If item is a simple value
                    treeview.insert("", "end", values=(item,))  # Default selected state

    def _add_item_to_list(self, config_key, entry_widget, is_int):
        new_value_str = entry_widget.get().strip()
        if not new_value_str:
            return

        try:
            new_value = int(new_value_str) if is_int else str(new_value_str)
        except ValueError:
            messagebox.showerror("Erreur de saisie", "Veuillez entrer un nombre entier valide pour ce champ.")
            return

        # Check for duplicates based on the 'value'
        if any(item_dict["value"] == new_value for item_dict in self.parametres[config_key]):
            messagebox.showinfo("Duplicata", "Cette valeur existe déjà dans la liste.")
            return

        new_item_dict = {"value": new_value, "selected": True} # New items are selected by default
        self.parametres[config_key].append(new_item_dict)
        self.treeviews[config_key].insert("", "end", values=(self._get_selected_display(True), new_value))
        entry_widget.delete(0, tk.END)

        self._set_unsaved_changes_state(True) # Enable save/cancel buttons after changes

    def _remove_item_from_list(self, config_key):
        selected_tree_items = self.treeviews[config_key].selection()
        if not selected_tree_items:
            messagebox.showinfo("Sélection manquante", "Veuillez sélectionner un élément à supprimer.")
            return

        values_to_remove = []
        for item_id in selected_tree_items:
            # Get the 'value' from the Treeview item (second column)
            value_to_remove = self.treeviews[config_key].item(item_id, "values")[1]
            values_to_remove.append(value_to_remove)
            self.treeviews[config_key].delete(item_id) # Remove from Treeview immediately

        # Update the actual config list
        self.parametres[config_key] = [item_dict for item_dict in self.parametres[config_key] if item_dict["value"] not in values_to_remove]
        self._set_unsaved_changes_state(True) # Enable save/cancel buttons after changes

    def _set_unsaved_changes_state(self, has_unsaved_changes):
        """Enables/disables save, cancel, and default buttons based on whether there are unsaved changes."""
        state = "normal" if has_unsaved_changes else "disabled"
        self.save_button.config(state=state)
        self.cancel_button.config(state=state)
        # self.default_button.config(state=state)

    def return_to_default(self):
        """Resets the settings to the default configuration."""
        self.parametres = load_config()
        self.load_settings_into_ui()
        self.update_main_app_callback(self.parametres) # Update the main app's config
        self._set_unsaved_changes_state(False)

    def import_config(self):
        file_path = filedialog.askopenfilename(
            title="Importer la configuration",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_config = json.load(f)
                
                # Update the main app's config first
                self.update_main_app_callback(imported_config) 
                
                messagebox.showinfo("Importation réussie", "La configuration a été importée avec succès. La fenêtre des paramètres va se rafraîchir.")
                
                # Destroy the current settings window so the main app can create a fresh one
                self.destroy() 
                self.open_settings_window_callback(has_imported_config=True)
                
            except Exception as e:
                messagebox.showerror("Erreur d'importation", f"Impossible d'importer la configuration : {e}")
                print(e)

    def export_config(self):
        file_path = filedialog.asksaveasfilename(
            title="Exporter la configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.parametres, f, indent=4)
                messagebox.showinfo("Exportation réussie", "La configuration a été exportée avec succès.")
            except Exception as e:
                messagebox.showerror("Erreur d'exportation", f"Impossible d'exporter la configuration : {e}")

    def save_settings(self,set_as_default=False):
        """Saves settings from UI to the configuration and updates main app."""
        try:
            if set_as_default:
                # Save the current parameters as the default configuration
                save_config(self.parametres)
            self.update_main_app_callback(self.parametres) # Update the main app's config
            messagebox.showinfo("Sauvegarde", "Paramètres sauvegardés avec succès.")
            self.destroy()
        except ValueError:
            messagebox.showerror("Erreur de saisie", "Veuillez entrer des nombres entiers valides pour les codes.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur inattendue est survenue lors de la sauvegarde: {e}")


class FECProcessor(tk.Tk):
    """
    Classe pour le traitement des fichiers FEC
    """
    # Initialisation de la classe
    def __init__(self):
        super().__init__()

        self.title("Automatisation de la revue analytique")
        self.geometry("1000x750")
        self.minsize(1000, 750)
        self.center_main_window()
        self.parametres = load_config() # Load configuration at startup
        self._settings_window = None # Initialize to None
        
        # Liste pour stocker les fichiers FEC et leurs années
        self.fec_files = []  # Liste de tuples (année, chemin_fichier)
        self.entry_widgets = {}  # Dictionnaire pour stocker les widgets Entry pour les années
        self.error_details = None  # Store error details
        self.output_path = ''
        
        # Configuration de l'interface
        self.create_widgets()

    # Création des widgets
    def create_widgets(self):
        """
        Crée les widgets de l'interface.
        """
        # Frame principale avec padding
        main_frame = tk.Frame(self)
        main_frame.pack(expand=True, fill='both', padx=40, pady=40)

        # Create a frame to hold the settings button and info icon
        top_right_frame = tk.Frame(main_frame)
        top_right_frame.place(relx=1.0, rely=0.0, anchor='ne', x=-30, y=10) # Position top-right

        # Settings Icon/Button
        settings_button = ttk.Button(
            top_right_frame,
            text="⚙️ Paramètres",
            command=self.open_settings_window
        )
        settings_button.pack(side=tk.RIGHT, padx=(5, 0)) # Pack to the right within its frame

        self.info_icon = tk.Label(
            top_right_frame,
            text="❓",
            font=("Segoe UI", 12),
            foreground="#00080E",  # Microsoft blue
            background=top_right_frame.cget('bg'), # Use the frame's background
            cursor="arrow",
            bd=0,
            highlightthickness=0,
            padx=0,
            pady=0
        )
        self.info_icon.pack(side=tk.RIGHT) # Pack to the right, next to the button

        # Bind hover events
        self.info_icon.bind("<Enter>", self.show_info_popup)
        self.info_icon.bind("<Leave>", self.close_info_popup)

        # Initialize popup as None
        self.info_popup = None

        # Configurer la grille
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Label titre
        title_label = tk.Label(
            main_frame, 
            text="Automate de revue analytique",
            font=('Inter', 20, 'bold')
        )
        title_label.grid(row=0, column=0, pady=30)

        # Frame pour l'ajout de fichiers
        add_frame = tk.Frame(main_frame)
        add_frame.grid(row=1, column=0, sticky='nsew', pady=10)
        add_frame.grid_columnconfigure(0, weight=1)

        # Label expliquant le format des années
        year_format_label = tk.Label(
            main_frame,
            text="Vous pouvez rentrer les années au format suivant : ANNEE (ex: 2025) ou ANNEE N-1-ANNEE N (ex: 2024-2025)",
            font=('Inter', 10, 'italic')
        )
        year_format_label.grid(row=2, column=0, pady=5)

        # Bouton pour sélectionner plusieurs fichiers FEC
        self.select_files_button = tk.Button(
            add_frame,
            text="Sélectionner les fichiers FEC",
            command=self.add_fec_files,
            font=('Inter', 11),
        )
        self.select_files_button.grid(row=0, column=0, padx=5)


        # Treeview pour afficher les fichiers sélectionnés avec années
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Inter", 8))
        style.configure("Treeview", font=("Inter", 10))

        self.files_tree = ttk.Treeview(
            main_frame, 
            columns=('Année', 'Fichier'),
            show='headings',
            height=10,
            selectmode='extended'  # ← allows multiple selections
        )
        self.files_tree.grid(row=3, column=0, pady=20, sticky='nsew')

        # Configuration des colonnes
        self.files_tree.heading('Année', text='Année')
        self.files_tree.heading('Fichier', text='Fichier')
        self.files_tree.column('Année', width=100)
        self.files_tree.column('Fichier', width=500)

        # Scrollbar pour le Treeview
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        scrollbar.grid(row=3, column=1, sticky='ns')
        self.files_tree.configure(yscrollcommand=scrollbar.set)

        # Frame pour les boutons de suppression
        delete_frame = tk.Frame(main_frame)
        delete_frame.grid(row=4, column=0, pady=10)

        # Bouton pour supprimer un fichier
        self.delete_button = tk.Button(
            delete_frame,
            text="Supprimer la sélection",
            command=self.delete_selected,
            font=('Inter', 11),
        )
        self.delete_button.pack(side='left', padx=5)

        # Bouton pour supprimer tous les fichiers
        self.delete_all_button = tk.Button(
            delete_frame,
            text="Supprimer tout",
            command=self.delete_all_files,
            font=('Inter', 11),
        )
        self.delete_all_button.pack(side='left', padx=5)

        # Bouton pour traiter les fichiers
        self.process_button = tk.Button(
            main_frame,
            text="Traiter les fichiers",
            command=self.process_files,
            font=('Inter', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            state='disabled',
            height=2
        )
        self.process_button.grid(row=5, column=0, pady=20)

        # Create frame for status and error details
        status_frame = tk.Frame(main_frame)
        status_frame.grid(row=6, column=0, pady=10)
        
        # Status label
        self.status_label = tk.Label(
            status_frame, 
            text="",
            font=('Inter', 9, 'italic')
        )
        self.status_label.pack(side='left', padx=5)
        
        # Error details button (hidden by default)
        self.error_details_button = tk.Button(
            status_frame,
            text="Voir détails",
            command=self.show_error_details,
            font=('Inter', 9),
            fg='red',
            state='disabled'
        )
        self.error_details_button.pack(side='left', padx=5)
        
        self.detect_encoding_button = tk.Button(
            status_frame,
            text="Détecter l'encodage",
            command=lambda: self.open_encoding_window(),
            font=('Inter', 9),
            state='disabled'  # Initially disabled until files are selected
        )
        self.detect_encoding_button.pack(side='left', padx=5)
        self.info_label = tk.Label(status_frame, text="Maintenez Ctrl (ou Shift) pour sélectionner plusieurs fichiers", font=('Inter', 9, 'italic'))
        self.info_label.pack(side='left', padx=5)

    def find_encoding(self,file_path):
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            detected = detect(raw_data)
            return detected['encoding']

    def open_encoding_window(self):
        
        selected_items = self.files_tree.selection()
        if not selected_items:
            messagebox.showinfo("Encodage", "Veuillez sélectionner un ou plusieurs fichiers.")
            return

        window = tk.Toplevel(self)
        window.title("Détection des encodages")
        window.geometry("800x400")
        window.grab_set()
        self.center_main_window()

        title = tk.Label(window, text="Encodage des fichiers sélectionnés", font=('Inter', 14, 'bold'))
        title.pack(pady=10)

        frame = tk.Frame(window)
        frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Treeview to display file path and encoding
        encoding_tree = ttk.Treeview(frame, columns=('Fichier', 'Encodage'), show='headings')
        encoding_tree.heading('Fichier', text='Fichier')
        encoding_tree.heading('Encodage', text='Encodage')
        encoding_tree.column('Fichier', width=600)
        encoding_tree.column('Encodage', width=150)
        encoding_tree.pack(side='left', fill='both', expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=encoding_tree.yview)
        encoding_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

        detected_encodings = []  # To store detected encodings
        # Detect encodings and populate the tree
        for iid in selected_items:
            file_path = iid  # Since iid is the full path
            try:
                encoding = self.find_encoding(file_path)
                detected_encodings.append(encoding)
            except Exception as e:
                encoding = f"Erreur: {e}"
            encoding_tree.insert('', 'end', values=(os.path.basename(file_path), encoding))

        # Close button
        close_btn = tk.Button(window, text="Fermer", command=window.destroy, font=('Inter', 10))
        close_btn.pack(pady=10)
        # Update the status label with detected encodings
        if len(detected_encodings) > 0:
            encodings_str = ', '.join(set(detected_encodings))
            self.detect_encoding_button.config(state='disabled')        
            self.info_label.config(text="Rajoutez les encodages ne figurant pas dans les paramètres: " + encodings_str)

    def show_info_popup(self,event=None):
        text = "\nLorsque vous sauvegardez les paramètres, un fichier nommé 'configuration.json' sera généré dans le dossier où se trouve l'application.\n\n" \
               "Il servira de fichier de configuration par défaut.\nVous pouvez également exporter ou importer les \nconfigurations afin de les réutiliser pour d'autres FECS.\n\n" \
               "En cochant la case 'Définir comme configuration par défaut' avant de sauvegarder, vous pouvez enregistrer la configuration actuelle comme la nouvelle configuration par défaut.\n"
        
        if self.info_popup is not None:
            return

        self.info_popup = tk.Toplevel(self)
        self.info_popup.overrideredirect(True)
        self.info_popup.configure(bg="#ffffff", padx=10, pady=5)
        self.info_popup.wm_attributes("-topmost", 1)

        # Position the popup near the icon
        x = self.info_icon.winfo_rootx() + 10
        y = self.info_icon.winfo_rooty() + 20
        self.info_popup.geometry(f"380x350+{x}+{y}")


        label= ttk.Label(
            self.info_popup,
            text=text,
            background="#ffffff",
            justify="left",
            wraplength=350,
            font=("Inter", 10, "italic") # Another font family and style
        )
        label.pack(pady=5)

    def close_info_popup(self, event=None):
        if self.info_popup is not None:
            self.info_popup.destroy()
            self.info_popup = None

    def center_main_window(self):
        """Center the window on the screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"+{x}+{y}")

    def center_config_window(self, window):
        """Center the given window relative to the main application."""
        self.update_idletasks()  # Ensure size is updated

        width = window.winfo_width()
        height = window.winfo_height()

        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()

        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2

        window.geometry(f"{width}x{height}+{x}+{y}")

    def open_settings_window(self, has_imported_config=False):
        # Always destroy any existing settings window to ensure a clean slate
        # and that the new window receives the latest configuration.
        if self._settings_window and self._settings_window.winfo_exists():
            self._settings_window.destroy()
        
        self._settings_window = SettingsWindow(
            self, 
            self.parametres, 
            self.update_app_config, 
            self.open_settings_window, 
            has_imported_config=has_imported_config
        )
        self._settings_window.transient(self)
        self._settings_window.grab_set()
        self.center_config_window(self._settings_window)
        
        self.wait_window(self._settings_window)
        # Crucial fix: After the window closes, clear the reference to the destroyed window
        self._settings_window = None

    def update_app_config(self, new_config):
        """Callback to update the main app's configuration after settings are saved."""
        self.parametres = new_config

    # Fonction pour ajouter plusieurs fichiers FEC
    def add_fec_files(self):
        """
        Ajoute plusieurs fichiers FEC à la liste des fichiers sélectionnés.
        """
        try:
            file_paths = filedialog.askopenfilenames(
                title="Sélectionner les fichiers FEC",
                filetypes=[("Fichiers texte", "*.txt")]
            )
            
            if file_paths:
                default_year = ''
                
                for file_path in file_paths:
                    # Extraire le nom du fichier sans le chemin
                    file_name = os.path.basename(file_path)
                    # default_year = file_name.split("FEC")[-1][:4] if file_name.split("FEC")[-1][:3] in [f"20{i}" for i in range(1,4)] else ''
                    matches = re.findall(r'FEC.*?(20\d{2})', file_name)
                    default_year = matches[-1] if matches else ''
                    
                    # Ajouter le fichier avec une année par défaut (année actuelle)
                    item_id = self.files_tree.insert('', 'end',iid=file_path, values=(default_year, file_name))
                    
                    # Créer un widget Entry pour l'année et l'associer à cette ligne
                    year_entry = ttk.Entry(self.files_tree, width=10)
                    year_entry.insert(0, str(default_year))
                    
                    # Ajouter une fonction de callback pour mettre à jour le TreeView et la liste lorsque l'année change
                    def update_year(event, item_id=item_id):
                        try:
                            year = self.entry_widgets[item_id].get()
                            self.files_tree.set(item_id, 'Année', year)
                            
                            # Mettre à jour l'année dans la liste des fichiers
                            for i, (old_year, file_path, entry_item_id) in enumerate(self.fec_files):
                                if entry_item_id == item_id:
                                    self.fec_files[i] = (year, file_path, item_id)
                                    break
                        except ValueError:
                            # Réinitialiser à la valeur précédente en cas d'entrée invalide
                            for year, _, entry_item_id in self.fec_files:
                                if entry_item_id == item_id:
                                    self.entry_widgets[item_id].delete(0, tk.END)
                                    self.entry_widgets[item_id].insert(0, str(year))
                                    break
                    
                    year_entry.bind('<FocusOut>', update_year)
                    year_entry.bind('<Return>', update_year)
                    
                    # Positionner l'Entry dans la colonne "Année"
                    self.files_tree.update()
                    bbox = self.files_tree.bbox(item_id, 'Année')
                    if bbox:
                        year_entry.place(x=bbox[0], y=bbox[1], width=bbox[2])
                    
                    # Stocker l'Entry dans le dictionnaire et les informations du fichier dans la liste
                    self.entry_widgets[item_id] = year_entry
                    self.fec_files.append((default_year, file_path, item_id))
                
                self.process_button.config(state='normal')
                self.detect_encoding_button.config(state='normal')
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite: {str(e)}")

    # Fonction pour supprimer un fichier FEC
    def delete_selected(self):
        """
        Supprime le fichier sélectionné de la liste.
        """
        selected_items = self.files_tree.selection()
        
        if not selected_items:
            messagebox.showinfo("Information", "Veuillez sélectionner un fichier à supprimer.")
            return
        
        for item_id in selected_items:
            # Trouver et supprimer l'entrée correspondante dans fec_files
            for i, (_, _, entry_item_id) in enumerate(self.fec_files):
                if entry_item_id == item_id:
                    # Supprimer l'Entry widget associé
                    if item_id in self.entry_widgets:
                        self.entry_widgets[item_id].destroy()
                        del self.entry_widgets[item_id]
                    
                    # Supprimer de la liste des fichiers
                    self.fec_files.pop(i)
                    break
            
            # Supprimer de l'arbre
            self.files_tree.delete(item_id)
        
        # Désactiver le bouton de traitement si plus de fichiers
        if not self.fec_files:
            self.process_button.config(state='disabled')

    # Fonction pour supprimer tous les fichiers
    def delete_all_files(self):
        """
        Supprime tous les fichiers de la liste.
        """
        # Supprimer tous les Entry widgets
        for entry in self.entry_widgets.values():
            entry.destroy()
        
        # Vider les dictionnaires et listes
        self.entry_widgets.clear()
        self.fec_files.clear()
        
        # Vider l'arbre
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
        
        # Désactiver le bouton de traitement
        self.process_button.config(state='disabled')

    # Fenêtre de configuration de l'entreprise
    def create_company_config_window(self):
        """
        Creates a configuration window to select company type and which sheets to include.
        This window appears after clicking 'Traiter les fichiers' but before processing.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configuration")
        config_window.geometry("600x600")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Store the result
        config_window.result = {
            'company_type': None,
            'sheets': {},
            'exclude_salaries': False
        }
        
        # Configure grid
        config_window.columnconfigure(0, weight=1)
        
        # Main frame with padding
        main_frame = tk.Frame(config_window, padx=20, pady=20)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = tk.Label(
            main_frame,
            text="Configuration du traitement",
            font=('Inter', 16, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky='w')
        
        # Company type frame
        company_type_frame = tk.LabelFrame(main_frame, text="Type d'entreprise", padx=15, pady=15)
        company_type_frame.grid(row=1, column=0, sticky='ew', pady=(0, 20))
        
        # Company type variables
        is_holding_var = tk.BooleanVar(value=False)
        is_other_var = tk.BooleanVar(value=True)
        
        # Company type checkboxes
        holding_cb = ttk.Checkbutton(
            company_type_frame,
            text="Holding",
            variable=is_holding_var,
            command=lambda: update_sheet_options('holding', is_holding_var.get())
        )
        holding_cb.grid(row=0, column=0, padx=10, sticky='w')
        
        other_cb = ttk.Checkbutton(
            company_type_frame,
            text="Autres",
            variable=is_other_var,
            command=lambda: update_sheet_options('other', is_other_var.get())
        )
        other_cb.grid(row=0, column=1, padx=10, sticky='w')
        
        # Sheets frame
        sheets_frame = tk.LabelFrame(main_frame, text="Feuilles à inclure", padx=15, pady=15)
        sheets_frame.grid(row=2, column=0, sticky='ew', pady=(0, 20))
        
        # Sheet options
        sheet_options = {
            'CA': tk.BooleanVar(value=True),
            'Trésorerie': tk.BooleanVar(value=True),
            'Achats et Sous-traitance': tk.BooleanVar(value=False),
            'Crédit Moyen FR': tk.BooleanVar(value=True),
            'TOP FR': tk.BooleanVar(value=True),
            'Frais personnels': tk.BooleanVar(value=True),
            'Crédit Moyen Client': tk.BooleanVar(value=True),
            'TOP Client': tk.BooleanVar(value=True),
            'Investissement' : tk.BooleanVar(value=False)
        }
        
        # Create sheet checkboxes (2 columns)
        sheet_checkboxes = {}
        row_idx = 0
        col_idx = 0
        
        for sheet_name, var in sheet_options.items():
            cb = ttk.Checkbutton(
                sheets_frame,
                text=sheet_name,
                variable=var
            )
            cb.grid(row=row_idx, column=col_idx, padx=10, pady=5, sticky='w')
            sheet_checkboxes[sheet_name] = cb
            
            col_idx += 1
            if col_idx > 1:
                col_idx = 0
                row_idx += 1
        
        # Salary exclusion frame
        salary_frame = tk.LabelFrame(main_frame, text="Options supplémentaires", padx=15, pady=15)
        salary_frame.grid(row=3, column=0, sticky='ew', pady=(0, 20))
        
        # Salary exclusion option
        exclude_salary_var = tk.BooleanVar(value=False)
        exclude_salary_cb = ttk.Checkbutton(
            salary_frame,
            text="Exclure Salaires et charges sociales",
            variable=exclude_salary_var
        )
        exclude_salary_cb.grid(row=0, column=0, padx=10, sticky='w')

        # Reporting generation option
        reporting_var = tk.BooleanVar(value=False)
        reporting_var_cb = ttk.Checkbutton(
            salary_frame,
            text="Générer un reporting",
            variable=reporting_var
        )
        reporting_var_cb.grid(row=0, column=1, padx=10, sticky='w')

        # Fournisseur exclusion option
        exclude_cpte_fournisseur = tk.BooleanVar(value=False)
        exclude_cpte_fournisseur_cb = ttk.Checkbutton(
            salary_frame,
            text="Exclure Fournisseurs pour CMF",
            variable=exclude_cpte_fournisseur
        )
        exclude_cpte_fournisseur_cb.grid(row=1, column=0, padx=10, sticky='w')

        # Confirm button
        confirm_button = ttk.Button(
            main_frame,
            text="Confirmer choix",
            command=lambda: confirm_selection()
        )
        confirm_button.grid(row=4, column=0, pady=20)
        
        # Function to update sheet options based on company type
        def update_sheet_options(company_type, is_checked):
            if company_type == 'holding' and is_checked:
                is_other_var.set(False)
                
                # Disable specific sheets for holdings
                sheet_options['Frais personnels'].set(False)
                sheet_options['Crédit Moyen Client'].set(False)
                sheet_options['TOP Client'].set(False)
                
                # Disable the checkboxes
                sheet_checkboxes['Frais personnels'].configure(state='disabled')
                sheet_checkboxes['Crédit Moyen Client'].configure(state='disabled')
                sheet_checkboxes['TOP Client'].configure(state='disabled')
                
            elif company_type == 'other' and is_checked:
                is_holding_var.set(False)
                
                # Enable all sheets for other company types
                for sheet_name in sheet_options:
                    if sheet_name not in ['Achats et Sous-traitance','Investissement']:
                        sheet_options[sheet_name].set(True)
                    else:
                        sheet_options[sheet_name].set(False)
                    sheet_checkboxes[sheet_name].configure(state='normal')
        
        # Function to confirm selection
        def confirm_selection():
            # Determine company type
            if is_holding_var.get():
                company_type = 'holding'
            elif is_other_var.get():
                company_type = 'other'
            else:
                messagebox.showwarning(
                    "Avertissement",
                    "Veuillez sélectionner un type d'entreprise"
                )
                return
            
            # Get selected sheets
            selected_sheets = {sheet: var.get() for sheet, var in sheet_options.items()}
            
            # Store results
            config_window.result = {
                'company_type': company_type,
                'sheets': selected_sheets,
                'exclude_salaries': exclude_salary_var.get(),
                'reporting': reporting_var.get(),
                'exclude_cpte_fournisseur':  exclude_cpte_fournisseur.get()
            }
            
            # Close window
            config_window.destroy()
        
        # Center window on screen
        config_window.update_idletasks()
        width = config_window.winfo_width()
        height = config_window.winfo_height()
        x = (config_window.winfo_screenwidth() // 2) - (width // 2)
        y = (config_window.winfo_screenheight() // 2) - (height // 2)
        config_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Wait for window
        config_window.wait_window()
        return config_window.result

    # Fonction pour ouvrir un fichier Excel
    def open_excel_file(self, file_path):
        """
        Ouvre le fichier Excel spécifié.

        Args:
            file_path (str): Chemin du fichier Excel à ouvrir.
        """
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # linux
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            messagebox.showwarning(
                "Attention",
                f"Le fichier a été créé mais n'a pas pu être ouvert automatiquement.\nVous pouvez le trouver ici : {file_path}"
            )

    # Fenêtre de configuration des frais personnels
    def create_frais_personnels_config_window(self, df_by_year):
        """
        Creates a configuration window with checkboxes for all years.
        Automatically checks boxes across years for matching accounts.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configuration")
        config_window.geometry("1000x700")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Create notebook
        notebook = ttk.Notebook(config_window)
        notebook.pack(expand=True, fill='both', padx=20, pady=20)
        
        # Store all selections and checkbuttons
        selections = {}
        all_checkbuttons = {
            'salaires': {},  # Will store {compte_lib: {year: checkbutton}}
            'charges': {}
        }
        
        # Function to handle checkbox changes
        def on_checkbox_change(compte_lib, compte_type, year, var):
            """Propagate selection to other years if they have the same account"""
            if var.get():  # Only propagate when checking, not unchecking
                for other_year, year_data in selections.items():
                    if other_year != year:  # Don't process the current year
                        # Check if this account exists in the other year
                        if compte_lib in year_data[compte_type]:
                            year_data[compte_type][compte_lib].set(True)

        # Create tabs for each year
        for year, df in sorted(df_by_year.items()):
            year_frame = ttk.Frame(notebook)
            notebook.add(year_frame, text=f"Année {year}")
            
            # Configure grid
            year_frame.columnconfigure(0, weight=1)
            year_frame.columnconfigure(1, weight=1)
            
            # Create frames
            salaires_frame = ttk.LabelFrame(year_frame, text="Comptes à exclure des salaires")
            salaires_frame.grid(row=0, column=0, padx=15, pady=15, sticky='nsew')
            
            charges_frame = ttk.LabelFrame(year_frame, text="Comptes à exclure des charges sociales")
            charges_frame.grid(row=0, column=1, padx=15, pady=15, sticky='nsew')
            
            # Create canvas and scrollbars
            salaires_canvas = tk.Canvas(salaires_frame)
            salaires_scrollbar = ttk.Scrollbar(salaires_frame, orient="vertical", command=salaires_canvas.yview)
            salaires_subframe = ttk.Frame(salaires_canvas)
            
            charges_canvas = tk.Canvas(charges_frame)
            charges_scrollbar = ttk.Scrollbar(charges_frame, orient="vertical", command=charges_canvas.yview)
            charges_subframe = ttk.Frame(charges_canvas)
            
            # Configure scrolling
            salaires_canvas.configure(yscrollcommand=salaires_scrollbar.set)
            charges_canvas.configure(yscrollcommand=charges_scrollbar.set)
            
            # Pack scrollbars and canvas
            salaires_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            salaires_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            charges_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            charges_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create windows in canvas
            salaires_canvas.create_window((0, 0), window=salaires_subframe, anchor='nw')
            charges_canvas.create_window((0, 0), window=charges_subframe, anchor='nw')
            
            # Get accounts
            salaires_comptes = df[
                df['CompteNum'].astype(str).str.startswith('641')
            ][['CompteNum', 'CompteLib']].drop_duplicates()
            
            charges_comptes = df[
                (df['CompteNum'].astype(str).str.startswith('64')) & 
                (~df['CompteNum'].astype(str).str.startswith('641'))
            ][['CompteNum', 'CompteLib']].drop_duplicates()
            
            salaires_vars = {}
            charges_vars = {}
            
            # Create checkbuttons for salaires
            for _, row in salaires_comptes.iterrows():
                var = tk.BooleanVar()
                compte_lib = row['CompteLib']
                
                # Initialize the dictionary for this account if it doesn't exist
                if compte_lib not in all_checkbuttons['salaires']:
                    all_checkbuttons['salaires'][compte_lib] = {}
                
                cb = ttk.Checkbutton(
                    salaires_subframe, 
                    text=f"{row['CompteNum']} - {compte_lib}", 
                    variable=var
                )
                cb.pack(anchor='w', padx=10, pady=5)
                
                # Store the variable and set up the callback
                salaires_vars[compte_lib] = var
                all_checkbuttons['salaires'][compte_lib][year] = var
                var.trace_add('write', lambda *args, 
                             cl=compte_lib, y=year, v=var: on_checkbox_change(cl, 'salaires', y, v))
                
            # Create checkbuttons for charges
            for _, row in charges_comptes.iterrows():
                var = tk.BooleanVar()
                compte_lib = row['CompteLib']
                
                # Initialize the dictionary for this account if it doesn't exist
                if compte_lib not in all_checkbuttons['charges']:
                    all_checkbuttons['charges'][compte_lib] = {}
                
                cb = ttk.Checkbutton(
                    charges_subframe, 
                    text=f"{row['CompteNum']} - {compte_lib}", 
                    variable=var
                )
                cb.pack(anchor='w', padx=10, pady=5)
                
                # Store the variable and set up the callback
                charges_vars[compte_lib] = var
                all_checkbuttons['charges'][compte_lib][year] = var
                var.trace_add('write', lambda *args, 
                             cl=compte_lib, y=year, v=var: on_checkbox_change(cl, 'charges', y, v))
            
            # Enable mousewheel scrolling
            def _on_mousewheel(event, canvas):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                
            salaires_canvas.bind_all("<MouseWheel>", lambda e: _on_mousewheel(e, salaires_canvas))
            charges_canvas.bind_all("<MouseWheel>", lambda e: _on_mousewheel(e, charges_canvas))
            
            # Update scroll region
            salaires_subframe.update_idletasks()
            charges_subframe.update_idletasks()
            salaires_canvas.configure(scrollregion=salaires_canvas.bbox("all"))
            charges_canvas.configure(scrollregion=charges_canvas.bbox("all"))
            
            # Store variables
            selections[year] = {
                'salaires': salaires_vars,
                'charges': charges_vars
            }
        
        # Add confirmation button
        confirm_button = ttk.Button(
            config_window, 
            text="Confirmer",
            command=lambda: confirm_selection()
        )
        confirm_button.pack(pady=20)
        
        def confirm_selection():
            config = {}
            for year, year_data in selections.items():
                config[year] = {
                    'excluded_salaires': [
                        compte for compte, var in year_data['salaires'].items()
                        if var.get()
                    ],
                    'excluded_charges': [
                        compte for compte, var in year_data['charges'].items()
                        if var.get()
                    ]
                }
            config_window.result = config
            config_window.destroy()
        
        # Center window on screen
        config_window.update_idletasks()
        width = config_window.winfo_width()
        height = config_window.winfo_height()
        x = (config_window.winfo_screenwidth() // 2) - (width // 2)
        y = (config_window.winfo_screenheight() // 2) - (height // 2)
        config_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Wait for window
        config_window.wait_window()
        return getattr(config_window, 'result', {})

    # Fenêtre de configuration pour les top dix clients et fournisseurs
    def create_top_dix_config_window(self, df_by_year, fournisseurs_codes=[], clients_codes=[], journal_code_achats = [], journal_code_ventes = []):
        """
        Creates a configuration window with checkboxes for suppliers (fournisseurs) and clients by year.
        Includes a search bar and propagates selections across years for matching accounts.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configuration Top 10 Fournisseurs et Clients (Exclusion Fournisseurs/Clients divers)")
        config_window.geometry("1000x700")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Create notebook
        notebook = ttk.Notebook(config_window)
        notebook.pack(expand=True, fill='both', padx=20, pady=20)
        
        # Store all selections and checkbuttons
        selections = {}
        all_checkbuttons = {
            'fournisseurs': {},  # Will store {compte_lib: {year: checkbutton}}
            'clients': {}
        }
        
        # Function to handle checkbox changes
        def on_checkbox_change(compte_lib, compte_type, year, var):
            """Propagate selection to other years if they have the same account"""
            if var.get():  # Only propagate when checking, not unchecking
                for other_year, year_data in selections.items():
                    if other_year != year:  # Don't process the current year
                        # Check if this account exists in the other year
                        if compte_lib in year_data[compte_type]:
                            year_data[compte_type][compte_lib].set(True)
        
        # Function to filter checkboxes based on search
        def filter_checkboxes(search_term, frame, checkbutton_frames):
            search_term = search_term.lower()
            
            for compte_lib, checkbutton_frame in checkbutton_frames.items():
                if search_term in compte_lib.lower():
                    checkbutton_frame.pack(anchor='w', padx=10, pady=5)
                else:
                    checkbutton_frame.pack_forget()
            
            # Update scroll region
            frame.update_idletasks()
            frame.master.configure(scrollregion=frame.master.bbox("all"))
        
        # Create tabs for each year
        for year, df in sorted(df_by_year.items()):
            year_frame = ttk.Frame(notebook)
            notebook.add(year_frame, text=f"Année {year}")
            
            # Configure grid
            year_frame.columnconfigure(0, weight=1)
            year_frame.columnconfigure(1, weight=1)
            
            # Create frames
            fournisseurs_frame = ttk.LabelFrame(year_frame, text="Comptes Fournisseurs à exclure")
            fournisseurs_frame.grid(row=1, column=0, padx=15, pady=5, sticky='nsew')
            
            clients_frame = ttk.LabelFrame(year_frame, text="Comptes Clients à exclure")
            clients_frame.grid(row=1, column=1, padx=15, pady=5, sticky='nsew')
            
            # Create search frames
            search_frame_fournisseurs = ttk.Frame(year_frame)
            search_frame_fournisseurs.grid(row=0, column=0, padx=15, pady=(15,0), sticky='ew')
            
            search_frame_clients = ttk.Frame(year_frame)
            search_frame_clients.grid(row=0, column=1, padx=15, pady=(15,0), sticky='ew')
            
            # Create search entries
            ttk.Label(search_frame_fournisseurs, text="Rechercher:").pack(side=tk.LEFT, padx=(0,5))
            search_entry_fournisseurs = ttk.Entry(search_frame_fournisseurs)
            search_entry_fournisseurs.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            ttk.Label(search_frame_clients, text="Rechercher:").pack(side=tk.LEFT, padx=(0,5))
            search_entry_clients = ttk.Entry(search_frame_clients)
            search_entry_clients.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # Create canvas and scrollbars
            fournisseurs_canvas = tk.Canvas(fournisseurs_frame)
            fournisseurs_scrollbar = ttk.Scrollbar(fournisseurs_frame, orient="vertical", command=fournisseurs_canvas.yview)
            fournisseurs_subframe = ttk.Frame(fournisseurs_canvas)
            
            clients_canvas = tk.Canvas(clients_frame)
            clients_scrollbar = ttk.Scrollbar(clients_frame, orient="vertical", command=clients_canvas.yview)
            clients_subframe = ttk.Frame(clients_canvas)
            
            # Configure scrolling
            fournisseurs_canvas.configure(yscrollcommand=fournisseurs_scrollbar.set)
            clients_canvas.configure(yscrollcommand=clients_scrollbar.set)
            
            # Pack scrollbars and canvas
            fournisseurs_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            fournisseurs_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            clients_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            clients_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create windows in canvas
            fournisseurs_canvas.create_window((0, 0), window=fournisseurs_subframe, anchor='nw')
            clients_canvas.create_window((0, 0), window=clients_subframe, anchor='nw')
            
            compte_num_col = 'CompAuxNum'
            compte_lib_col = 'CompAuxLib'
            
            if len(list(df["CompAuxLib"].value_counts())) == 0 or len(list(df["CompAuxNum"].value_counts())) ==0:
                compte_num_col = 'CompteNum'
                compte_lib_col = 'CompteLib'

            liste_journal_code = journal_code_achats + journal_code_ventes

            journal_code_filtered = df["JournalCode"].astype(str).str.strip().isin(liste_journal_code)


            # Get accounts
            fournisseurs_comptes = df[
                (journal_code_filtered) & 
                (df['CompteNum'].astype(str).str[:3].isin(fournisseurs_codes))
            ][[compte_num_col, compte_lib_col]].drop_duplicates()
            
            clients_comptes = df[
                (journal_code_filtered) & 
                (df['CompteNum'].astype(str).str[:3].isin(clients_codes))
            ][[compte_num_col, compte_lib_col]].drop_duplicates()
            
            fournisseurs_vars = {}
            clients_vars = {}
            fournisseurs_frames = {}
            clients_frames = {}
            
            # Create checkbuttons for fournisseurs
            for _, row in fournisseurs_comptes.iterrows():
                var = tk.BooleanVar()
                compte_lib = row[compte_lib_col]
                compte_num = row[compte_num_col]
                display_text = f"{compte_num} - {compte_lib}"
                
                # Initialize the dictionary for this account if it doesn't exist
                if compte_lib not in all_checkbuttons['fournisseurs']:
                    all_checkbuttons['fournisseurs'][compte_lib] = {}
                
                # Create a frame for this checkbutton (for filtering)
                cb_frame = ttk.Frame(fournisseurs_subframe)
                cb_frame.pack(anchor='w', padx=10, pady=5)
                fournisseurs_frames[compte_lib] = cb_frame
                
                cb = ttk.Checkbutton(
                    cb_frame, 
                    text=display_text, 
                    variable=var
                )
                cb.pack(anchor='w')
                
                # Store the variable and set up the callback
                fournisseurs_vars[compte_lib] = var
                all_checkbuttons['fournisseurs'][compte_lib][year] = var
                var.trace_add('write', lambda *args, 
                            cl=compte_lib, y=year, v=var: on_checkbox_change(cl, 'fournisseurs', y, v))
                
            # Create checkbuttons for clients
            for _, row in clients_comptes.iterrows():
                var = tk.BooleanVar()
                compte_lib = row[compte_lib_col]
                compte_num = row[compte_num_col]
                display_text = f"{compte_num} - {compte_lib}"
                
                # Initialize the dictionary for this account if it doesn't exist
                if compte_lib not in all_checkbuttons['clients']:
                    all_checkbuttons['clients'][compte_lib] = {}
                
                # Create a frame for this checkbutton (for filtering)
                cb_frame = ttk.Frame(clients_subframe)
                cb_frame.pack(anchor='w', padx=10, pady=5)
                clients_frames[compte_lib] = cb_frame
                
                cb = ttk.Checkbutton(
                    cb_frame, 
                    text=display_text, 
                    variable=var
                )
                cb.pack(anchor='w')
                
                # Store the variable and set up the callback
                clients_vars[compte_lib] = var
                all_checkbuttons['clients'][compte_lib][year] = var
                var.trace_add('write', lambda *args, 
                            cl=compte_lib, y=year, v=var: on_checkbox_change(cl, 'clients', y, v))
            
            # Configure search functionality
            search_var_fournisseurs = tk.StringVar()
            search_entry_fournisseurs.config(textvariable=search_var_fournisseurs)

            def create_filter_callback(var, frame, frames_dict):
                return lambda *args: filter_checkboxes(var.get(), frame, frames_dict)

            search_var_fournisseurs.trace_add('write', create_filter_callback(search_var_fournisseurs, fournisseurs_subframe, fournisseurs_frames))

                        
            search_var_clients = tk.StringVar()
            search_entry_clients.config(textvariable=search_var_clients)
            search_var_clients.trace_add('write', create_filter_callback(search_var_clients, clients_subframe, clients_frames))

            
            # Enable mousewheel scrolling
            def _on_mousewheel(event, canvas):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            def _bind_mousewheel(event, canvas):
                canvas.bind_all("<MouseWheel>", lambda e: _on_mousewheel(e, canvas))
                
            def _unbind_mousewheel(event):
                event.widget.unbind_all("<MouseWheel>")
                
            # Bind mousewheel only when mouse is over the specific canvas
            fournisseurs_canvas.bind("<Enter>", lambda e: _bind_mousewheel(e, fournisseurs_canvas))
            fournisseurs_canvas.bind("<Leave>", _unbind_mousewheel)
            
            clients_canvas.bind("<Enter>", lambda e: _bind_mousewheel(e, clients_canvas))
            clients_canvas.bind("<Leave>", _unbind_mousewheel)
            
            # Update scroll region
            fournisseurs_subframe.update_idletasks()
            clients_subframe.update_idletasks()
            fournisseurs_canvas.configure(scrollregion=fournisseurs_canvas.bbox("all"))
            clients_canvas.configure(scrollregion=clients_canvas.bbox("all"))
            
            # Store variables
            selections[year] = {
                'fournisseurs': fournisseurs_vars,
                'clients': clients_vars
            }
        
        # Add confirmation button
        confirm_button = ttk.Button(
            config_window, 
            text="Confirmer",
            command=lambda: confirm_selection()
        )
        confirm_button.pack(pady=20)
        
        def confirm_selection():
            config = {}
            for year, year_data in selections.items():
                config[year] = {
                    'excluded_fournisseurs': [
                        compte for compte, var in year_data['fournisseurs'].items()
                        if var.get()
                    ],
                    'excluded_clients': [
                        compte for compte, var in year_data['clients'].items()
                        if var.get()
                    ]
                }
            config_window.result = config
            config_window.destroy()
        
        # Center window on screen
        config_window.update_idletasks()
        width = config_window.winfo_width()
        height = config_window.winfo_height()
        x = (config_window.winfo_screenwidth() // 2) - (width // 2)
        y = (config_window.winfo_screenheight() // 2) - (height // 2)
        config_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Wait for window
        config_window.wait_window()
        return getattr(config_window, 'result', {})

    # Fenêtre de configuration pour l'exclusion des fournisseurs dans le CMF
    def exclusion_credit_moyen(self, df_by_year, fournisseurs_codes=[], journal_code_achats = []):
        """
        Creates a configuration window with checkboxes for suppliers (fournisseurs) by year.
        Includes a search bar and propagates selections across years for matching accounts.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configuration Crédit moyen Fournisseurs (Exclusion)")
        config_window.geometry("1000x700")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Create notebook
        notebook = ttk.Notebook(config_window)
        notebook.pack(expand=True, fill='both', padx=20, pady=20)
        
        # Store all selections and checkbuttons
        selections = {}
        fournisseurs_checked = {}
        
        # Function to handle checkbox changes
        def on_checkbox_change(compte_lib, year, var):
            """Propagate selection to other years if they have the same account"""
            if var.get():  # Only propagate when checking, not unchecking
                for other_year, year_data in selections.items():
                    if other_year != year:  # Don't process the current year
                        # Check if this account exists in the other year
                        if compte_lib in year_data:
                            year_data[compte_lib].set(True)
        
        # Function to filter checkboxes based on search
        def filter_checkboxes(search_term, frame, checkbutton_frames):
            search_term = search_term.lower()
            
            for compte_lib, checkbutton_frame in checkbutton_frames.items():
                if search_term in compte_lib.lower():
                    checkbutton_frame.pack(anchor='w', padx=10, pady=5)
                else:
                    checkbutton_frame.pack_forget()
            
            # Update scroll region
            frame.update_idletasks()
            frame.master.configure(scrollregion=frame.master.bbox("all"))
        
        # Create tabs for each year
        for year, df in sorted(df_by_year.items()):
            year_frame = ttk.Frame(notebook)
            notebook.add(year_frame, text=f"Année {year}")
            
            # Configure grid
            year_frame.columnconfigure(0, weight=1)
            
            # Create frames
            fournisseurs_frame = ttk.LabelFrame(year_frame, text="Comptes Fournisseurs à exclure")
            fournisseurs_frame.grid(row=1, column=0, padx=15, pady=5, sticky='nsew')
                        
            # Create search frames
            search_frame_fournisseurs = ttk.Frame(year_frame)
            search_frame_fournisseurs.grid(row=0, column=0, padx=15, pady=(15,0), sticky='ew')
            
            
            # Create search entries
            ttk.Label(search_frame_fournisseurs, text="Rechercher:").pack(side=tk.LEFT, padx=(0,5))
            search_entry_fournisseurs = ttk.Entry(search_frame_fournisseurs)
            search_entry_fournisseurs.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            
            # Create canvas and scrollbars
            fournisseurs_canvas = tk.Canvas(fournisseurs_frame)
            fournisseurs_scrollbar = ttk.Scrollbar(fournisseurs_frame, orient="vertical", command=fournisseurs_canvas.yview)
            fournisseurs_subframe = ttk.Frame(fournisseurs_canvas)
            
            
            # Configure scrolling
            fournisseurs_canvas.configure(yscrollcommand=fournisseurs_scrollbar.set)
            
            # Pack scrollbars and canvas
            fournisseurs_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            fournisseurs_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
            
            # Create windows in canvas
            fournisseurs_canvas.create_window((0, 0), window=fournisseurs_subframe, anchor='nw')
            
            compte_num_col = 'CompAuxNum'
            compte_lib_col = 'CompAuxLib'
            
            if len(list(df["CompAuxLib"].value_counts())) == 0 or len(list(df["CompAuxNum"].value_counts())) ==0:
                compte_num_col = 'CompteNum'
                compte_lib_col = 'CompteLib'


            journal_code_filtered = df["JournalCode"].astype(str).str.strip().isin(journal_code_achats)

            # Get accounts
            fournisseurs_comptes = df[
                (journal_code_filtered) & 
                (df['CompteNum'].astype(str).str[:3].isin(fournisseurs_codes))
            ][[compte_num_col, compte_lib_col]].drop_duplicates()

            fournisseurs_vars = {}
            fournisseurs_frames = {}
            
            # Create checkbuttons for fournisseurs
            for _, row in fournisseurs_comptes.iterrows():
                var = tk.BooleanVar()
                compte_lib = row[compte_lib_col]
                compte_num = row[compte_num_col]
                display_text = f"{compte_num} - {compte_lib}"
                
                # Initialize the dictionary for this account if it doesn't exist
                if compte_lib not in fournisseurs_checked:
                    fournisseurs_checked[compte_lib] = {}
                
                # Create a frame for this checkbutton (for filtering)
                cb_frame = ttk.Frame(fournisseurs_subframe)
                cb_frame.pack(anchor='w', padx=10, pady=5)
                fournisseurs_frames[compte_lib] = cb_frame
                
                cb = ttk.Checkbutton(
                    cb_frame, 
                    text=display_text, 
                    variable=var
                )
                cb.pack(anchor='w')
                
                # Store the variable and set up the callback
                fournisseurs_vars[compte_lib] = var
                fournisseurs_checked[compte_lib][year] = var
                var.trace_add('write', lambda *args, 
                            cl=compte_lib, y=year, v=var: on_checkbox_change(cl, y, v))
                
            # Create checkbuttons for clients

            
            # Configure search functionality
            search_var_fournisseurs = tk.StringVar()
            search_entry_fournisseurs.config(textvariable=search_var_fournisseurs)

            def create_filter_callback(var, frame, frames_dict):
                return lambda *args: filter_checkboxes(var.get(), frame, frames_dict)

            search_var_fournisseurs.trace_add('write', create_filter_callback(search_var_fournisseurs, fournisseurs_subframe, fournisseurs_frames))
            
            # Enable mousewheel scrolling
            def _on_mousewheel(event, canvas):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            def _bind_mousewheel(event, canvas):
                canvas.bind_all("<MouseWheel>", lambda e: _on_mousewheel(e, canvas))
                
            def _unbind_mousewheel(event):
                event.widget.unbind_all("<MouseWheel>")
                
            # Bind mousewheel only when mouse is over the specific canvas
            fournisseurs_canvas.bind("<Enter>", lambda e: _bind_mousewheel(e, fournisseurs_canvas))
            fournisseurs_canvas.bind("<Leave>", _unbind_mousewheel)
        
            
            # Update scroll region
            fournisseurs_subframe.update_idletasks()
            fournisseurs_canvas.configure(scrollregion=fournisseurs_canvas.bbox("all"))

            # Store variables
            selections[year] = fournisseurs_vars
        
        # Add confirmation button
        confirm_button = ttk.Button(
            config_window, 
            text="Confirmer",
            command=lambda: confirm_selection()
        )
        confirm_button.pack(pady=20)
        
        def confirm_selection():
            config = {}
            for year, year_data in selections.items():
                config[year] = [compte for compte, var in year_data.items() if var.get()]
            config_window.result = config
            config_window.destroy()
        
        # Center window on screen
        config_window.update_idletasks()
        width = config_window.winfo_width()
        height = config_window.winfo_height()
        x = (config_window.winfo_screenwidth() // 2) - (width // 2)
        y = (config_window.winfo_screenheight() // 2) - (height // 2)
        config_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Wait for window
        config_window.wait_window()
        return getattr(config_window, 'result', {})

    # Fenêtre de configuration pour le crédit moyen client
    def create_credit_moyen_client_config_window(self, df_by_year,taux_cmc):
        """
        Creates a configuration window for credit moyen client calculations.
        Allows users to select which accounts to include and their VAT rates.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configuration Crédit Moyen Client")
        config_window.geometry("800x600")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Create notebook
        notebook = ttk.Notebook(config_window)
        notebook.pack(expand=True, fill='both', padx=20, pady=20)
        
        # Store all selections
        selections = {}
        
        # Track accounts across all years for synchronization
        all_accounts = {}  # Format: {account_lib: {year: selection_vars}}
        
        # Flag to prevent recursive callbacks during synchronization
        is_syncing = False
        
        # Create tabs for each year
        for year, df in sorted(df_by_year.items()):
            year_frame = ttk.Frame(notebook)
            notebook.add(year_frame, text=f"Année {year}")
            
            # Configure grid
            year_frame.columnconfigure(0, weight=1)
            
            # Create main frame
            main_frame = ttk.LabelFrame(year_frame, text="Sélection des comptes et taux de TVA")
            main_frame.grid(row=0, column=0, padx=15, pady=15, sticky='nsew')
            
            # Create "Select All" button and controls frame
            controls_frame = ttk.Frame(year_frame)
            controls_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky='w')
            
            # Select All checkbox
            select_all_var = tk.BooleanVar(value=False)
            select_all_cb = ttk.Checkbutton(
                controls_frame,
                text="Tout sélectionner",
                variable=select_all_var,
                command=lambda y=year: toggle_all_selections(y)
            )
            select_all_cb.pack(side=tk.LEFT, padx=(0, 20))
            
            # Add TVA dropdown for "Select All"
            ttk.Label(controls_frame, text="TVA par défaut:").pack(side=tk.LEFT, padx=(0, 5))
            default_tva_var = tk.StringVar(value='0%')
            default_tva_combo = ttk.Combobox(
                controls_frame,
                values=taux_cmc,
                textvariable=default_tva_var,
                width=8,
                state='readonly'
            )
            default_tva_combo.set('0%')
            default_tva_combo.pack(side=tk.LEFT)
            default_tva_combo.bind("<<ComboboxSelected>>", 
                                lambda event, y=year: apply_default_tva(y))
            
            # Create canvas with scrollbar
            canvas = tk.Canvas(main_frame)
            scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
            subframe = ttk.Frame(canvas)
            
            # Configure scrolling
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Pack scrollbar and canvas
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create window in canvas
            canvas.create_window((0, 0), window=subframe, anchor='nw')
            
            # Get accounts (only 7xx accounts) with proper duplicate handling
            comptes = df[
                df['CompteNum'].astype(str).str[:3].astype(int).between(700, 709)
            ][['CompteNum', 'CompteLib']].copy()
            
            # Clean the data to prevent duplicate issues
            comptes = comptes.dropna()  # Remove null values
            comptes = comptes[comptes['CompteLib'].str.strip() != '']  # Remove empty labels
            comptes['CompteLib'] = comptes['CompteLib'].str.strip()  # Clean whitespace
            
            # Remove duplicates more aggressively - prioritize by CompteNum first, then CompteLib
            comptes = comptes.drop_duplicates(subset=['CompteNum'], keep='first')
            comptes = comptes.sort_values(['CompteNum', 'CompteLib'])
            
            # Create header
            header_frame = ttk.Frame(subframe)
            header_frame.pack(fill='x', padx=10, pady=5)
            
            ttk.Label(header_frame, text="Compte", width=40).grid(row=0, column=0, padx=5)
            ttk.Label(header_frame, text="TVA", width=10).grid(row=0, column=1, padx=5)
            
            account_vars = {}
            
            # Create checkbuttons and comboboxes for each account
            for _, row in comptes.iterrows():
                frame = ttk.Frame(subframe)
                frame.pack(fill='x', padx=10, pady=2)
                
                # Use CompteNum as the unique key instead of CompteLib to handle duplicates
                compte_lib = row['CompteLib']
                compte_num = row['CompteNum']
                unique_key = f"{compte_num}"  # Use CompteNum as unique identifier
                
                # Skip if we already processed this account number in this year
                if unique_key in account_vars:
                    continue
                
                # Create checkbox variable with explicit initialization
                var = tk.BooleanVar()
                var.set(False)  # Explicitly set to False
                
                # Create checkbox with custom command that forces state
                def create_checkbox_command(uk, y):
                    def checkbox_command():
                        # Force the checkbox to be in a definite state
                        current_state = selections[y]['accounts'][uk]['selected'].get()
                        selections[y]['accounts'][uk]['checkbox'].state(['!alternate'])
                        if current_state:
                            selections[y]['accounts'][uk]['checkbox'].state(['selected'])
                        else:
                            selections[y]['accounts'][uk]['checkbox'].state(['!selected'])
                        
                        # Then handle the synchronization
                        on_account_change(uk, y, 'selected')
                    return checkbox_command
                
                cb = ttk.Checkbutton(
                    frame,
                    text=f"{compte_num} - {compte_lib}",
                    variable=var,
                    command=create_checkbox_command(unique_key, year)
                )
                cb.grid(row=0, column=0, sticky='w')
                
                # Force checkbox to normal state
                cb.state(['!alternate', '!selected'])
                
                # TVA Combobox
                tva_var = tk.StringVar(value='0%')
                tva_combo = ttk.Combobox(
                    frame,
                    values=taux_cmc,
                    textvariable=tva_var,
                    width=8,
                    state='readonly'
                )
                tva_combo.set('0%')
                tva_combo.grid(row=0, column=1, padx=5)
                tva_combo.bind("<<ComboboxSelected>>", 
                            lambda event, uk=unique_key, y=year: on_account_change(uk, y, 'tva'))
                
                account_vars[unique_key] = {
                    'selected': var,
                    'tva': tva_var,
                    'compte_num': compte_num,
                    'compte_lib': compte_lib,
                    'checkbox': cb
                }
                
                # Track accounts across years for synchronization using unique_key
                if unique_key not in all_accounts:
                    all_accounts[unique_key] = {}
                all_accounts[unique_key][year] = {
                    'selected': var,
                    'tva': tva_var,
                    'compte_num': compte_num,
                    'compte_lib': compte_lib,
                    'checkbox': cb
                }
            
            # Enable mousewheel scrolling
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
            # Update scroll region
            subframe.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # Store variables
            selections[year] = {
                'accounts': account_vars,
                'select_all': select_all_var,
                'default_tva': default_tva_var,
                'select_all_cb': select_all_cb
            }
        
        # Function to force checkbox to proper state
        def force_checkbox_state(checkbox, is_selected):
            """Force a checkbox to a definite state, removing any indeterminate state."""
            checkbox.state(['!alternate'])  # Remove indeterminate state
            if is_selected:
                checkbox.state(['selected'])
            else:
                checkbox.state(['!selected'])
        
        # Function to handle account selection or TVA change with auto-sync
        def on_account_change(unique_key, source_year, change_type):
            nonlocal is_syncing
            
            if is_syncing:
                return
            
            try:
                is_syncing = True
                
                # Get current values from source
                source_selected = all_accounts[unique_key][source_year]['selected'].get()
                source_tva = all_accounts[unique_key][source_year]['tva'].get()
                
                # Sync to all other years
                for year, data in all_accounts[unique_key].items():
                    if year != source_year:
                        if change_type == 'selected':
                            data['selected'].set(source_selected)
                            force_checkbox_state(data['checkbox'], source_selected)
                        data['tva'].set(source_tva)
                
                # Update select all states for all years
                for year in selections:
                    update_select_all_state(year)
                    
            finally:
                is_syncing = False
        
        # Function to toggle all selections
        def toggle_all_selections(year):
            nonlocal is_syncing
            
            if is_syncing:
                return
            
            try:
                is_syncing = True
                state = selections[year]['select_all'].get()
                
                # Force select all checkbox to proper state
                force_checkbox_state(selections[year]['select_all_cb'], state)
                
                # Set all accounts in current year
                for unique_key, var in selections[year]['accounts'].items():
                    var['selected'].set(state)
                    force_checkbox_state(var['checkbox'], state)
                
                # Sync to other years
                for unique_key in all_accounts:
                    if year in all_accounts[unique_key]:
                        for other_year in all_accounts[unique_key]:
                            if other_year != year and unique_key in selections[other_year]['accounts']:
                                selections[other_year]['accounts'][unique_key]['selected'].set(state)
                                force_checkbox_state(selections[other_year]['accounts'][unique_key]['checkbox'], state)
                
                # Update select all checkboxes in other years
                for other_year in selections:
                    if other_year != year:
                        update_select_all_state(other_year)
                
                # If selecting all, apply default TVA
                if state:
                    apply_default_tva(year)
                    
            finally:
                is_syncing = False
        
        # Function to apply default TVA to all selected accounts
        def apply_default_tva(year):
            nonlocal is_syncing
            
            if is_syncing:
                return
            
            try:
                is_syncing = True
                tva_value = selections[year]['default_tva'].get()
                
                # Set TVA for all selected accounts in current year
                for unique_key, var in selections[year]['accounts'].items():
                    if var['selected'].get():
                        var['tva'].set(tva_value)
                
                # Sync to other years
                for unique_key in all_accounts:
                    if year in all_accounts[unique_key] and selections[year]['accounts'][unique_key]['selected'].get():
                        for other_year in all_accounts[unique_key]:
                            if other_year != year and unique_key in selections[other_year]['accounts']:
                                other_account = selections[other_year]['accounts'][unique_key]
                                if other_account['selected'].get():
                                    other_account['tva'].set(tva_value)
            finally:
                is_syncing = False
        
        # Function to update "Select All" state
        def update_select_all_state(year):
            nonlocal is_syncing
            if is_syncing:
                return
            
            try:
                # Temporarily prevent syncing during state update
                is_syncing = True
                
                all_selected = all(var['selected'].get() for var in selections[year]['accounts'].values())
                none_selected = not any(var['selected'].get() for var in selections[year]['accounts'].values())
                
                if all_selected:
                    selections[year]['select_all'].set(True)
                    force_checkbox_state(selections[year]['select_all_cb'], True)
                elif none_selected:
                    selections[year]['select_all'].set(False)
                    force_checkbox_state(selections[year]['select_all_cb'], False)
                else:
                    # Partial selection - use indeterminate state only for "Select All"
                    selections[year]['select_all'].set(False)
                    selections[year]['select_all_cb'].state(['alternate'])
                    
            finally:
                is_syncing = False
        
        # Add confirmation button
        confirm_button = ttk.Button(
            config_window,
            text="Confirmer",
            command=lambda: confirm_selection()
        )
        confirm_button.pack(pady=20)
        
        def confirm_selection():
            config = {}
            for year, year_data in selections.items():
                config[year] = {
                    unique_key: {
                        'selected': var['selected'].get(),
                        'tva': float(var['tva'].get().strip('%')) / 100,
                        'compte_num': var['compte_num'],
                        'compte_lib': var['compte_lib']
                    }
                    for unique_key, var in year_data['accounts'].items()
                    if var['selected'].get()
                }
            config_window.result = config
            config_window.destroy()
        
        # Center window
        config_window.update_idletasks()
        width = config_window.winfo_width()
        height = config_window.winfo_height()
        x = (config_window.winfo_screenwidth() // 2) - (width // 2)
        y = (config_window.winfo_screenheight() // 2) - (height // 2)
        config_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Wait for window
        config_window.wait_window()
        return getattr(config_window, 'result', {})

    def report_generator_config(self, revue_ana_path):
        """Creates and manages the report generator configuration window."""
        window = tk.Toplevel(self)
        window.title("Générateur de reporting")
        window.geometry("900x800")
        window.transient(self)
        window.grab_set()
        self.center_config_window(window)


        window.configure(bg='#f0f2f5')
        
        style = ttk.Style()
        style.configure('Container.TFrame', background='#f0f2f5')
        style.configure('Card.TFrame', background='#ffffff')
        style.configure('SectionTitle.TLabel', font=('Inter', 12, 'bold'), background='#ffffff', foreground='#1e293b')
        style.configure('NormalText.TLabel', font=('Inter', 10), background='#ffffff', foreground='#1e293b')
        
        container = ttk.Frame(window, style='Container.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=40, pady=30)
        
        card_frame = ttk.Frame(container, style='Card.TFrame')
        card_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        file_frame = ttk.Frame(card_frame, style='Card.TFrame')
        file_frame.pack(fill=tk.X, padx=30, pady=10)
        
        base_name = Path(revue_ana_path).stem
        ttk.Label(file_frame, text="Fichier Excel:", style="SectionTitle.TLabel").pack(anchor='w', pady=(0, 5))
        file_label = ttk.Label(file_frame, text=base_name, style="NormalText.TLabel")
        file_label.pack(side=tk.LEFT, pady=10)
        
        template_path = resource_path("assets\template_reporting.docx")
        
        template_frame = ttk.Frame(card_frame, style='Card.TFrame')
        template_frame.pack(fill=tk.X, padx=30, pady=10)
        ttk.Label(template_frame, text="Template de reporting:", style="SectionTitle.TLabel").pack(anchor='w', pady=(0, 5))
        template_label = ttk.Label(template_frame, text= f"Par défaut: {template_path}", style="NormalText.TLabel", 
                           wraplength=300,  # Adjust this value based on your window width
                           justify=tk.LEFT)
        template_label.pack(side=tk.LEFT, pady=10)

        def choose_template():
            nonlocal template_path
            new_template_path = filedialog.askopenfilename(
                title="Sélectionner la template de reporting",
                filetypes=[("Fichiers Word", "*.docx")],initialfile=template_path
            )
            if new_template_path:
                template_path = new_template_path
                template_label.config(text=f"Emplacement: {new_template_path}")

        ttk.Button(template_frame, text="Choisir la template", command=choose_template).pack(side=tk.RIGHT, pady=10)

        output_frame = ttk.Frame(card_frame, style='Card.TFrame')
        output_frame.pack(fill=tk.X, padx=30, pady=10)
        ttk.Label(output_frame, text="Emplacement de sortie:", style="SectionTitle.TLabel").pack(anchor='w', pady=(0, 5))
        output_label = ttk.Label(output_frame, text="Par défaut: ./generated_reports", style="NormalText.TLabel")
        output_label.pack(side=tk.LEFT, pady=10)
        
        output_folder = os.path.join(os.getcwd(), "generated_reports")
        os.makedirs(output_folder, exist_ok=True)
        
        def choose_output_location():
            nonlocal output_folder
            folder_path = filedialog.askdirectory(title="Choisir l'emplacement de sortie", initialdir=output_folder)
            if folder_path:
                output_folder = folder_path
                output_label.config(text=f"Emplacement: {folder_path}")
        
        ttk.Button(output_frame, text="Choisir l'emplacement", command=choose_output_location).pack(side=tk.RIGHT, pady=10)
        
        client_date_frame = ttk.Frame(card_frame, style='Card.TFrame')
        client_date_frame.pack(fill=tk.X, padx=30, pady=10)
        
        ttk.Label(client_date_frame, text="Client:", style="SectionTitle.TLabel").pack(anchor='w', pady=(0, 5))
        client_var = tk.StringVar()
        client_entry = ttk.Entry(client_date_frame, textvariable=client_var, width=40)
        client_entry.pack(anchor='w', pady=(0, 15))
        
        ttk.Label(client_date_frame, text="Date:", style="SectionTitle.TLabel").pack(anchor='w', pady=(0, 5))
        date_var = tk.StringVar(value=datetime.datetime.today().strftime("%d/%m/%Y"))
        date_entry = ttk.Entry(client_date_frame, textvariable=date_var, width=40)
        date_entry.pack(anchor='w', pady=(0, 10))
 

        ttk.Label(
            card_frame,
            text="Veuillez fermer le fichier Excel avant de générer le reporting",
            font=('Inter', 10, 'italic')
        ).pack(pady=5)
        
        action_frame = ttk.Frame(card_frame, style='Card.TFrame')
        action_frame.pack(fill=tk.X, padx=30, pady=10)
        
        # progress = ttk.Progressbar(action_frame, mode='indeterminate', length=400)
        # status_label = ttk.Label(action_frame, text="", style="SectionTitle.TLabel")
        # status_label.pack(pady=5)

        def create_circular_loading(parent):
            """
            Create a simple circular loading indicator
            
            Args:
                parent: The parent widget to place the loading indicator
            
            Returns:
                A Label widget that can be animated to simulate loading
            """
            loading_label = ttk.Label(parent, text="⧗", style="SectionTitle.TLabel", font=("Arial", 20))
            
            # Animation function to create spinning effect
            def animate_loading(label, cycle=0):
                # Unicode characters for spinning effect
                spinner_chars = ['⧗', '⌛', '⏳']
                
                if hasattr(label, '_animation_running'):
                    label.config(text=spinner_chars[cycle % len(spinner_chars)])
                    label.after(200, lambda: animate_loading(label, cycle + 1))
            
            def start_loading(label):
                label._animation_running = True
                animate_loading(label)
            
            def stop_loading(label):
                label._animation_running = False
                label.config(text="")
            
            # Attach methods to the label
            loading_label.start = lambda: start_loading(loading_label)
            loading_label.stop = lambda: stop_loading(loading_label)
            
            return loading_label

        progress = create_circular_loading(action_frame)
        status_label = ttk.Label(action_frame, text="", style="SectionTitle.TLabel")
        status_label.pack(pady=5)

        def process_file():
            
            progress.pack(pady=5)
            progress.start()
            status_label.config(text="Génération du reporting...")
            
            client = client_var.get().strip()
            date = date_var.get().strip()
            
            output_path = os.path.join(output_folder, f'{" ".join(list(reversed(date.split("/"))))} {client} Reporting.docx')
            
            # thread = threading.Thread(target=generate_report, args=(revue_ana_path,output_path,config_elements_fixes,config_elements_variables,client,date))
            # thread.start()
            
            def threaded_report_generation():
                try:
                    generate_report(revue_ana_path, output_path, config_elements_fixes, config_elements_variables, client, date)
                    # When done, stop the loading and update status
                    progress.stop()
                    status_label.config(text="Reporting généré avec succès!")
                except Exception as e:
                    # Handle any errors
                    progress.stop()
                    status_label.config(text=f"Erreur: {str(e)}")

            
            thread = threading.Thread(target=threaded_report_generation)
            thread.start()
       
        def generate_report(excel_path, output_path,config_elements_fixes,config_elements_variables,client,date):
            try:
                # Placeholder for actual report generation function
                insert_images_from_excel_to_word(excel_path, template_path, output_path,config_elements_fixes,config_elements_variables, client, date)
                window.after(0, handle_success, output_path)
            except Exception as e:
                window.after(0, handle_error, str(e))
        
        def handle_success(output_path):
            progress.stop()
            progress.pack_forget()
            status_label.config(text="Reporting généré avec succès!")
            if messagebox.askyesno("Succès", "Reporting généré avec succès! Voulez-vous l'ouvrir?"):
                open_file(output_path)
        
        def handle_error(error_message):
            progress.stop()
            progress.pack_forget()
            status_label.config(text="Erreur lors de la génération du reporting")
            messagebox.showerror("Erreur", f"Erreur lors de la génération du reporting: {error_message}")
        
        def open_file(file_path):
            if os.path.exists(file_path):
                try:
                    if os.name == 'nt':
                        os.startfile(file_path)
                    else:
                        subprocess.run(['open' if os.name == 'darwin' else 'xdg-open', file_path])
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de l'ouverture du fichier: {str(e)}")
            else:
                messagebox.showerror("Erreur", "Fichier non trouvé")
        
        generate_button = ttk.Button(action_frame, text="Générer le reporting", command=process_file)
        generate_button.pack(pady=10)
        
        window.transient(self)
        window.grab_set()
        window.protocol("WM_DELETE_WINDOW", window.destroy)

        return window

    # Fonction pour afficher les détails de l'erreur
    def show_error_details(self):
        """Shows error details in a new window"""
        if self.error_details:
            details_window = tk.Toplevel(self)
            details_window.title("Détails de l'erreur")
            details_window.geometry("600x400")
            
            # Create text widget with scrollbar
            text_frame = tk.Frame(details_window)
            text_frame.pack(fill='both', expand=True, padx=10, pady=10)
            
            scrollbar = tk.Scrollbar(text_frame)
            scrollbar.pack(side='right', fill='y')
            
            text_widget = tk.Text(
                text_frame,
                wrap=tk.WORD,
                yscrollcommand=scrollbar.set,
                font=('Courier', 10)
            )
            text_widget.pack(side='left', fill='both', expand=True)
            
            scrollbar.config(command=text_widget.yview)
            
            # Insert error details
            text_widget.insert('1.0', self.error_details)
            text_widget.configure(state='disabled')
            
            # Add close button
            close_button = tk.Button(
                details_window,
                text="Fermer",
                command=details_window.destroy
            )
            close_button.pack(pady=10)

    # Fonction pour gérer les erreurs
    def handle_error(self, error, error_type="Erreur"):
        """Centralized error handling"""
        error_message = str(error)
        self.error_details = f"{error_type}:\n\n{error_message}\n\nTraceback:\n{traceback.format_exc()}"
        
        # Show error message
        messagebox.showerror(error_type, error_message)
        
        # Update status and error button
        self.status_label.config(text=f"{error_type} - voir détails pour plus d'informations")
        self.error_details_button.config(state='normal')
        
        # Log error
        print(f"{error_type}: {error_message}")
        traceback.print_exc()

    def process_files(self):
        """
        Traite les fichiers FEC sélectionnés.
        """
        if not self.fec_files:
            messagebox.showwarning("Attention", "Aucun fichier à traiter")
            return
        
        processed_files = []
        liste_encoding = self.parametres['liste_encoding']
        for year, file_path,_ in self.fec_files:
            try:                
                # Vérifier les doublons d'années
                for processed_year, _ in processed_files:
                    if processed_year == year:
                        raise ValueError(f"Plusieurs fichiers pour l'année {year}")
                
                processed_files.append((year, file_path))
                
            except ValueError as ve:
                messagebox.showerror("Erreur", str(ve))
                return

        fournisseurs_codes = [item['value'] for item in self.parametres['fournisseurs_code'] if item['selected']]
        clients_codes = [item['value'] for item in self.parametres['clients_code'] if item['selected']]
        journal_code_achats = [item['value'] for item in self.parametres['journal_codes_achats'] if item['selected']]
        journal_code_ventes = [item['value'] for item in self.parametres['journal_codes_ventes'] if item['selected']]
        taux_cmc = self.parametres['taux_cmc']

        # Show configuration window first
        config = self.create_company_config_window()
        
        # If window was closed without confirming
        if not config:
            return
        
        # Determine whether to exclude salaries
        exclude_salaries = config['exclude_salaries']
        reporting = config["reporting"]
        exclude_cpte_fournisseur = config['exclude_cpte_fournisseur']
                
        account_config = {}
        excluded_cmf_config = {}
        cmc_account_config = {}
        top_dix_account_config = {}

        try:
            self.status_label.config(text="Traitement en cours...")
            self.update()
            # Reset error details
            self.error_details = None
            self.error_details_button.config(state='disabled')
            # Load all DataFrames first
            df_by_year = {}
            # Désactiver les boutons pendant le traitement
            self.select_files_button.config(state='disabled')
            self.process_button.config(state='disabled')
            self.delete_button.config(state='disabled')
            self.delete_all_button.config(state='disabled')

            # Créer le workbook
            wb = Workbook()

            # Dictionnaires pour stocker les DataFrames par année
            all_results = {
                'Tresorerie': [],
                'CMF': [],
                'Top_FR': [],
                "Achats et Sous-traitance": [],
                'CMC': [],
                'Top_CLT': [],
                'Frais_Personnel':[]
            }
            
            # Liste pour stocker les DataFrames pour chaque année
            ca_dataframes = []
            ca_dataframes_cumule = []

            cmf_jours_dataframes = []
            cmf_top_5_by_year_dataframes = []
            cmf_top_5_jours_dataframes = {}

            treso_brute_dataframes = []
            treso_nette_dataframes = []

            achats_et_st_dataframes = []
            achats_et_st_dataframes_cumules = []

            table_investissement_dataframes = []
            total_investissement_dataframes = []

            cmc_jours_dataframes = []
            cmc_top_5_by_year_dataframes = []
            cmc_top_5_jours_dataframes = {}

            frais_personnels_cumules_dataframes = []
            frais_personnels_annuel_dataframes = []

            max_length_tresorerie = 0
            max_length_cmc = 0
            previous_year_data_cmf = None
            previous_year_data_cmc = None

            # Initialize dictionaries to store data by year
            cmf_by_year = {}
            cmc_by_year = {}
            
            # Load all DataFrames first
            df_by_year = {}
            for year, file_path,_ in sorted(self.fec_files, key=lambda x: x[0], reverse=False):
                i = 0
                success = False
                while i < len(liste_encoding) and not success:
                    try:
                        df = pd.read_csv(file_path, sep='\t', encoding=liste_encoding[i], engine='python')
                        df_by_year[year] = df
                        success = True
                    except (UnicodeDecodeError,UnicodeError):
                        i += 1
            
            # Open the account configuration window
            # Only open the salary config window if user requested to exclude salaries
            if exclude_salaries:
                account_config = self.create_frais_personnels_config_window(df_by_year)
            if exclude_cpte_fournisseur and config['sheets']['Crédit Moyen FR']:
                excluded_cmf_config = self.exclusion_credit_moyen(df_by_year,
                                                                  fournisseurs_codes,
                                                                  journal_code_achats
                                                                  )
            if config['sheets']['Crédit Moyen Client']:
                cmc_account_config = self.create_credit_moyen_client_config_window(df_by_year,taux_cmc)
            if config['sheets']['TOP FR'] or config['sheets']['TOP Client']: 
                top_dix_account_config = self.create_top_dix_config_window(df_by_year,
                                                                           fournisseurs_codes=fournisseurs_codes,
                                                                           clients_codes=clients_codes,
                                                                           journal_code_achats=journal_code_achats,
                                                                           journal_code_ventes=journal_code_ventes
                                                                           )

            for year, df in df_by_year.items():
                    
                # Préparation des données
                df.columns = df.columns.str.strip()
                df["EcritureDate"] = pd.to_datetime(df["EcritureDate"], format='%Y%m%d', errors='coerce')
                df['Année'] = df["EcritureDate"].dt.year
                df['Mois'] = df["EcritureDate"].dt.month_name(locale='French')
                df["Credit"] = pd.to_numeric(df["Credit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0)
                df["Debit"] = pd.to_numeric(df["Debit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0)

                cmc_top_5_jours_dataframes[year] = []
                cmf_top_5_jours_dataframes[year] = []

                # Calculs pour l'année en cours
                if config['sheets']['CA']:
                    ca_mensuel, ca_mensuel_cumule = chiffres_affaires(df)
                    ca_dataframes.append(ca_mensuel)
                    ca_dataframes_cumule.append(ca_mensuel_cumule)

                if config['sheets']['Achats et Sous-traitance']:
                    table_achat_sous_traitance, table_achat_sous_traitance_cumule = achats_et_sous_traitances(df)
                    achats_et_st_dataframes.append(table_achat_sous_traitance[["Mois", "Total"]].rename(columns={"Total": str(year)}))
                    achats_et_st_dataframes_cumules.append(table_achat_sous_traitance_cumule[["Mois", "Total"]].rename(columns={"Total": str(year)}))
                
                if config['sheets']['Trésorerie']:
                    tresorerie_resultat = tresorerie(df)
                    treso_brute_df = tresorerie_resultat[["Mois", "Trésorerie Brute"]].rename(columns={"Trésorerie Brute": str(year)})
                    treso_nette_df = tresorerie_resultat[["Mois", "Trésorerie nette des emprunts et CC"]].rename(columns={"Trésorerie nette des emprunts et CC": str(year)})
                    treso_brute_dataframes.append(treso_brute_df)
                    treso_nette_dataframes.append(treso_nette_df)
                    max_length_tresorerie = max(max_length_tresorerie, tresorerie_resultat.shape[1])
                    all_results['Tresorerie'].append((f"Evolution de la trésorerie {year}", tresorerie_resultat))

                if config['sheets']['Investissement']:
                    table_investissement,total_investissement = investissement(df,year)
                    table_investissement_dataframes.append(table_investissement)
                    total_investissement_dataframes.append(total_investissement)

                if config['sheets']['Frais personnels']:
                    if exclude_salaries:
                        year_config = account_config.get(year, {})
                        excluded_salaires = year_config.get('excluded_salaires', [])
                        excluded_charges = year_config.get('excluded_charges', [])
                        
                        # Use the excluded accounts in your calculations
                        salaires = df[
                            (df['CompteNum'].astype(str).str.startswith('641')) & 
                            (~df['CompteLib'].isin(excluded_salaires))
                        ]
                        
                        charges_sociales = df[
                            (df['CompteNum'].astype(str).str.startswith('64')) & 
                            (~df['CompteNum'].astype(str).str.startswith('641')) & 
                            (~df['CompteLib'].isin(excluded_charges))
                        ]
                    else:
                        salaires = df[df['CompteNum'].astype(str).str.startswith('641')]
                        charges_sociales = df[df['CompteNum'].astype(str).str.startswith('64') & ~df['CompteNum'].astype(str).str.startswith('641')]
                    
                    frais_personnels_resultat, frais_personnels_resultat_cumules = frais_personnels(salaires, charges_sociales)
                    frais_personnels_cumules_dataframes.append(frais_personnels_resultat_cumules[["Mois", "Total Salaires et Charges Sociales"]].rename(columns={"Total Salaires et Charges Sociales": year}))
                    frais_personnels_annuel_dataframes.append(frais_personnels_resultat[["Mois", "Total Salaires et Charges Sociales"]].rename(columns={"Total Salaires et Charges Sociales": year}))
                    all_results['Frais_Personnel'].append((f"Evolution des frais de personnel {year}", frais_personnels_resultat))
                                           
                if config['sheets']['Crédit Moyen FR']:
                    previous_year_data_cmf = cmf_by_year.get("-".join([str(int(year)-1) for year in str(year).split("-")]))
                    cmf_data = df
                    if exclude_cpte_fournisseur:
                        excluded_cmf_list = excluded_cmf_config.get(year,[])
                        cmf_data = df[~(df['CompAuxLib'].isin(excluded_cmf_list))] if len(excluded_cmf_list) != 0 else df
                    
                    cmf = credit_moyen_fournisseur(cmf_data,
                                                   previous_year_data=previous_year_data_cmf,
                                                   fournisseurs_codes=fournisseurs_codes,
                                                   clients_codes=clients_codes,
                                                   journal_code_achats=journal_code_achats
                                                   )
                    cmf_by_year[year] = cmf
                    cmf_jours = cmf[["Mois", "Jours"]].rename(columns={"Jours": f"{year}"}).iloc[:-1]
                    moyenne = cmf_jours[f"{year}"].replace(0,pd.NA).mean(skipna=True)
                    cmf_jours.loc[len(cmf_jours)] = ["Moyenne", custom_round(moyenne) if not pd.isna(moyenne) else '']
                    cmf_jours_dataframes.append(cmf_jours)
                    all_results['CMF'].append((f"Crédit Moyen Fournisseurs {year}", cmf))
                
                if config['sheets']['Crédit Moyen Client']:
                    previous_year_data_cmc = cmc_by_year.get("-".join([str(int(year)-1) for year in str(year).split("-")]))
                    # Calculate CMC and store it
                    cmc = credit_moyen_client(df,
                                              annee=year, 
                                              previous_year_data=previous_year_data_cmc, 
                                              account_config=cmc_account_config,
                                              fournisseurs_codes=fournisseurs_codes,
                                              clients_codes=clients_codes
                                              )
                    max_length_cmc = max(max_length_cmc,cmc.shape[1])
                    cmc_by_year[year] = cmc
                    cmc_jours = cmc[["Mois", "Jours"]].rename(columns={"Jours": f"{year}"}).iloc[:-1]
                    moyenne = cmc_jours[f"{year}"].replace(0,pd.NA).mean(skipna=True)
                    cmc_jours.loc[len(cmc_jours)] = ["Moyenne", custom_round(moyenne) if not pd.isna(moyenne) else '']
                    cmc_jours_dataframes.append(cmc_jours)
                    all_results['CMC'].append((f"Crédit Moyen Clients {year}", cmc))
                
                if config['sheets']['TOP FR']:
                    year_config_top_dix = top_dix_account_config.get(year, {})
                    excluded_fournisseurs = year_config_top_dix.get('excluded_fournisseurs', [])
                    fournisseurs = df[~(df['CompAuxLib'].isin(excluded_fournisseurs))] if len(excluded_fournisseurs) != 0 else df
                    if fournisseurs.shape == df.shape:
                        fournisseurs = df[~(df['CompteLib'].isin(excluded_fournisseurs))] if len(excluded_fournisseurs) != 0 else df
                    try:
                        # On utilise pas top_dix_hors_taxes pour l'instant
                        top_dix_fournisseur,compte_lib_column_fournisseur = top_dix(fournisseurs, 
                                                                                               type = 'Fournisseurs',
                                                                                               fournisseurs_codes=fournisseurs_codes,
                                                                                               clients_codes=clients_codes,
                                                                                               journal_code_achats=journal_code_achats,
                                                                                               journal_code_ventes=journal_code_ventes
                                                                                               )
                    except KeyError:
                        top_dix_fournisseur,compte_lib_column_fournisseur = top_dix(fournisseurs, 
                                                                                    type = 'Fournisseurs',
                                                                                    fournisseurs_codes=fournisseurs_codes,
                                                                                    clients_codes=clients_codes,
                                                                                    journal_code_achats=journal_code_achats,
                                                                                    journal_code_ventes=journal_code_ventes
                                                                                    )
                    percentage_column_fournisseur = "Pourcentage HT" if "Pourcentage HT" in top_dix_fournisseur.columns else "Pourcentage"
                    top_cinq_fournisseur = list(top_dix_fournisseur[compte_lib_column_fournisseur].head(5))
                    all_results['Top_FR'].append((f"Top 10 Fournisseurs {year}", top_dix_fournisseur))
                
                if config['sheets']['TOP Client']:
                    year_config_top_dix = top_dix_account_config.get(year, {})
                    excluded_clients = year_config_top_dix.get('excluded_clients', [])
                    clients = df[~(df['CompAuxLib'].isin(excluded_clients))] if len(excluded_clients) != 0 else df
                    try:
                        # On utilise pas top_dix_hors_taxes pour l'instant
                        top_dix_client,compte_lib_column_client = top_dix(clients, 
                                                                                     type ='Clients',
                                                                                     fournisseurs_codes=fournisseurs_codes,
                                                                                     clients_codes=clients_codes,
                                                                                     journal_code_achats=journal_code_achats,
                                                                                     journal_code_ventes=journal_code_ventes
                                                                                    )
                    except KeyError:
                        top_dix_client,compte_lib_column_client = top_dix(clients, 
                                                                        type ='Clients',
                                                                        fournisseurs_codes=fournisseurs_codes,
                                                                        clients_codes=clients_codes,
                                                                        journal_code_achats=journal_code_achats,
                                                                        journal_code_ventes=journal_code_ventes
                                                                    )
                    percentage_column_client = "Pourcentage HT" if "Pourcentage HT" in top_dix_client.columns else "Pourcentage"
                    top_cinq_clients = list(top_dix_client[compte_lib_column_client].head(5))
                    all_results['Top_CLT'].append((f"Top 10 Clients {year}", top_dix_client))
                
                if config['sheets']['Achats et Sous-traitance']:
                    all_results["Achats et Sous-traitance"].append((f"Evolutions des achats et de la sous-traitance {year}", table_achat_sous_traitance))

                if config['sheets']['Crédit Moyen Client'] and config['sheets']['TOP Client']:
                    cmc_top_5_dataframes_with_titles = [credit_moyen_client_top_cinq(df,
                                                                                     client,annee = year,
                                                                                     previous_year_data=previous_year_data_cmc,
                                                                                     account_config=cmc_account_config,
                                                                                     fournisseurs_codes=fournisseurs_codes,
                                                                                     clients_codes=clients_codes
                                                                                     ) 
                                                        for client in top_cinq_clients
                                                        ]
                    for title, cmc_top_5 in cmc_top_5_dataframes_with_titles:
                        client = title.split("Crédit moyen client")[-1]
                        cmc_top_5_jours = cmc_top_5[["Mois", "Jours"]].rename(columns={"Jours": f"{client}"}).iloc[:-1]
                        moyenne = cmc_top_5_jours[f"{client}"].replace(0,pd.NA).mean(skipna=True)
                        cmc_top_5_jours.loc[len(cmc_top_5_jours)] = ["Moyenne", custom_round(moyenne) if not pd.isna(moyenne) else '']
                        cmc_top_5_jours_dataframes[year].append(cmc_top_5_jours)
                    cmc_top_5_by_year_dataframes.append(cmc_top_5_dataframes_with_titles)

                if config['sheets']['Crédit Moyen FR'] and config['sheets']['TOP FR']:
                    cmf_top_5_dataframes_with_titles = [credit_moyen_fournisseur_top_cinq(df,
                                                                                          year,
                                                                                          fournisseur,
                                                                                          previous_year_data=previous_year_data_cmf,
                                                                                          fournisseurs_codes=fournisseurs_codes,
                                                                                          clients_codes=clients_codes,
                                                                                          journal_code_achats=journal_code_achats) 
                                                        for fournisseur in top_cinq_fournisseur
                                                        ]
                    for title, cmf_top_5 in cmf_top_5_dataframes_with_titles:
                        fournisseur = title.split("Crédit moyen fournisseur")[-1]
                        cmf_top_5_jours = cmf_top_5[["Mois", "Jours"]].rename(columns={"Jours": f"{fournisseur}"}).iloc[:-1]
                        moyenne = cmf_top_5_jours[f"{fournisseur}"].replace(0,pd.NA).mean(skipna=True)
                        cmf_top_5_jours.loc[len(cmf_top_5_jours)] = ["Moyenne", custom_round(moyenne) if not pd.isna(moyenne) else '']
                        cmf_top_5_jours_dataframes[year].append(cmf_top_5_jours)
                    cmf_top_5_by_year_dataframes.append(cmf_top_5_dataframes_with_titles)
                
            # Renverser l'ordre pour avoir les années dans l'ordre croissant
            for key in all_results:
                if all_results[key]:
                    all_results[key] = list(reversed(all_results[key]))

            # Fusionner tous les DataFrames en un seul
            if config['sheets']['CA'] and ca_dataframes:
                ca_final = pd.concat(ca_dataframes, axis=1)
                ca_final_cumule = pd.concat(ca_dataframes_cumule, axis=1)    
                # Supprimer les colonnes en double (comme "Mois")
                columns_to_keep = ['Mois'] + [col for col in ca_final.columns if col != 'Mois']
                ca_final = ca_final[columns_to_keep]
                ca_final_cumule = ca_final_cumule[columns_to_keep]
                ca_final = ca_final.loc[:, ~ca_final.columns.duplicated()]
                ca_final_cumule = ca_final_cumule.loc[:, ~ca_final_cumule.columns.duplicated()]

            if config['sheets']['Crédit Moyen FR'] and cmf_jours_dataframes:
                cmf_jours_dataframes = pd.concat(cmf_jours_dataframes, axis=1)
                cmf_jours_dataframes = cmf_jours_dataframes.loc[:, ~cmf_jours_dataframes.columns.duplicated()]

            if config['sheets']['Crédit Moyen FR'] and cmf_top_5_jours_dataframes and config['sheets']['TOP FR']:
                for year,_ in cmc_top_5_jours_dataframes.items():
                    cmf_top_5_jours_dataframes[year] = pd.concat(cmf_top_5_jours_dataframes[year], axis=1)
                    cmf_top_5_jours_dataframes[year] = cmf_top_5_jours_dataframes[year].loc[:, ~cmf_top_5_jours_dataframes[year].columns.duplicated()]

            if config['sheets']['Trésorerie'] and treso_brute_dataframes:
                treso_brute_dataframes = pd.concat(treso_brute_dataframes, axis=1)
                treso_nette_dataframes = pd.concat(treso_nette_dataframes, axis=1)
                treso_brute_dataframes = treso_brute_dataframes.loc[:, ~treso_brute_dataframes.columns.duplicated()]
                treso_nette_dataframes = treso_nette_dataframes.loc[:, ~treso_nette_dataframes.columns.duplicated()]

            if config['sheets']['Achats et Sous-traitance'] and achats_et_st_dataframes:
                achats_et_st_dataframes = pd.concat(achats_et_st_dataframes, axis=1)
                achats_et_st_dataframes_cumules = pd.concat(achats_et_st_dataframes_cumules, axis=1)
                achats_et_st_dataframes = achats_et_st_dataframes.loc[:, ~achats_et_st_dataframes.columns.duplicated()]
                achats_et_st_dataframes_cumules = achats_et_st_dataframes_cumules.loc[:, ~achats_et_st_dataframes_cumules.columns.duplicated()]

            if config['sheets']['Investissement']:
                table_investissement_dataframes = pd.concat(table_investissement_dataframes,axis = 1)
                total_investissement_dataframes = pd.concat(total_investissement_dataframes,axis = 0)
                table_investissement_dataframes = table_investissement_dataframes.loc[:, ~table_investissement_dataframes.columns.duplicated()]
                total_investissement_dataframes = total_investissement_dataframes.loc[:, ~total_investissement_dataframes.columns.duplicated()]
                
                table_investissement_dataframes.loc[len(table_investissement_dataframes)] = ['','Total'] + [table_investissement_dataframes[column].sum() for column in table_investissement_dataframes.columns[2:]]

            if config['sheets']['Crédit Moyen Client'] and cmc_jours_dataframes:
                cmc_jours_dataframes = pd.concat(cmc_jours_dataframes, axis=1)
                cmc_jours_dataframes = cmc_jours_dataframes.loc[:, ~cmc_jours_dataframes.columns.duplicated()]
            
            if config['sheets']['Crédit Moyen Client'] and cmc_top_5_jours_dataframes and config['sheets']['TOP Client']:
                for year,_ in cmc_top_5_jours_dataframes.items():
                    cmc_top_5_jours_dataframes[year] = pd.concat(cmc_top_5_jours_dataframes[year], axis=1)
                    cmc_top_5_jours_dataframes[year] = cmc_top_5_jours_dataframes[year].loc[:, ~cmc_top_5_jours_dataframes[year].columns.duplicated()]
            
            if config['sheets']['Frais personnels'] and frais_personnels_annuel_dataframes:
                frais_personnels_final_cumules = pd.concat(frais_personnels_cumules_dataframes, axis=1)
                frais_personnels_annuel = pd.concat(frais_personnels_annuel_dataframes, axis=1)
                frais_personnels_final_cumules = frais_personnels_final_cumules.loc[:, ~frais_personnels_final_cumules.columns.duplicated()]
                frais_personnels_annuel = frais_personnels_annuel.loc[:, ~frais_personnels_annuel.columns.duplicated()]

            # Créer les feuilles Excel avec les résultats
            if config['sheets']['CA'] and ca_dataframes:
                add_df_to_sheet(wb, "CA", ("Evolution du CA mensuel", ca_final), ("Evolution du CA mensuel cumulé", ca_final_cumule), plot=True, remove_total=True)
            
            if config['sheets']['Trésorerie'] and all_results['Tresorerie']:
                add_df_to_sheet(wb, "Trésorerie", *all_results['Tresorerie'],spacing=4)
                if treso_brute_dataframes is not None and treso_nette_dataframes is not None:
                    add_df_to_sheet(wb, "Trésorerie", ("Evolution de la trésorerie brute", treso_brute_dataframes), 
                                ("Evolution de la trésorerie nette", treso_nette_dataframes), 
                                start_column=max_length_tresorerie + 4, plot=True, remove_total=False)

            if config['sheets']['Achats et Sous-traitance'] and all_results["Achats et Sous-traitance"]:
                add_df_to_sheet(wb, "Achats et Sous-traitance", *all_results["Achats et Sous-traitance"])
                if achats_et_st_dataframes is not None and achats_et_st_dataframes_cumules is not None:
                    add_df_to_sheet(wb, "Achats et Sous-traitance", ("Evolution des achats et sous-traitance", achats_et_st_dataframes),
                                ("Evolution des achats et sous-traitance cumulés", achats_et_st_dataframes_cumules),
                                plot=True, remove_total=False, start_column=7)
                    
            if config['sheets']['Investissement']:
                add_df_to_sheet(wb, "Investissement", (f"Evolution des investissements {year}", table_investissement_dataframes), ("Total des investissements par année", total_investissement_dataframes), plot=True,remove_total=True,chart_type="barchart",x_axis ="CompteLib")

            if config['sheets']['Frais personnels'] and all_results['Frais_Personnel']:
                add_df_to_sheet(wb, "Salaires", *all_results['Frais_Personnel'])
                if frais_personnels_annuel is not None and frais_personnels_final_cumules is not None:
                    add_df_to_sheet(wb, "Salaires", ("Evolution mensuelle des frais de personnel", frais_personnels_annuel), 
                                ("Evolution mensuelle des frais de personnel cumulés", frais_personnels_final_cumules), 
                                start_column=8, plot=True, remove_total=True)
        
            if config['sheets']['Crédit Moyen Client'] and all_results['CMC']:
                add_df_to_sheet(wb, "Crédit Moyen CLT", *all_results['CMC'])
                if cmc_jours_dataframes is not None:
                    add_df_to_sheet(wb, "Crédit Moyen CLT", ("Evolution du Crédit moyen client en jours", cmc_jours_dataframes), 
                                start_column=max_length_cmc + 4, plot=True, remove_total=True)
            
            if config['sheets']['Crédit Moyen Client'] and cmc_top_5_by_year_dataframes and config['sheets']['TOP Client']:
                add_df_to_sheet(wb, "Crédit Moyen CLT TOP 5", *cmc_top_5_by_year_dataframes[-1])
                if len(cmc_top_5_by_year_dataframes) > 1:
                    add_df_to_sheet(wb, "Crédit Moyen CLT TOP 5", *cmc_top_5_by_year_dataframes[-2],start_column=max_length_cmc+3)
                start_column = 2*(max_length_cmc + 3) if len(cmc_top_5_by_year_dataframes) > 1 else max_length_cmc + 3
                add_df_to_sheet(wb, "Crédit Moyen CLT TOP 5", ("Evolution du Crédit moyen client pour le top 5 des clients en jours",next(reversed(cmc_top_5_jours_dataframes.values()))),start_column=start_column,plot = True)
            
            if config['sheets']['TOP Client'] and all_results['Top_CLT']:
                add_df_to_sheet(wb, "TOP CLT", *all_results['Top_CLT'], plot=True, x_axis=compte_lib_column_client, 
                            chart_type="pie", percentage_col=percentage_column_client, remove_total=True)

            if config['sheets']['Crédit Moyen FR'] and all_results['CMF']:
                add_df_to_sheet(wb, "Crédit Moyen FR", *all_results['CMF'])
                if cmf_jours_dataframes is not None:
                    add_df_to_sheet(wb, "Crédit Moyen FR", ("Evolution du Crédit moyen fournisseur en jours", cmf_jours_dataframes),
                                start_column=8, plot=True, remove_total=True)

            if config['sheets']['Crédit Moyen FR'] and cmf_top_5_by_year_dataframes and config['sheets']['TOP FR']:
                add_df_to_sheet(wb, "Crédit Moyen FR TOP 5", *cmf_top_5_by_year_dataframes[-1])
                if len(cmf_top_5_by_year_dataframes) > 1:
                    add_df_to_sheet(wb, "Crédit Moyen FR TOP 5", *cmf_top_5_by_year_dataframes[-2],start_column=8)
                start_column = 16 if len(cmf_top_5_by_year_dataframes) > 1 else 8
                add_df_to_sheet(wb, "Crédit Moyen FR TOP 5", ("Evolution du Crédit moyen fournisseur pour le top 5 des fournisseurs en jours",next(reversed(cmf_top_5_jours_dataframes.values()))),start_column= start_column,plot = True)
              
            if config['sheets']['TOP FR'] and all_results['Top_FR']:
                add_df_to_sheet(wb, "TOP FR", *all_results['Top_FR'], plot=True, x_axis=compte_lib_column_fournisseur, 
                            chart_type="pie", percentage_col=percentage_column_fournisseur, remove_total=True)
            
            # Supprimer la feuille par défaut
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']

            # Sauvegarder le fichier Excel
            while True:
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Fichiers Excel", "*.xlsx")],
                    title="Enregistrer le fichier Excel"
                )

                if not output_path:  # User canceled the save dialog
                    self.status_label.config(text="Opération annulée")
                    break

                try:

                    wb.save(output_path)
                    messagebox.showinfo(
                        "Succès",
                        f"Le traitement est terminé.\nLe fichier a été enregistré : {output_path}"
                    )
                    self.status_label.config(text="Traitement terminé avec succès")
                    self.output_path = output_path

                    if not reporting:
                        self.open_excel_file(output_path)
                    break

                except PermissionError as e:
                    self.handle_error(e, "Erreur de permission")
                except FileNotFoundError as e:
                    self.handle_error(e, "Fichier non trouvé")
                except IOError as e:
                    self.handle_error(e, "Erreur I/O")
                except Exception as e:
                    self.handle_error(e, "Erreur générale")

            if reporting:
                reporting_window = self.report_generator_config(output_path)
                self.wait_window(reporting_window)
                self.open_excel_file(output_path)

        except Exception as e:
            self.handle_error(e, "Erreur générale")

        finally:
            # Réactiver les boutons après le traitement
            self.select_files_button.config(state='normal')
            self.process_button.config(state='normal' if self.fec_files else 'disabled')
            self.delete_button.config(state='normal')
            self.delete_all_button.config(state='normal')

# Exécution de l'application
if __name__ == "__main__":
    app = FECProcessor()
    app.mainloop()
      