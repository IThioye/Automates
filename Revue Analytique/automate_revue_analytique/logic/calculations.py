import pandas as pd
import numpy as np 

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
    
    mask = (ca_mensuel_cumule[colonne_annee] == total_cumule)
    # Assign an occurrence number for each True in the mask
    occurrence = mask.cumsum()
    # Set every even occurrence to 0
    ca_mensuel_cumule.loc[mask & (occurrence > 1), colonne_annee] = 0
    
    if len(df["Mois"].unique()) == 12:
        ca_mensuel_cumule[f"% {colonne_annee}"] = ca_mensuel_cumule.apply(lambda row: f'{round((row[colonne_annee]/total_cumule)*100)}%',axis = 1)
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
