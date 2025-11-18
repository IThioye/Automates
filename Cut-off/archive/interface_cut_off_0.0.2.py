import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import pandas as pd
import warnings
import numpy as np
from typing import Literal
from unidecode import unidecode
import re
import traceback
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side,Alignment
from openpyxl import Workbook

warnings.filterwarnings("ignore")

def reorder_months(df):
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

def custom_round(value):
    # If the decimal part is greater than or equal to 0.5, round up
    if value - np.floor(value) >= 0.5:
        return int(np.ceil(value))  # Round up and convert to int
    else:
        return int(np.floor(value))  # Round down and convert to int

def adjust_column_width(ws,min_row=2,min_col=2,start = 2):
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
    for col_idx, column_cells in enumerate(ws.iter_cols(min_row=min_row, min_col=min_col), start=start):
        column_empty = True
        column_max_width = 10  # Default minimum width
        
        row_index = start  # Start at second row
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

def design_title_cell(ws, title, sheet_name,start_column,end_column):
    """
    Colore la cellule de titre en fonction du nom de la feuille et ajuste sa largeur de colonne.
    
    Args:
        ws: Feuille Excel
        title: Contenu de la cellule à trouver
        sheet_name: Nom de la feuille pour déterminer la couleur
    """
    # Définir les couleurs pour différentes feuilles
    title_colors = {
        "Cut-Off Ventes": "BDD7EE",  # Blue
        "Cut-Off Avoirs de vente": "FFE699",  # Orange
        "Cut-Off Achats": "E5E5FF", # Purple
        "Cut-Off Avoirs d'achat": "C1F0C8",  # Green
    }

    # Définition du style de bordure noire fine
    black_border = Border(
        left=Side(style='medium', color='000000'),
        right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'),
        bottom=Side(style='medium', color='000000')
    )

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
                    cell.border = black_border
                    ws.merge_cells(start_row=cell.row, start_column=start_column, end_row=cell.row, end_column=end_column)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    return

def color_between_values(ws,first_value,last_value,fill_color,include_start_col_and_end_col=True):
    """
    Colore les cellules entre deux colonnes identifiées par leurs noms d'en-tête.
    
    Args:
        ws: La feuille de calcul
        sheet_name: Le nom de la feuille de calcul
        df_width: La largeur du DataFrame
        start_column: La colonne de départ du DataFrame
    """

    value_to_find = first_value
        
    # Trouver toutes les lignes d'en-tête
    rows_to_colour = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
        for cell in row:
            if cell.value == value_to_find:
                rows_to_colour.append(row_idx)

    # Pour chaque ligne d'en-tête trouvée
    for row_idx in rows_to_colour:        
        first_col = None
        second_col = None

        # Obtenir la ligne contenant l'en-tête
        header_row = ws[row_idx]
        
        # Trouver les colonnes dans cette table spécifique
        for cell in header_row:
            cell_value = str(cell.value).strip() if cell.value else ""
            if cell_value == first_value:
                first_col = cell.column
            if cell_value == last_value:
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

def design_cells(ws, type, sheet_name):
    """
    Design cells in an Excel worksheet based on the type and sheet name.

    Applies a bold font with black text to cells containing "Total sélectionné", 
    "% [type] sélectionné(e)s", "Nombre de factures totales", and "Nombre de factures sélectionnées".
    Also applies a bold font with black text to cells containing "Total [type]".

    Args:
        ws: Excel worksheet object.
        type: Type of cut-off (e.g. "Ventes" or "Achats").
        sheet_name: Name of the Excel sheet to design the cells in.

    """

    # Définir les couleurs pour différentes feuilles
    colors = {
        "Cut-Off Ventes": "BDD7EE",  # Blue
        "Cut-Off Avoirs de vente": "FFE699",  # Orange
        "Cut-Off Achats": "E5E5FF", # Purple
        "Cut-Off Avoirs d'achat": "C1F0C8",  # Green
    }

    values_to_color = ["Total sélectionné",f"% {type} sélectionné(e)s","Nombre de factures totales","Nombre de factures sélectionnées"]
    
    # Trouver la cellule de titre
    if sheet_name in list(colors.keys()):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value in values_to_color:
                    # Obtenir la couleur pour la feuille actuelle (par défaut à blanc si la feuille n'est pas dans le dictionnaire)
                    color = colors.get(sheet_name, "FFFFFF")                    
                    # Appliquer la couleur
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    cell.font = Font(bold=True, color="000000")  # Texte noir en gras

                elif cell.value is not None and f"Total {type}" in str(cell.value):
                    cell.font = Font(bold=True, color="000000")

def add_empty_df_to_sheet(wb, sheet_name, df, titles, error=None, start_column = 2, start_row = 2, empty_rows_to_add = 0):
    """
    Adds an empty DataFrame as an Excel table to an Excel sheet with a title and optional error message.
    Clears existing tables to avoid conflicts.

    Args:
        wb: Excel workbook object.
        sheet_name: Name of the Excel sheet to add the DataFrame to.
        df: DataFrame to add to the sheet, only headers are used since it's empty.
        title: Title of the DataFrame, added as a header in the sheet.
        error: Optional error message to display in the sheet, defaults to None.
        start_column: Column index to start writing the DataFrame at, defaults to 2.
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

    current_row = start_row

    titre_wb = titles["titre_wb"]
    titre_selection_cut_off = titles["titre_selection_cut_off"] 
    titre_document_apportes = titles["titre_document_apportes"]

    if error is not None:
        # Add and color wb title
        ws.cell(row=current_row, column=start_column, value=titre_wb)
        design_title_cell(ws, 
                      titre_wb, 
                      sheet_name,
                      start_column,
                      end_column = start_column + len(df.columns) + 9
                      )
        current_row += 2

        # Add and color selection title
        ws.cell(row=current_row, column=start_column, value=titre_selection_cut_off)
        design_title_cell(ws, 
                      titre_selection_cut_off, 
                      sheet_name,
                      start_column,
                      end_column = start_column + len(df.columns) - 1
                      )
        current_row += 2
    
    else : 
        ws.cell(row=current_row, column=start_column, value = titre_document_apportes)
        design_title_cell(ws, 
                          titre_document_apportes, 
                          sheet_name,
                          start_column,
                          end_column = start_column + len(df.columns)-1
                          )
        current_row += 2

    black_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # Write DataFrame headers
    for col_idx, column in enumerate(df.columns, start=start_column):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = str(column)  # Convert to string to avoid type issues
        cell.font = Font(bold=True)
        cell.border = black_border
   
    if error is not None:
        current_row += 2
        cell = ws.cell(row=current_row, column=start_column, value=error)
        cell.font = Font(color="FF0000",bold=True)
    else:        
        # Insert empty rows with border if needed
        if empty_rows_to_add > 0 :
            current_row += 1  # Move below header space
            for _ in range(empty_rows_to_add):
                for col_idx in range(start_column, start_column + len(df.columns)):
                    cell = ws.cell(row=current_row, column=col_idx)
                    cell.value = ""  # Set as visually empty
                    cell.border = black_border
                current_row += 1  # Move to next line
        
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        color_between_values(ws,
                        first_value=list(df)[0],
                        last_value=list(df)[-1],
                        fill_color=gray_fill,
                        include_start_col_and_end_col=True
                            )   
        
    # Adjust column widths at the end
    try:
        adjust_column_width(ws)
    except Exception as e:
        print(f"Warning: Could not adjust column widths: {str(e)}")
    
def add_df_to_sheet(wb, sheet_name,df ,titles,type, start_column=2, start_row = 2):

    """
    Adds a DataFrame as an Excel table to an Excel sheet.
    Includes error handling and safeguards for file opening.
    
    Args:
        wb: Excel workbook
        sheet_name: Name of the Excel sheet to add the DataFrame to
        df: DataFrame to add to the sheet
        title: Title of the DataFrame, used for the chart title and table name
        type: Type of the cut-off to color the cells
        start_column: Column index to start writing the DataFrame at
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

    current_row = start_row
    used_table_names = set()  # Track used table names

    titre_wb = titles["titre_wb"]
    titre_selection_cut_off = titles["titre_selection_cut_off"]

    # Add and color wb title
    ws.cell(row=current_row, column=start_column, value=titre_wb)
    design_title_cell(ws, 
                      titre_wb, 
                      sheet_name,
                      start_column,
                      end_column=start_column+len(df.columns)+9
                      )
    current_row += 2

    # Add and color selection title
    ws.cell(row=current_row, column=start_column, value=titre_selection_cut_off)
    design_title_cell(ws, 
                      titre_selection_cut_off, 
                      sheet_name,
                      start_column,
                      end_column=start_column+len(df.columns)-1
                      )
    current_row += 2
    df_start_row = current_row

    # Write DataFrame headers
    for col_idx, column in enumerate(df.columns, start=start_column):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = str(column)  # Convert to string to avoid type issues
        cell.font = Font(bold=True)
        
    # Write DataFrame data
    for row_idx, row in enumerate(df.values, start=current_row + 1):
        for col_idx, value in enumerate(row, start=start_column):
            cell = ws.cell(row=row_idx, column=col_idx)
            # Handle different types of zero values and None
            if value in (0, "0", 0.0, None, ""):
                cell.value = ""
            else:
                cell.value = value
    
    # Color special cells
    design_cells(ws, type, sheet_name)

    # Create unique table name
    base_table_name = re.sub(r'[^a-zA-Z0-9_]', '_', unidecode(titre_selection_cut_off))  # Remove special characters
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
        print(f"Warning: Could not create table for {titre_selection_cut_off}: {str(e)}")
    
        
    # Adjust column widths at the end
    try:
        adjust_column_width(ws,min_row=df_start_row,min_col=start_column,start = df_start_row)
    except Exception as e:
        print(f"Warning: Could not adjust column widths: {str(e)}")

def cut_off(df,type:Literal["achats","ventes","avoirs d'achat","avoirs de vente"],intervalle_periode=None,compte_num_column = "CompteNum",mois_en_ordre=None,num_months = 2,top_max=20,seuil=3000):

    """
    Select and calculate cut-off for given type of transactions.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing the transactions.
    type : Literal["achats","ventes","avoirs d'achat","avoirs de vente"]
        The type of transactions to select and calculate cut-off for.
    periode : str, optional
        The period for which to select and calculate cut-off. If None, defaults to the last two months of the year.
    compte_num_column : str, optional
        The column containing the account numbers. Defaults to "CompteNum".
    mois_en_ordre : list, optional
        The list of months in order. Defaults to None.
    top_max : int, optional
        The maximum number of rows to select for the top selection. Defaults to 20.

    Returns
    -------
    tuple
        A tuple containing the title of the selection, the DataFrame of the selection, and an error message if there are no data for the given period.
    """
    
    compte_num,liste_journal_code = ('401',["AC","HA"]) if type in ["achats","avoirs d'achat"] else ('411',["VE","VT","VTE"])
    colonne_de_calcul = "Credit" if type in ["achats","avoirs de vente"] else "Debit"
    
    periode = None
    if mois_en_ordre is not None:
        if intervalle_periode is not None:
            from_,to_ = intervalle_periode
            periode = mois_en_ordre[from_:to_+1]
        else:
            periode = mois_en_ordre[:num_months] if type in ["achats","avoirs de vente"] else mois_en_ordre[-num_months:]

    data = df[(df["JournalCode"].str.strip().isin(liste_journal_code)) & (df[compte_num_column].astype(str).str.startswith(compte_num))].loc[df['Mois'].isin((periode))]
    periode_annee = f"{'-'.join([mois for mois in periode])} {df['Année'].unique()[-1]}"

    titles =   {"titre_wb" : f'Cut-off des {type} à la clôture ({periode_annee.split("-")[-1].split(" ")[0]})',
                "titre_selection_cut_off": f'Sélection cut-off des {type} sur la période {periode_annee}',
                "titre_document_apportes": 'Documents apportés par le client'
                }
    
    empty_df = pd.DataFrame(columns=["Date d'écriture comptable", 'Compte', 'Libellé', 'Debit', 'Credit'])
    error = f"Il n'y a pas de données pour les {type} durant la période {periode_annee}"    

    if data.shape[0] != 0 : 
        classement = data[["JournalCode",'EcritureDate',compte_num_column,'PieceRef','EcritureLib',colonne_de_calcul]].sort_values(by=colonne_de_calcul,ascending=False)
        
        if int(classement[colonne_de_calcul].unique()[0]) == 0:
                return titles,empty_df,error
        
        classement[colonne_de_calcul] = classement[colonne_de_calcul].apply(custom_round)
        classement=classement[~(classement[colonne_de_calcul] == 0)]
        classement.columns = ["Code","Date d'écriture comptable", 'Compte', 'PieceRef','Libellé', colonne_de_calcul]
        classement["Date d'écriture comptable"] = classement["Date d'écriture comptable"].dt.strftime('%d/%m/%Y')

        periode_annee = f"{'-'.join([mois for mois in periode])} {data['Année'].unique()[0]}"

        top = classement.head(top_max)
        top = top[top[colonne_de_calcul]>=seuil] if type in ["avoirs d'achat","avoirs de vente"] else top
        if top.shape[0] == 0:
            error = f"Il n'y a pas d'{type} durant la période {periode_annee} au dessus de {seuil}€"
            return titles,empty_df,error
        
        top["Sélection"] = top[colonne_de_calcul]
        total_colonne_calcul = classement[colonne_de_calcul].sum()
        total_selection = top["Sélection"].sum()

        pourcentage_de_selectionnés = f"{round((total_selection/total_colonne_calcul)*100)} %" if total_colonne_calcul not in [None,0,0.0] and total_selection not in [None,0,0.0] else None
        bottom = pd.DataFrame([
                                {   "Code":'',
                                    "Date d'écriture comptable" : "",
                                    "Compte" : "",
                                    "PieceRef":'',
                                    "Libellé" : f"Total {type} {periode_annee} (hors avoirs)" if type in ["achats","ventes"] else f"Total {type} {periode_annee}",
                                    colonne_de_calcul: total_colonne_calcul,
                                    "Sélection" : total_selection
                                },
                                {   "Code":'',
                                    "Date d'écriture comptable" : "",
                                    "Compte" : "",
                                    "PieceRef":'',
                                    "Libellé" : "",
                                    colonne_de_calcul : "Total sélectionné",
                                    "Sélection" : total_selection
                                },
                                {   "Code":'',
                                    "Date d'écriture comptable" : "",
                                    "Compte" : "",
                                    "PieceRef":'',
                                    "Libellé" : "",
                                    colonne_de_calcul: f"% {type} sélectionné(e)s",
                                    "Sélection" : pourcentage_de_selectionnés
                                },
                                {   "Code":'',
                                    "Date d'écriture comptable" : "",
                                    "Compte" : "",
                                    "PieceRef":'',
                                    "Libellé" : "",
                                    colonne_de_calcul : "Nombre de factures totales",
                                    "Sélection" : classement.shape[0]
                                },
                                {   "Code":'',
                                    "Date d'écriture comptable" : "",
                                    "Compte" : "",
                                    "PieceRef":'',
                                    "Libellé" : "",
                                    colonne_de_calcul : "Nombre de factures sélectionnées",
                                    "Sélection" : len(top)
                                }
                            ])

        top_5_final = pd.concat([top,bottom],axis=0)
        top_5_final = top_5_final.reset_index(drop=True)

        titles =   {"titre_wb" : f'Cut-off des {type} à la clôture ({periode_annee.split("-")[-1].split(" ")[0]})',
                    "titre_selection_cut_off": f'Sélection cut-off des {type} sur la période {periode_annee}',
                    "titre_document_apportes": 'Documents apportés par le client'
                    }
        return titles,top_5_final,None

    return titles,empty_df,error

class CutOffApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Cut-Off")
        self.geometry("800x650")
        self.minsize(800, 650)
        self.configure(bg="#f0f2f5")
        self.center_main_window()
        
        # Set theme to use system theme if available
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        
        # Configure styles
        style.configure("Treeview", 
                        background="#ffffff",
                        foreground="#333333", 
                        rowheight=30,
                        fieldbackground="#ffffff",
                        borderwidth=0,
                        font=('Segoe UI', 10))
        
        style.configure("Treeview.Heading", 
                        font=('Segoe UI', 11, 'bold'),
                        background="#e1e5ea",
                        foreground="#333333",
                        padding=5)
        
        style.map('Treeview', 
                  background=[('selected', '#4a6da7')],
                  foreground=[('selected', '#ffffff')])
                  
        style.configure("TButton", 
                        padding=10, 
                        font=('Segoe UI', 11),
                        background="#4a6da7",
                        foreground="#ffffff")
                        
        style.map("TButton",
                  background=[('active', '#3a5d97')],
                  relief=[('pressed', 'groove'), ('!pressed', 'ridge')])
                  
        style.configure("TCombobox", 
                        padding=8,
                        font=('Segoe UI', 11))
                        
        style.configure("TFrame", background="#f0f2f5")
        
        # Create main containers
        self.main_frame = ttk.Frame(self, padding="20", style="TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title with emoji
        self.title_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.title_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.title_label = ttk.Label(
            self.title_frame, 
            text="📁 Cut-Off", 
            font=('Segoe UI', 16, 'bold'),
            background="#f0f2f5",
            foreground="#333333"
        )
        self.title_label.pack(side=tk.LEFT)
        
        # Month input
        self.month_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.month_frame.pack(fill=tk.X, pady=(0, 15))

        self.month_label = ttk.Label(
            self.month_frame, 
            text="Nombre de mois à considérer avant/après la clôture:", 
            font=('Segoe UI', 11),
            background="#f0f2f5",
            foreground="#333333"
        )
        self.month_label.pack(side=tk.LEFT, padx=(0, 10))

        self.month_var = tk.StringVar(value=2)  # Default value

        self.month_entry = ttk.Entry(
            self.month_frame, 
            textvariable=self.month_var,
            width=20
        )
        self.month_entry.pack(side=tk.LEFT)

        self.info_icon = tk.Label(
            self.month_frame,
            text="❓",
            font=("Segoe UI", 9),
            foreground="#0078D7",  # Microsoft blue
            background="#f0f2f5",
            cursor="arrow",            
            bd=0,  # no border
            highlightthickness=0,  # no highlight border
            padx=0,
            pady=0
        )
        self.info_icon.pack(side=tk.LEFT, padx=(5, 0))

        # Bind hover events
        self.info_icon.bind("<Enter>", self.show_info_popup)
        self.info_icon.bind("<Leave>", self.close_info_popup)

        # Initialize popup as None
        self.info_popup = None

        # Define the callback function
        def toggle_month_entry():
            if self.checkbox_var.get():
                self.month_entry.config(state='disabled')
            else:
                self.month_entry.config(state='normal')

        # Create the BooleanVar to hold the state of the checkbox
        self.checkbox_var = tk.BooleanVar(value=False)  # Default is unchecked

        # Add the checkbox to the right of the frame
        self.checkbox = ttk.Checkbutton(
            self.month_frame,
            text="Choisir les périodes manuellement",
            variable=self.checkbox_var,
            command=toggle_month_entry
        )
        self.checkbox.pack(side=tk.RIGHT, padx=(10, 0))

        # note line
        self.note_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.note_frame.pack(fill=tk.X, pady=(0, 15))

        self.note_label = ttk.Label(
            self.note_frame, 
            text="La version actuelle nécessite deux FECS pour le résultat final : le FEC de l'année N-1 pour la clôture \net le FEC de l'année N pour la post-clôture.", 
            font=('Segoe UI', 11, 'italic'),
            background="#f0f2f5",
            foreground="#333333"
        )
        self.note_label.pack(side=tk.LEFT, padx=(0, 10))

        # File controls
        self.file_controls_frame = ttk.Frame(self.main_frame, style="TFrame")
        self.file_controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.browse_button = ttk.Button(
            self.file_controls_frame,
            text="🔍 Sélectionner fichiers",
            command=self.browse_files
        )
        self.browse_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(
            self.file_controls_frame,
            text="🗑️ Tout supprimer",
            command=self.clear_files
        )
        self.clear_button.pack(side=tk.LEFT)
        
        # File treeview with scrollbar in a container
        self.tree_container = ttk.Frame(self.main_frame, style="TFrame")
        self.tree_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Add vertical scrollbar
        self.tree_scrollbar_y = ttk.Scrollbar(self.tree_container)
        self.tree_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create the treeview
        self.file_tree = ttk.Treeview(
            self.tree_container,
            columns=("Année", "Fichier"),
            show="headings",
            yscrollcommand=self.tree_scrollbar_y.set
        )
        
        # Configure scrollbars
        self.tree_scrollbar_y.config(command=self.file_tree.yview)
        
        # Configure columns
        self.file_tree.heading("Année", text="Année")
        self.file_tree.heading("Fichier", text="Fichier")
        
        # Set column widths
        self.file_tree.column("Année", width=100, anchor=tk.CENTER)
        self.file_tree.column("Fichier", width=300)
        
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        
        # Bind double-click to edit year
        self.file_tree.bind("<Double-1>", self.edit_year)
        
        # Process button at the bottom
        self.process_button = ttk.Button(
            self.main_frame,
            text="✅ Générer le cut-off",
            command=self.process_files
        )
        self.process_button.pack(side=tk.RIGHT, pady=(0, 10))
        
        self.year_editor = None
        self.files_data = []
        self.current_year = str(datetime.now().year)
        self.output_path = ''

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

    def show_info_popup(self, event):
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

    def browse_files(self):
        """Open file dialog to select multiple files"""
        filepaths = filedialog.askopenfilenames(
                title="Sélectionner les fichiers FEC",
                filetypes=[("Fichiers texte", "*.txt")]
            )
        
        if not filepaths:
            return
            
        for filepath in filepaths:
            filename = os.path.basename(filepath)
            default_year = str(filename.split("FEC")[-1][:4]) if filename.split("FEC")[-1][:3] in [f"20{i}" for i in range(1,4)] else self.current_year
            # Add to our treeview with default year as current year
            item_id = self.file_tree.insert("", tk.END, values=(default_year, filename,))
            self.files_data.append({
                "id": item_id,
                "year":default_year,
                "filename": filename,
                "filepath": filepath
            })
        
    def clear_files(self):
        """Clear all files from the treeview"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self.files_data = []

    def edit_year(self, event):
        """Handle double-click to edit year"""
        # Get the item that was clicked
        item_id = self.file_tree.identify_row(event.y)
        if not item_id:
            return
            
        # Get the column that was clicked
        column = self.file_tree.identify_column(event.x)
        
        # Only allow editing the Year column (column #1)
        if column != "#1":
            return
            
        # Get the bbox of the cell
        x, y, width, height = self.file_tree.bbox(item_id, column)
        
        # Create an Entry widget for editing
        if self.year_editor:
            self.year_editor.destroy()
            
        current_value = self.file_tree.item(item_id)['values'][0]
        
        self.year_editor = ttk.Entry(self.file_tree, width=10)
        self.year_editor.place(x=x, y=y, width=width, height=height)
        self.year_editor.insert(0, current_value)
        self.year_editor.select_range(0, tk.END)
        self.year_editor.focus()
        
        # Bind events for when editing is complete
        self.year_editor.bind("<Return>", lambda e, item=item_id: self.save_year(item))
        self.year_editor.bind("<FocusOut>", lambda e, item=item_id: self.save_year(item))

    def save_year(self, item_id):
        """Save the edited year value"""
        if not self.year_editor:
            return
            
        try:
            new_year = self.year_editor.get()
            # Basic validation: year should be reasonable
            if 1900 <= int(new_year) <= 2100 :
                # Get the current values
                values = list(self.file_tree.item(item_id)['values'])
                # Update the year
                values[0] = new_year
                # Update the treeview
                self.file_tree.item(item_id, values=values)
                
                # Update our stored data
                for item in self.files_data:
                    if item["id"] == item_id:
                        item["year"] = new_year
                        break
        except ValueError:
            try:
                if 1900 <= int(new_year.split("-")[0]) <= 2100 and 1900 <= int(new_year.split("-")[-1]) <= 2100:
                    # Get the current values
                    values = list(self.file_tree.item(item_id)['values'])
                    # Update the year
                    values[0] = new_year
                    # Update the treeview
                    self.file_tree.item(item_id, values=values)
                    
                    # Update our stored data
                    for item in self.files_data:
                        if item["id"] == item_id:
                            item["year"] = new_year
                            break
                else:
                    self.handle_error("Veuillez entrer une année entre 1900 et 2100")
            except ValueError:
                self.handle_error("Veuillez enter une année valide")
            
        self.year_editor.destroy()
        self.year_editor = None

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

    def handle_error(self, error, error_type="Erreur"):
        """Centralized error handling"""
        error_message = str(error)
        self.error_details = f"{error_type}:\n\n{error_message}\n\nTraceback:\n{traceback.format_exc()}"
        
        # Show error message
        messagebox.showerror(error_type, error_message)
            
        # Log error
        print(f"{error_type}: {error_message}")
        traceback.print_exc()

    def period_choice(self):
        config_window = tk.Toplevel(self)
        config_window.title("Choix des périodes pour l'achat et la vente")
        config_window.geometry("800x600")
        config_window.transient(self)
        config_window.grab_set()
        self.center_config_window(config_window)
        
        # Main frame with scrollbar
        main_frame = tk.Frame(config_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Add a canvas with scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        
        content_frame = tk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        # Configure scroll region when frame size changes
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        content_frame.bind("<Configure>", configure_scroll_region)
        
        # Ensure canvas resizes with window
        def configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind("<Configure>", configure_canvas)
        
        # Store all dropdown variables to access later
        self.period_selections = {}
        first_year = min(self.files_data, key=lambda x: x["year"])["year"]
        
        # Create frame for each file
        for i, file_data in enumerate(self.files_data):
            year = file_data["year"]
            file_path = file_data["filepath"]
            file_name = os.path.basename(file_path)
            liste_encoding = ["utf-8","MacRoman"]

            try: 
                df = pd.read_csv(file_path, sep='\t', encoding= liste_encoding[0], engine='python')
                if df.shape[1] == 1:
                    df = pd.read_csv(file_path, sep='|', encoding= liste_encoding[0], engine='python')
            except UnicodeDecodeError:
                    df = pd.read_csv(file_path, sep='\t', encoding= liste_encoding[1], engine='python')
                    if df.shape[1] == 1:
                        df = pd.read_csv(file_path, sep='|', encoding= liste_encoding[1], engine='python')

            df["EcritureDate"] = pd.to_datetime(df["EcritureDate"],format='%Y%m%d', errors='coerce')
            df['Année'] = df["EcritureDate"].dt.year
            df['Mois'] = df["EcritureDate"].dt.month_name(locale='French')
            
            months = reorder_months(df)
            file_frame = tk.LabelFrame(content_frame, text=f"Fichier: {file_name} ({year})", padx=10, pady=10)
            file_frame.pack(fill=tk.X, expand=True, pady=10)
            
            def create_callback(year, is_first_year):
                def callback(*args):
                    # Find the corresponding file_data entry
                    for file_data in self.files_data:
                        if file_data["year"] == year:
                            if is_first_year:
                                # For first year, update ventes period
                                ventes_from_val = self.period_selections[year]["ventes_from"].get()
                                ventes_to_val = self.period_selections[year]["ventes_to"].get()
                                ventes_from_idx = months.index(ventes_from_val)
                                ventes_to_idx = months.index(ventes_to_val)
                                file_data["periode"] = (ventes_from_idx, ventes_to_idx)
                                print(f"Updated ventes for year {year}: {ventes_from_val} to {ventes_to_val}")
                            else:
                                # For second year, update achats period
                                achats_from_val = self.period_selections[year]["achats_from"].get()
                                achats_to_val = self.period_selections[year]["achats_to"].get()
                                achats_from_idx = months.index(achats_from_val)
                                achats_to_idx = months.index(achats_to_val)
                                file_data["periode"] = (achats_from_idx, achats_to_idx)
                                print(f"Updated achats for year {year}: {achats_from_val} to {achats_to_val}")
                            break
                return callback
                        
            if year == first_year :            
                # First year - Ventes section
                ventes_frame = tk.Frame(file_frame)
                ventes_frame.pack(fill=tk.X, pady=5)
                
                ventes_label = tk.Label(ventes_frame, text=f"{year} - Ventes (& avoir d'achats) :")
                ventes_label.grid(row=0, column=0, sticky="w", padx=5)
                
                # From dropdown
                ventes_from_var = tk.StringVar(value=months[0])
                ventes_from_label = tk.Label(ventes_frame, text="de")
                ventes_from_label.grid(row=0, column=1, padx=5)
                ventes_from_dropdown = tk.OptionMenu(ventes_frame, ventes_from_var, *months)
                ventes_from_dropdown.grid(row=0, column=2, padx=5)
                
                # To dropdown
                ventes_to_var = tk.StringVar(value=months[-1])
                ventes_to_label = tk.Label(ventes_frame, text="à")
                ventes_to_label.grid(row=0, column=3, padx=5)
                ventes_to_dropdown = tk.OptionMenu(ventes_frame, ventes_to_var, *months)
                ventes_to_dropdown.grid(row=0, column=4, padx=5)

                self.period_selections[year] = {
                    "ventes_from": ventes_from_var,
                    "ventes_to": ventes_to_var
                }
                
                # Add trace to update values when dropdowns change
                ventes_from_var.trace_add("write", create_callback(year,True))
                ventes_to_var.trace_add("write", create_callback(year,True))

            else:
                # Second year - Achats section
                achats_frame = tk.Frame(file_frame)
                achats_frame.pack(fill=tk.X, pady=5)
                
                achats_label = tk.Label(achats_frame, text=f"{year} - Achats (& avoir de ventes) :")
                achats_label.grid(row=0, column=0, sticky="w", padx=5)
                
                # From dropdown
                achats_from_var = tk.StringVar(value=months[0])
                achats_from_label = tk.Label(achats_frame, text="de")
                achats_from_label.grid(row=0, column=1, padx=5)
                achats_from_dropdown = tk.OptionMenu(achats_frame, achats_from_var, *months)
                achats_from_dropdown.grid(row=0, column=2, padx=5)
                
                # To dropdown
                achats_to_var = tk.StringVar(value=months[-1])
                achats_to_label = tk.Label(achats_frame, text="à")
                achats_to_label.grid(row=0, column=3, padx=5)
                achats_to_dropdown = tk.OptionMenu(achats_frame, achats_to_var, *months)
                achats_to_dropdown.grid(row=0, column=4, padx=5)    
                
                # Store selections for this file
                self.period_selections[year] = {
                    "achats_from": achats_from_var,
                    "achats_to": achats_to_var
                }
                
                achats_from_var.trace_add("write", create_callback(year,False))
                achats_to_var.trace_add("write", create_callback(year,False))

        # Validate button   
        def validate_periods(window):
            """Process the period selections and close the window"""            
            # Close the config window
            window.destroy()

        # Validate button frame
        button_frame = tk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        validate_button = tk.Button(button_frame, text="Valider", 
                                command=lambda: validate_periods(config_window))
        validate_button.pack(side=tk.RIGHT, padx=10)
        
        cancel_button = tk.Button(button_frame, text="Annuler", 
                                command=config_window.destroy)
        cancel_button.pack(side=tk.RIGHT, padx=10)

        config_window.wait_window()

        
        # # You might want to update the main UI or continue processing here
        # self.update_status(f"Périodes sélectionnées pour {len(self.period_selections)} fichiers")      

    def process_files(self):
        """Process the files with the selected month and years"""
        if not self.files_data:
            self.handle_error("Aucun fichier sélectionné")
            return
        
        num_month_for_period = int(self.month_var.get()) if not self.checkbox_var.get() else None
        if num_month_for_period is not None : 
            if num_month_for_period < 1 or num_month_for_period > 12:
                self.handle_error("Veuillez choisir un mois entre 1 et 12 !")
                return
        
        first_year = min(self.files_data, key=lambda x: x["year"])["year"]

        if self.checkbox_var.get():
            self.period_choice()

        wb = Workbook()
        try:
            for file_data in self.files_data:
                year = file_data["year"]
                file_path = file_data["filepath"]
                intervalle_periode = file_data["periode"] if self.checkbox_var.get() else None

                liste_encoding = ["utf-8","MacRoman"]
                try: 
                    df = pd.read_csv(file_path, sep='\t', encoding= liste_encoding[0], engine='python')
                    if df.shape[1] == 1:
                        df = pd.read_csv(file_path, sep='|', encoding= liste_encoding[0], engine='python')
                except UnicodeDecodeError:
                        df = pd.read_csv(file_path, sep='\t', encoding= liste_encoding[1], engine='python')
                        if df.shape[1] == 1:
                            df = pd.read_csv(file_path, sep='|', encoding= liste_encoding[1], engine='python')

                df["EcritureDate"] = pd.to_datetime(df["EcritureDate"],format='%Y%m%d', errors='coerce')
                df['Année'] = df["EcritureDate"].dt.year
                df['Mois'] = df["EcritureDate"].dt.month_name(locale='French')

                # Replace the original mois_en_ordre assignment with:
                mois_en_ordre = reorder_months(df)
                df["Credit"] = pd.to_numeric(df["Credit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0)
                df["Debit"] = pd.to_numeric(df["Debit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0)
                
                type_list = ["ventes","avoirs d'achat"] if year == first_year else ["achats","avoirs de vente"]

                for type in type_list:
                    titles,selection_cut_off,error = cut_off(df,
                                                             type= type,
                                                             mois_en_ordre= mois_en_ordre,
                                                             intervalle_periode=intervalle_periode,
                                                             num_months=num_month_for_period
                                                             )
                    if error is None:
                        add_df_to_sheet(wb,
                                        sheet_name = f'Cut-Off {type.capitalize()}',
                                        df = selection_cut_off,
                                        titles = titles,
                                        type = type
                                        )
                        
                        document_apportes = pd.DataFrame(columns=["Commentaire ATC", "Date facture", "Montant facture", "Date du BL", "BL signé ?",	"Date d'entrée en stock", "Ecart date",	"Commentaire"])
                        add_empty_df_to_sheet(wb,
                                        sheet_name = f'Cut-Off {type.capitalize()}',
                                        df = document_apportes,
                                        titles = titles,
                                        start_column= len(list(selection_cut_off)) + 4,
                                        start_row= 4,
                                        empty_rows_to_add = len(selection_cut_off) - 5
                                        )  
                                          
                    else:
                        add_empty_df_to_sheet(wb,
                                        sheet_name = f'Cut-Off {type.capitalize()}',
                                        df = selection_cut_off,
                                        titles = titles,
                                        error = error
                                        )

            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']        

            while True:
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Fichiers Excel", "*.xlsx")],
                    title="Enregistrer le fichier Excel"
                )

                if not output_path:  # User canceled the save dialog
                    break

                try:
                    wb.save(output_path)
                    messagebox.showinfo(
                        "Succès",
                        f"Le traitement est terminé.\nLe fichier a été enregistré : {output_path}"
                    )
                    
                    self.output_path = output_path
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

        except Exception as e:
            self.handle_error(e, "Erreur générale")  

if __name__ == "__main__":
    app = CutOffApp()
    app.mainloop()