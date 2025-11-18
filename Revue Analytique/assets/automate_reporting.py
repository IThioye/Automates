import tkinter as tk
from tkinter import ttk, messagebox,filedialog
import os
import subprocess
import threading
from pathlib import Path

from docx import Document
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import plotly.io as pio
import plotly.express as px
from io import BytesIO
from docx.shared import Inches
import numpy as np
from PIL import Image
import warnings
warnings.filterwarnings("ignore")

config_elements_fixes = {  
            
            "Evolution_du_CA_mensuel": {"sheet_name":"CA",
                                        "marker_text_tableau":"{CA}",
                                        "chart_type": "line",
                                        "marker_text_graphique": "{GRAPHIQUE_CA}",
                                        "remove_total": True,
                                        "section_title": "Evolution du chiffre d’affaires"
                                        },
            "Evolution_du_CA_mensuel_cumule":   {"sheet_name":"CA",
                                                 "marker_text_tableau":"{CA_CUMULE}",
                                                 "chart_type": "line",
                                                 "marker_text_graphique": "{GRAPHIQUE_CA_CUMULE}",
                                                 "remove_total": False,
                                                 "section_title": "Evolution du chiffre d’affaires"
                                                 },
            "Evolution_de_la_tresorerie_nette": {"sheet_name":"Trésorerie",
                                                 "marker_text_tableau":"{TRESORERIE_NETTE}",
                                                 "chart_type": "line",
                                                 "marker_text_graphique": "{GRAPHIQUE_TRESORERIE_NETTE}",
                                                 "remove_total": False,
                                                 "section_title": "Evolution mensuelle de la trésorerie"
                                                 },
            "Evolution_de_la_tresorerie_brute": {"sheet_name":"Trésorerie",
                                                 "marker_text_tableau":"{TRESORERIE_BRUTE}",
                                                 "chart_type": "line",
                                                 "marker_text_graphique": "{GRAPHIQUE_TRESORERIE_BRUTE}",
                                                 "remove_total": False,
                                                 "section_title": "Evolution mensuelle de la trésorerie"
                                                 },
            "Evolution_mensuelle_des_frais_de_personnel": {"sheet_name":"Salaires",
                                                       "marker_text_tableau":"{FRAIS_PERSONNEL}",
                                                       "chart_type": "line",
                                                       "marker_text_graphique": "{GRAPHIQUE_FRAIS_PERSONNEL}",
                                                       "remove_total": True,
                                                       "section_title": "Evolution des frais de personnel"
                                                       },
            "Evolution_mensuelle_des_frais_de_personnel_cumules": {"sheet_name":"Salaires",
                                                                 "marker_text_tableau":"{FRAIS_PERSONNEL_CUMULE}",
                                                                 "chart_type": "line",
                                                                 "marker_text_graphique": "{GRAPHIQUE_FRAIS_PERSONNEL_CUMULE}",
                                                                 "remove_total": False,
                                                                 "section_title": "Evolution des frais de personnel"
                                                                 },
            "Evolution_du_Credit_moyen_client_en_jours": {"sheet_name":"Crédit Moyen CLT",
                                                          "marker_text_tableau":"{CREDIT_MOYEN_CLT}",
                                                          "chart_type": "line",
                                                          "marker_text_graphique": "{GRAPHIQUE_CREDIT_MOYEN_CLT}",
                                                          "remove_total": True,
                                                          "section_title": "Crédit moyen clients & fournisseurs"
                                                          },
            "Evolution_du_Credit_moyen_fournisseur_en_jours": {"sheet_name":"Crédit Moyen FR",
                                                               "marker_text_tableau":"{CREDIT_MOYEN_FR}",
                                                               "chart_type": "line",
                                                               "marker_text_graphique": "{GRAPHIQUE_CREDIT_MOYEN_FR}",
                                                               "remove_total": True,
                                                               "section_title": "Crédit moyen clients & fournisseurs"
                                                               },
            "Evolution_des_achats_et_sous_traitance": {"sheet_name":"Achats et Sous-traitance",
                                                               "marker_text_tableau":"{ACHATS_ET_SOUS_TRAITANCE}",
                                                               "chart_type": "line",
                                                               "marker_text_graphique": "{GRAPHIQUE_ACHATS_ET_SOUS_TRAITANCE}",
                                                               "remove_total": True,
                                                               "section_title": "Evolution des achats de matières et sous-traitance"
                                                               },
            "Evolution_des_achats_et_sous_traitance_cumules": {"sheet_name":"Achats et Sous-traitance",
                                                               "marker_text_tableau":"{ACHATS_ET_SOUS_TRAITANCE_CUMULEE}",
                                                               "chart_type": "line",
                                                               "marker_text_graphique": "{GRAPHIQUE_ACHATS_ET_SOUS_TRAITANCE_CUMULEE}",
                                                               "remove_total": False,
                                                               "section_title": "Evolution des achats de matières et sous-traitance"
                                                               },
            "Total_des_investissements_par_annee": {"sheet_name":"Investissement",
                                                               "marker_text_tableau":"{TOTAL_INVESTISSEMENT}",
                                                               "chart_type": "barchart",
                                                               "marker_text_graphique": "{GRAPHIQUE_INVESTISSEMENT}",
                                                               "remove_total": True,
                                                               "section_title": "Evolution des investissements"
                                                               }
                        
                        }

config_elements_variables = {

            "TOP CLT": {
                        "chart_type": "pie",
                        "remove_total": True,
                        "marker_text":"{TOP_CLT}",
                        },
            "TOP FR":  { 
                        "chart_type": "pie",
                        "remove_total": True,
                        "marker_text":"{TOP_FR}",
                        }
                            
                            }          

def cm_to_pixels(cm, dpi=300):
    """Convertit des centimètres en pixels selon la résolution (DPI)"""
    return int(cm * dpi / 2.54)

def dataframe_to_image(df):
    """Convertit un DataFrame en image"""
    # Calculer les largeurs optimales des colonnes
    col_widths = []
    for col in df.columns:
        header_len = len(str(col))
        col_values = df[col].astype(str)
        max_data_len = col_values.map(len).max() if len(col_values) > 0 else 0
        col_widths.append(max(header_len, max_data_len))
    
    # Paramètres pour les dimensions
    char_width = 0.1  # Largeur d'un caractère en pouces
    row_height = 0.3  # Hauteur d'une ligne en pouces
    
    # Calculer dimensions
    table_width = sum(w * char_width for w in col_widths) + 1
    table_height = (len(df) + 1) * row_height
    
    # Créer figure
    fig = plt.figure(figsize=(table_width, table_height), frameon=False)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis('off')
    
    # Calculer largeurs relatives
    total_chars = sum(col_widths)
    col_width_ratios = [width/total_chars for width in col_widths]
    
    # Créer tableau
    table = ax.table(
        cellText=df.values.astype(str),
        colLabels=df.columns,
        cellLoc='center',
        loc='center',
        colWidths=col_width_ratios
    )
    
    table.scale(1, 1)
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    
    # Formater cellules
    for i, key in enumerate(table._cells):
        cell = table._cells[key]
        if key[0] == 0:  # En-têtes
            cell.set_text_props(weight='bold', color='white')
            cell.set_facecolor('#4472C4')
        else:
            if key[0] % 2:
                cell.set_facecolor('#E6EFF9')
            else:
                cell.set_facecolor('#ffffff')
        cell.set_edgecolor('#D4D4D4')
    
    # Sauvegarder en mémoire
    img_stream = BytesIO()
    plt.savefig(img_stream, format='png', bbox_inches='tight', dpi=300, pad_inches=0, 
                transparent=True)
    plt.close(fig)
    img_stream.seek(0)
    
    return img_stream, table_width, table_height

def crop_image_stream(img_stream, crop_top_cm=1.2, crop_bottom_cm=1.2, dpi=300):
    """Rogne l'image en haut et en bas avec une précision en centimètres"""
    # Ouvrir l'image avec PIL
    img = Image.open(img_stream)
    width, height = img.size
    
    # Convertir les cm en pixels
    crop_top_px = cm_to_pixels(crop_top_cm, dpi)
    crop_bottom_px = cm_to_pixels(crop_bottom_cm, dpi)
    
    # Vérifier que l'image est suffisamment grande pour le rognage demandé
    if height <= (crop_top_px + crop_bottom_px):
        print("Avertissement: Image trop petite pour le rognage demandé, rognage ajusté")
        # Ajuster proportionnellement
        total_crop = height // 2
        crop_top_px = total_crop // 2
        crop_bottom_px = total_crop // 2
    
    # Calculer les coordonnées du rectangle à conserver
    left = 0
    top = crop_top_px
    right = width
    bottom = height - crop_bottom_px
    
    # Rogner l'image
    cropped_img = img.crop((left, top, right, bottom))
    
    # Enregistrer en mémoire
    cropped_stream = BytesIO()
    cropped_img.save(cropped_stream, format='PNG')
    cropped_stream.seek(0)
    
    # Calculer les nouvelles dimensions en pouces
    new_width_inches = width / dpi
    new_height_inches = (bottom - top) / dpi
    
    return cropped_stream, new_width_inches, new_height_inches

def plot_to_image(dataframe,chart_type,title,remove_total=False):
    
    img_stream = BytesIO()

    if remove_total:
        data = dataframe.iloc[:-1]
    else:
        data = dataframe

    if chart_type == "line":
        plt.figure(figsize=(8, 4))

        # Select even indexed columns (0-based index) for plotting
        columns_to_plot = data.iloc[:, 1:].columns  # Start from the second column and take every column
        pourcentage_columns = [col for col in columns_to_plot if "Pourcentage" in str(col)]  # Find columns with percentage values
        if pourcentage_columns:
            columns_to_plot = columns_to_plot.drop(pourcentage_columns)  # Exclude percentage columns

        data[columns_to_plot] = (
            data[columns_to_plot]
            .astype(str)
            .replace("", 0)
            .replace(".", ",")
            .astype(int)
        )  # Convert columns to float

        values_to_plot = data[columns_to_plot].replace(0, np.nan)  # Replace 0 with NaN

        for col in columns_to_plot:
            plt.plot(data["Mois"], values_to_plot[col], label=col)

        plt.title(title)
        plt.xticks(rotation=45)
        plt.grid(axis="y")  # Only show horizontal gridlines

        # Dynamically set y-axis limits
        y_min = 0  # Start from 0
        y_max = values_to_plot.max().max() * 1.1  if values_to_plot.max().max() > 0 else 1 # Add 10% margin above the highest value
        plt.ylim(y_min, y_max)

        plt.gca().get_yaxis().set_major_formatter(
            plt.FuncFormatter(lambda x, _: f"{int(x):}")
        )  # Format y-axis to show complete values

        plt.legend()
        plt.tight_layout()

        plt.savefig(img_stream, format='png', bbox_inches='tight', dpi=300, pad_inches=0, transparent=False)
        plt.close()    
        img_stream.seek(0)       
        # Calculate the width and height of the figure in inches
        width, height = plt.gcf().get_size_inches()
            
    elif chart_type == "barchart" :

        categories = data['CompteLib'].tolist()
        # Columns after 'CompteLib' will be the legends (2024 in this case)
        columns = data.columns[2:]
        # Define bar width
        bar_width = 0.2
        # Create figure
        _, ax = plt.subplots(figsize=(10, 6))
        # Generate X locations for each category
        x_indexes = np.arange(len(categories))
        # Plot each year (or column) for each 'CompteLib'
        for i, col in enumerate(columns):
            ax.bar(x_indexes + i * bar_width, data[col], width=bar_width, label=col)

        # Adjust X-axis to have labels centered and corresponding to 'CompteLib'
        ax.set_xticks(x_indexes + bar_width * (len(columns) - 1) / 2)
        ax.set_xticklabels(categories)

        # Labels and legend
        ax.set_title("Evolution des investissements par nature")
        ax.legend(title="Annee")

        plt.xticks(rotation=45, ha='right')  # Rotate labels for readability
        plt.tight_layout()

        plt.savefig(img_stream, format='png', bbox_inches='tight', dpi=300, pad_inches=0, transparent=False)
        plt.close()    
        img_stream.seek(0)       
        # Calculate the width and height of the figure in inches
        width, height = plt.gcf().get_size_inches()

    else:
        
        fig = px.pie(
            data, 
            values="Pourcentage", 
            names=data.columns[1], 
            title=title,
            color_discrete_sequence=px.colors.qualitative.Set3  # Nice distinct colors
        )
        fig.update_layout(
            width=800,   # Set the width of the chart
            height=500,  # Set the height of the chart
        )
        dpi = 300
        width = fig.layout.width / dpi
        height = fig.layout.height / dpi

        pio.write_image(fig, img_stream, format='png', scale=3)  # scale factor for higher resolution
        img_stream.seek(0)

    
    
    return img_stream, width,height

def copy_run_style(source_run, target_run):
    """Copy run style (bold, italic, underline, font name, size, color, highlight)"""
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline

    # Font properties
    target_run.font.name = source_run.font.name
    target_run.font.size = source_run.font.size
    target_run.font.color.rgb = source_run.font.color.rgb if source_run.font.color else None
    target_run.font.highlight_color = source_run.font.highlight_color

def delete_section_after(doc, section_title):
    """
    Deletes a section starting from a Heading 1 with the given title 
    and all following content until the next Heading 1.
    """
    deleting = False  # Flag to start deletion
    paragraphs_to_delete = []  # Store paragraphs to delete
    section_found = False
    
    for i, para in enumerate(doc.paragraphs):
        if para.style.name == "Heading 1" and para.text.strip() == section_title:
            deleting = True  # Start deleting
            section_found = True
            paragraphs_to_delete.append(para)  # Mark the heading for deletion
            continue  # Move to next paragraph

        if deleting and para.style.name == "Heading 1":
            deleting = False  # Stop at the next Heading 1
            break
    
        if deleting:
            paragraphs_to_delete.append(para)  # Mark paragraphs for deletion

    # Actually remove the marked paragraphs (in reverse order to avoid parent issues)
    for para in reversed(paragraphs_to_delete):
        para._element.getparent().remove(para._element)
        
    # Clean up the document
    doc._body._element.xpath('//w:p[not(node())]')
    
    return section_found  # Return whether the section was found

def insert_images_from_excel_to_word(excel_path, word_template_path, word_output_path, config_elements_fixes,config_elements_variables,client,date):
    """Extrait les tableaux d'Excel, les rogne et les insère dans Word"""
    doc = Document(word_template_path)
    wb = load_workbook(excel_path)

    for table_name, config_item in config_elements_fixes.items():
        sheet_name = config_item["sheet_name"]
        marker_text_tableau = config_item["marker_text_tableau"]
        marker_text_graphique = config_item["marker_text_graphique"]
        chart_type = config_item["chart_type"]
        remove_total = config_item["remove_total"]
        title = table_name.replace("_"," ")
        section_title = config_item["section_title"]
        if sheet_name in wb.sheetnames:
            try:
                ws = wb[sheet_name]
                
                if table_name not in ws.tables:
                    print(f"Tableau '{table_name}' non trouvé dans la feuille '{sheet_name}'")
                    continue
                
                # Extraire données
                table_range = ws[ws.tables[table_name].ref]
                headers = [cell.value if cell.value is not None else "" for cell in table_range[0]]
                data = []
                for row in table_range[1:]:
                    row_data = [cell.value if cell.value is not None else "" for cell in row]
                    if any(row_data):
                        data.append(row_data)

                if not data:
                    print(f"Le tableau '{table_name}' ne contient pas de données")
                    continue
                
                # Créer DataFrame et convertir en image
                df = pd.DataFrame(data, columns=headers)
                percentage_columns = df.loc[:,df.columns.str.contains("%")]
                df = df.drop(columns=percentage_columns.columns)
                img_stream, _, _ = dataframe_to_image(df)
                
                # Rogner l'image (1.2cm en haut et en bas)
                cropped_stream, img_width, img_height = crop_image_stream(
                    img_stream, crop_top_cm=1.2, crop_bottom_cm=1.2)
                
                chart_img, chart_width, chart_height = plot_to_image(df,chart_type,title,remove_total)
                # Vérifier la largeur disponible
                section = doc.sections[0]
                available_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches
                                
                
                # Insérer l'image dans le document
                marker_tableau_found = False
                marker_graphique_found = False
                marker_tableau_replaced = False
                

                for paragraph in doc.paragraphs:
                    if marker_text_tableau in paragraph.text:
                        marker_tableau_found = True

                        if img_width > available_width:
                            scale_factor = available_width / img_width
                            width_to_use = Inches(available_width)
                            height_to_use = Inches(img_height * scale_factor)
                        else:
                            # Dimensions naturelles
                            width_to_use = Inches(img_width)
                            height_to_use = Inches(img_height)                                               
                        
                        # Supprimer espaces
                        paragraph.paragraph_format.space_before = 0
                        
                        paragraph.paragraph_format.line_spacing = 1.0
                        
                        # Remplacer le marqueur
                        paragraph.text = ""
                        run = paragraph.add_run()
                        run.add_picture(cropped_stream, width=width_to_use, height=height_to_use)
                        marker_tableau_replaced = True if paragraph.text == "" else False
                        break

                if not marker_tableau_replaced:
                    delete_section_after(doc,section_title)

                for paragraph in doc.paragraphs:    
                    if marker_text_graphique in paragraph.text:
                        marker_graphique_found = True

                        # Ajuster la taille de l'image du graphique
                        if chart_width > available_width:
                            if chart_type == "pie":
                                available_width = section.page_width.inches - (section.left_margin.inches + section.right_margin.inches)
                            width_to_use = Inches(available_width)
                        else:
                            # Dimensions naturelles
                            width_to_use = Inches(chart_width)

                        if chart_type == "pie":
                            width_to_use = width_to_use
                            height_to_use = Inches(chart_height)
                        else:
                            height_to_use = width_to_use/2

                        paragraph.text = ""
                        run = paragraph.add_run()
                        run.add_picture(chart_img, width=width_to_use, height=height_to_use)
                        break
                        
                if not marker_tableau_found:
                    print(f"Marqueur '{marker_text_tableau}' non trouvé dans le document")
                if not marker_graphique_found:
                    print(f"Marqueur '{marker_text_graphique}' non trouvé dans le document")

            except Exception as e:
                print(f"Erreur lors du traitement du tableau '{table_name}': {str(e)}")
        else :
            for paragraph in doc.paragraphs:
                if section_title in paragraph.text:
                    delete_section_after(doc,section_title) 
        
        
    sheet_variables = list(config_elements_variables.keys())
    
    # First find all markers and their positions
    marker_positions = {}
    for sheet_name in sheet_variables:
        marker_text = config_elements_variables[sheet_name]["marker_text"]
        if marker_text not in marker_positions:
            marker_positions[marker_text] = []
    
    # Find all marker occurrences in the document
    for i, paragraph in enumerate(doc.paragraphs):
        for marker_text in marker_positions.keys():
            if marker_text in paragraph.text:
                marker_positions[marker_text].append((i, paragraph))
    
    # Now process each sheet and find its corresponding marker
    section = doc.sections[0]
    available_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches
    
    for sheet_name in sheet_variables:
        marker_text = config_elements_variables[sheet_name]["marker_text"]
        chart_type = config_elements_variables[sheet_name]["chart_type"]
        remove_total = config_elements_variables[sheet_name]["remove_total"]
        
        # Skip if no markers found for this sheet
        if not marker_positions[marker_text]:
            print(f"Marqueur '{marker_text}' non trouvé dans le document")
            continue

        # Get the first occurrence of this marker
        _, marker_paragraph = marker_positions[marker_text][0]
        marker_positions[marker_text].pop(0)  # Remove this occurrence from the list


        # Clear the marker paragraph
        marker_paragraph.text = ""

        # Process all tables for this sheet
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            for table in list(ws.tables.keys()):
                try:
                    paragraph_title = table.split("_")[-1]
                    title = table.replace("_", " ")

                    # Process table data
                    table_range = ws[ws.tables[table].ref]
                    headers = [cell.value if cell.value is not None else "" for cell in table_range[0]]
                    data = []
                    for row in table_range[1:]:
                        row_data = [cell.value if cell.value is not None else "" for cell in row]
                        if any(row_data):
                            data.append(row_data)
                    if not data:
                        print(f"Le tableau '{table}' ne contient pas de données")
                        continue

                    df = pd.DataFrame(data, columns=headers)
                    percentage_columns = df.loc[:, df.columns.str.contains("%")]
                    df = df.drop(columns=percentage_columns.columns)
                    img_stream, _, _ = dataframe_to_image(df)
                    cropped_stream, img_width, img_height = crop_image_stream(
                        img_stream, crop_top_cm=1.2, crop_bottom_cm=1.2)
                    chart_img, chart_width,chart_height = plot_to_image(df, chart_type, title, remove_total)

                    # Calculate dimensions
                    if img_width > available_width:
                        scale_factor = available_width / img_width
                        width_to_use_tableau = Inches(available_width)
                        height_to_use_tableau = Inches(img_height * scale_factor)
                    else:
                        width_to_use_tableau = Inches(img_width)
                        height_to_use_tableau = Inches(img_height)


                    width_to_use_graphique = Inches(chart_width)*2.5

                    if chart_type == "pie":
                        height_to_use_graphique = Inches(chart_height)*2.5
                    else:
                        height_to_use_graphique = width_to_use_graphique / 2

                    # 1. Add Heading 3 paragraph **WITHOUT** extra line at the top
                    heading_para = marker_paragraph.insert_paragraph_before(paragraph_title, style='Heading 3')
                    # Prevent page break before this heading
                    heading_para._p.get_or_add_pPr().set('keepNext', '1')

                    # 2. Add table image paragraph with appropriate spacing
                    table_para = marker_paragraph.insert_paragraph_before()
                    table_run = table_para.add_run()
                    table_run.add_picture(cropped_stream, width=width_to_use_tableau, height=height_to_use_tableau)
                    # Prevent page break between heading and table
                    table_para._p.get_or_add_pPr().set('keepWithNext', '1')

                    # 3. Add chart image paragraph
                    chart_para = marker_paragraph.insert_paragraph_before()
                    chart_run = chart_para.add_run()
                    chart_run.add_picture(chart_img, width=width_to_use_graphique, height=height_to_use_graphique)
                    
                except Exception as e:
                    print(f"Erreur lors du traitement du tableau '{table}': {str(e)}")
    
    marker_descriptifs = ["{CLIENT}",'{DATE}']

    marker_titre = f'{marker_descriptifs[0]} au {marker_descriptifs[1]}'

    marker_titre_found = False
    
    for paragraph in doc.paragraphs:
        source_run = None
        if "Revue analytique" in paragraph.text:
            source_run = paragraph.runs[0]
            break

    for paragraph in doc.paragraphs:    
        target_run = None
        if marker_titre in paragraph.text:
            marker_titre_found = True
            paragraph.text = paragraph.text.replace(marker_titre,f'{client} au {date}')
            target_run = paragraph.runs[0]
            copy_run_style(source_run, target_run)
            break


    if not marker_titre_found:
        print(f"Did not find {marker_titre} in the document !") 


    # Sauvegarder le document
    doc.save(word_output_path)

# class ModernButton(ttk.Button):
#     """Custom styled button class"""
#     def __init__(self, master=None, **kwargs):
#         super().__init__(master, **kwargs)
#         self.bind('<Enter>', lambda e: self.on_hover())
#         self.bind('<Leave>', lambda e: self.on_leave())
    
#     def on_hover(self):
#         self.state(['pressed'])
    
#     def on_leave(self):
#         self.state(['!pressed'])

# class ReportGeneratorApp:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("Générateur de reporting")
#         self.root.geometry("900x800")  # Increased height to accommodate new client/date fields
#         self.root.configure(bg='#f0f2f5')
        
#         self.setup_styles()
        
#         self.container = ttk.Frame(root, style='Container.TFrame')
#         self.container.pack(fill=tk.BOTH, expand=True, padx=40, pady=30)
        
#         self.header_frame = ttk.Frame(self.container, style='Header.TFrame')
#         self.header_frame.pack(fill=tk.X, pady=(0, 30))
        
#         title_label = ttk.Label(
#             self.header_frame,
#             text="Générateur de reporting",
#             style="Title.TLabel"
#         )
#         title_label.pack()
        
#         subtitle_label = ttk.Label(
#             self.header_frame,
#             text="Générer un reporting à partir d'une revue analytique",
#             style="Subtitle.TLabel"
#         )
#         subtitle_label.pack(pady=(5, 0))
        
#         self.card_frame = ttk.Frame(self.container, style='Card.TFrame')
#         self.card_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
#         # Input file selection area
#         self.file_frame = ttk.Frame(self.card_frame, style='Inner.TFrame')
#         self.file_frame.pack(fill=tk.X, padx=30, pady=20)
        
#         input_label = ttk.Label(
#             self.file_frame,
#             text="Fichier Excel:",
#             style="SectionTitle.TLabel"
#         )
#         input_label.pack(anchor='w', pady=(0, 5))
        
#         self.file_label = ttk.Label(
#             self.file_frame,
#             text="Aucun fichier sélectionné",
#             style="FileLabel.TLabel"
#         )
#         self.file_label.pack(side=tk.LEFT, pady=10)
        
#         self.browse_button = ModernButton(
#             self.file_frame,
#             text="Parcourir le fichier Excel",
#             style="Action.TButton",
#             command=self.browse_file
#         )
#         self.browse_button.pack(side=tk.RIGHT, pady=10)
        
#         # Separator
#         ttk.Separator(self.card_frame, orient='horizontal').pack(fill=tk.X, padx=30)
        
#         # Output location selection area
#         self.output_frame = ttk.Frame(self.card_frame, style='Inner.TFrame')
#         self.output_frame.pack(fill=tk.X, padx=30, pady=20)
        
#         output_label = ttk.Label(
#             self.output_frame,
#             text="Emplacement de sortie:",
#             style="SectionTitle.TLabel"
#         )
#         output_label.pack(anchor='w', pady=(0, 5))
        
#         self.output_label = ttk.Label(
#             self.output_frame,
#             text="Par défaut: ./generated_reports",
#             style="FileLabel.TLabel"
#         )
#         self.output_label.pack(side=tk.LEFT, pady=10)
        
#         self.output_button = ModernButton(
#             self.output_frame,
#             text="Choisir l'emplacement",
#             style="Action.TButton",
#             command=self.choose_output_location
#         )
#         self.output_button.pack(side=tk.RIGHT, pady=10)
        
#         # Separator
#         ttk.Separator(self.card_frame, orient='horizontal').pack(fill=tk.X, padx=30)
        
#         # Client and Date input area
#         self.client_date_frame = ttk.Frame(self.card_frame, style='Inner.TFrame')
#         self.client_date_frame.pack(fill=tk.X, padx=30, pady=20)
        
#         # Client input
#         client_label = ttk.Label(
#             self.client_date_frame,
#             text="Client:",
#             style="SectionTitle.TLabel"
#         )
#         client_label.pack(anchor='w', pady=(0, 5))
        
#         self.client_var = tk.StringVar()
#         self.client_entry = ttk.Entry(
#             self.client_date_frame,
#             textvariable=self.client_var,
#             width=40
#         )
#         self.client_entry.pack(anchor='w', pady=(0, 15))
        
#         # Date input
#         date_label = ttk.Label(
#             self.client_date_frame,
#             text="Date:",
#             style="SectionTitle.TLabel"
#         )
#         date_label.pack(anchor='w', pady=(0, 5))
        
#         self.date_var = tk.StringVar()
#         self.date_entry = ttk.Entry(
#             self.client_date_frame,
#             textvariable=self.date_var,
#             width=40
#         )
#         self.date_entry.pack(anchor='w', pady=(0, 10))
        
#         # Separator
#         ttk.Separator(self.card_frame, orient='horizontal').pack(fill=tk.X, padx=30)
        
#         # Action area
#         self.action_frame = ttk.Frame(self.card_frame, style='Inner.TFrame')
#         self.action_frame.pack(fill=tk.X, padx=30, pady=20)
        
#         self.generate_button = ModernButton(
#             self.action_frame,
#             text="Générer le reporting",
#             style="Generate.TButton",
#             command=self.generate_clicked,
#             state="disabled"
#         )
#         self.generate_button.pack(pady=10)
        
#         # Progress frame
#         self.progress_frame = ttk.Frame(self.card_frame, style='Inner.TFrame')
#         self.progress_frame.pack(fill=tk.X, padx=30, pady=(0, 20))
        
#         self.progress = ttk.Progressbar(
#             self.progress_frame,
#             mode='indeterminate',
#             length=400,
#             style='Modern.Horizontal.TProgressbar'
#         )
        
#         self.status_label = ttk.Label(
#             self.progress_frame,
#             text="",
#             style="Status.TLabel"
#         )
#         self.status_label.pack(pady=5)
        
#         # Store paths
#         self.template_path = "./CODE/template reporting.docx"
#         self.output_folder = os.path.join(os.getcwd(), "output")
#         self.selected_file = None
        
#         # Create default output folder
#         os.makedirs(self.output_folder, exist_ok=True)

#     def setup_styles(self):
#         """Configure custom styles for the application"""
#         style = ttk.Style()
        
#         # Previous style configurations remain the same
#         bg_color = '#f0f2f5'
#         card_bg = '#ffffff'
#         accent_color = '#2563eb'
#         hover_color = '#1d4ed8'
#         success_color = '#059669'
        
#         # Add new style for section titles
#         style.configure('SectionTitle.TLabel',
#             font=('Segoe UI', 12, 'bold'),
#             background=card_bg,
#             foreground='#1e293b'
#         )
        
#         # Previous style configurations remain the same
#         style.configure('Container.TFrame', background=bg_color)
#         style.configure('Header.TFrame', background=bg_color)
#         style.configure('Card.TFrame', background=card_bg)
#         style.configure('Inner.TFrame', background=card_bg)
        
#         style.configure('Title.TLabel',
#             font=('Segoe UI', 24, 'bold'),
#             background=bg_color,
#             foreground='#1e293b'
#         )
        
#         style.configure('Subtitle.TLabel',
#             font=('Segoe UI', 12),
#             background=bg_color,
#             foreground='#64748b'
#         )
        
#         style.configure('FileLabel.TLabel',
#             font=('Segoe UI', 11),
#             background=card_bg,
#             foreground='#64748b'
#         )
        
#         style.configure('Status.TLabel',
#             font=('Segoe UI', 11),
#             background=card_bg,
#             foreground='#64748b'
#         )
        
#         style.configure('Action.TButton',
#             font=('Segoe UI', 11),
#             background=accent_color,
#             padding=(15, 8)
#         )
        
#         style.configure('Generate.TButton',
#             font=('Segoe UI', 11, 'bold'),
#             background=accent_color,
#             padding=(20, 10)
#         )
        
#         style.map('Action.TButton',
#             background=[('pressed', hover_color), ('active', hover_color)],
#             foreground=[('pressed', 'white'), ('active', 'white')]
#         )
        
#         style.map('Generate.TButton',
#             background=[('pressed', hover_color), ('active', hover_color)],
#             foreground=[('pressed', 'white'), ('active', 'white')]
#         )
        
#         style.configure('Modern.Horizontal.TProgressbar',
#             troughcolor=bg_color,
#             background=success_color,
#             thickness=8
#         )

#     def browse_file(self):
#         """Open file dialog to select Excel file"""
#         file_path = filedialog.askopenfilename(
#             filetypes=[("Excel files", "*.xlsx")]
#         )
#         if file_path:
#             self.selected_file = file_path
#             filename = os.path.basename(file_path)
#             self.file_label.config(text=f"Selected: {filename}")
#             self.generate_button.config(state="normal")
#             self.status_label.config(text="")

#     def generate_clicked(self):
#         """Handle generate button click"""
#         if self.selected_file:
#             self.process_file(self.selected_file)

#     def choose_output_location(self):
#         """Allow user to choose output location"""
#         folder_path = filedialog.askdirectory(
#             title="Choisir l'emplacement de sortie",
#             initialdir=self.output_folder
#         )
#         if folder_path:
#             self.output_folder = folder_path
#             display_path = self.truncate_path(folder_path)
#             self.output_label.config(text=f"Emplacement: {display_path}")

#     def truncate_path(self, path, max_length=50):
#         """Truncate path if it's too long"""
#         if len(path) <= max_length:
#             return path
        
#         parts = Path(path).parts
#         if len(parts) <= 2:
#             return path
        
#         return os.path.join('...', *parts[-2:])

#     def process_file(self, excel_path):
#         """Process the Excel file and generate report"""
#         self.browse_button.config(state="disabled")
#         self.output_button.config(state="disabled")
#         self.generate_button.config(state="disabled")
        
#         self.progress.pack(fill=tk.X, pady=(10, 5))
#         self.progress.start()
#         self.status_label.config(text="Génération du reporting...")
        
#         # Get client and date values from the entry fields
#         client = self.client_var.get().strip()
#         date = self.date_var.get().strip()

#         base_name = Path(excel_path).stem
#         output_path = os.path.join(
#             self.output_folder,
#             f'{" ".join(list(reversed(date.split("/"))))} {client} Reporting.docx'
#         )

#         thread = threading.Thread(
#             target=self.generate_report,
#             args=(excel_path, output_path, client, date)
#         )
#         thread.start()

#     def generate_report(self, excel_path, output_path, client, date):
#         """Generate the report in a separate thread"""
#         try:
#             insert_images_from_excel_to_word(
#                 excel_path,
#                 self.template_path,
#                 output_path,
#                 config_elements_fixes,
#                 config_elements_variables,
#                 client,
#                 date
#             )
#             self.root.after(0, self.handle_success, output_path)
#         except Exception as e:
#             self.root.after(0, self.handle_error, str(e))

#     def handle_success(self, output_path):
#         """Handle successful report generation"""
#         self.progress.stop()
#         self.progress.pack_forget()
#         self.status_label.config(
#             text="Reporting généré avec succès!"
#         )
        
#         self.browse_button.config(state="normal")
#         self.output_button.config(state="normal")
#         self.generate_button.config(state="normal")
        
#         if messagebox.askyesno(
#             "Succès",
#             "Reporting généré avec succès! Voulez-vous l'ouvrir?"
#         ):
#             self.open_file(output_path)

#     def handle_error(self, error_message):
#         """Handle report generation error"""
#         self.progress.stop()
#         self.progress.pack_forget()
#         self.status_label.config(text="Erreur lors de la génération du reporting")
        
#         self.browse_button.config(state="normal")
#         self.output_button.config(state="normal")
#         self.generate_button.config(state="normal")
        
#         messagebox.showerror("Erreur", f"Erreur lors de la génération du reporting: {error_message}")

#     def open_file(self, file_path):
#         """Open the generated report with default application"""
#         if os.path.exists(file_path):
#             try:
#                 if os.name == 'nt':  # Windows
#                     os.startfile(file_path)
#                 elif os.name == 'posix':  # macOS and Linux
#                     subprocess.run(['open' if os.name == 'darwin' else 'xdg-open', file_path])
#             except Exception as e:
#                 messagebox.showerror("Erreur", f"Erreur lors de l'ouverture du fichier: {str(e)}")
#         else:
#             messagebox.showerror("Erreur", "Fichier non trouvé")

# def main():
#     root = tk.Tk()
#     app = ReportGeneratorApp(root)
#     root.mainloop()

# if __name__ == "__main__":
#     main()