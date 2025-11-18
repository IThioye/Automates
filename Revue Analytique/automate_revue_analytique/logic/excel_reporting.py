from openpyxl import Workbook
from openpyxl.chart import (LineChart, PieChart, BarChart3D, Reference)
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill,Alignment,Border, Side
from openpyxl.drawing.text import Paragraph, ParagraphProperties, RegularTextRun,CharacterProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from unidecode import unidecode
import re

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
    from openpyxl.styles import Font

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
                                # Bold cells in "Total" rows
                                if start_col == "Total":
                                    cell.font = Font(bold=True)
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
    positive_green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    negative_red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    
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
                        if " Variation" in cell.value:
                            for row in range(current_row + 1, last_row + 1):
                                target_cell = ws.cell(row=row, column=cell.column)
                                try:
                                    cell_value = target_cell.value
                                    if "-" in cell_value:
                                        target_cell.fill = negative_red_fill
                                    elif cell_value in ["", None,0,0.0]:
                                        continue
                                    else:
                                        target_cell.fill = positive_green_fill
                                except (ValueError, TypeError):
                                    continue
                            
                        else:
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
        thin_border = Border(
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
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        ignore_sheetnames = ["TOP CLT","TOP FR"]

            
        # Write DataFrame data
        for row_idx, row in enumerate(df.values, start=current_row + 1):
            for col_idx, value in enumerate(row, start=start_column):
                cell = ws.cell(row=row_idx, column=col_idx)
                
                if sheet_name not in ignore_sheetnames:
                    if value in (0, "0", 0.0, None, ""):
                        cell.value = ""
                    else:
                        cell.value = value
                        # Add thousand separator for numeric values
                        if isinstance(value, (int, float)) and not isinstance(value, bool):
                            cell.number_format = '_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-'


                else:
                    cell.value = value
                    # Add thousand separator for numeric values
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        cell.number_format = '_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-'
                
                # Center all cells
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border

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

