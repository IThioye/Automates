from utils.utils import resource_path
from .setting_windows import SettingsWindow
from config.config_manager import load_config

from logic.word_reporting import insert_images_from_excel_to_word,config_elements_fixes,config_elements_variables
from logic.calculations import *
from logic.excel_reporting import add_df_to_sheet

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from chardet import detect
import sys
import os
import re
from pathlib import Path
import datetime
import threading
import pandas as pd
import traceback
from openpyxl import Workbook


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

    def check_duplicated_years(self,processed_files):
        """
        Vérifie les doublons d'années dans les fichiers FEC sélectionnés.
        """
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

    def load_files(self,liste_encoding):
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
        return df_by_year


    def clean_data(self,df):
        """
        Nettoie et prépare les données du DataFrame FEC.
        
        """
        df.columns = df.columns.str.strip()
        df["EcritureDate"] = pd.to_datetime(df["EcritureDate"], format='%Y%m%d', errors='coerce')
        df['Année'] = df["EcritureDate"].dt.year
        df['Mois'] = df["EcritureDate"].dt.month_name(locale='French')
        df["Credit"] = pd.to_numeric(df["Credit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0)
        df["Debit"] = pd.to_numeric(df["Debit"].str.replace(',', '.', regex=True), errors='coerce').fillna(0) 

    def process_files(self):
        """
        Traite les fichiers FEC sélectionnés.
        """
        if not self.fec_files:
            messagebox.showwarning("Attention", "Aucun fichier à traiter")
            return
        
        processed_files = []        
        self.check_duplicated_years(processed_files)
        
        liste_encoding = self.parametres['liste_encoding']
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
            df_by_year = self.load_files(liste_encoding)
            
            # Open the account configuration window
            # Only open the salary config window if user requested to exclude salaries
            if exclude_salaries:
                account_config = self.create_frais_personnels_config_window(df_by_year)
                if not account_config:
                    return
                
            if exclude_cpte_fournisseur and config['sheets']['Crédit Moyen FR']:
                excluded_cmf_config = self.exclusion_credit_moyen(df_by_year,
                                                                  fournisseurs_codes,
                                                                  journal_code_achats
                                                                  )
                if not excluded_cmf_config:
                    return
                
            if config['sheets']['Crédit Moyen Client']:
                cmc_account_config = self.create_credit_moyen_client_config_window(df_by_year,taux_cmc)
                if not cmc_account_config:
                    return

            if config['sheets']['TOP FR'] or config['sheets']['TOP Client']: 
                top_dix_account_config = self.create_top_dix_config_window(df_by_year,
                                                                           fournisseurs_codes=fournisseurs_codes,
                                                                           clients_codes=clients_codes,
                                                                           journal_code_achats=journal_code_achats,
                                                                           journal_code_ventes=journal_code_ventes
                                                                           )
                if not top_dix_account_config:
                    return

            for year, df in df_by_year.items():
                    
                # Préparation des données
                self.clean_data(df)

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
                    if top_dix_client is not None and not top_dix_client.empty and compte_lib_column_client is not None:
                        top_cinq_clients = list(top_dix_client[compte_lib_column_client].head(5))
                    else:
                        print("⚠️ Skipping top 5 clients: no valid data.")
                        top_cinq_clients = []

                    all_results['Top_CLT'].append((f"Top 10 Clients {year}", top_dix_client))
                
                if config['sheets']['Achats et Sous-traitance']:
                    all_results["Achats et Sous-traitance"].append((f"Evolutions des achats et de la sous-traitance {year}", table_achat_sous_traitance))

                if config['sheets']['Crédit Moyen Client'] and config['sheets']['TOP Client'] and len(top_cinq_clients) > 0:
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
                ca_final_cumule_year_columns = [col for col in ca_final_cumule.columns if col != 'Mois' and '%' not in col][-2:] # garder seulement les deux dernières années

                if len(ca_final_cumule_year_columns) >= 2:                
                    # Define the two most recent year columns
                    year_prev, year_curr = ca_final_cumule_year_columns[0], ca_final_cumule_year_columns[1]

                    # Define the name of the new column dynamically
                    variation_col = f"% Variation {year_curr}-{year_prev}"

                    # Calculate year-over-year % variation (only if both values are valid and non-zero)
                    def compute_variation(row):
                        val_prev = row[year_prev]
                        val_curr = row[year_curr]
                        
                        if pd.notna(val_prev) and pd.notna(val_curr) and val_prev != 0 and val_curr != 0:
                            variation = ((val_curr - val_prev) / abs(val_prev)) * 100
                            return f"{round(variation)}%"
                        else:
                            return 0

                    # Apply to each row
                    ca_final_cumule[variation_col] = ca_final_cumule.apply(compute_variation, axis=1)


            if config['sheets']['Crédit Moyen FR'] and cmf_jours_dataframes:
                cmf_jours_dataframes = pd.concat(cmf_jours_dataframes, axis=1)
                cmf_jours_dataframes = cmf_jours_dataframes.loc[:, ~cmf_jours_dataframes.columns.duplicated()]

            if config['sheets']['Crédit Moyen FR'] and cmf_top_5_jours_dataframes and config['sheets']['TOP FR']:
                for year,_ in cmf_top_5_jours_dataframes.items():
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
            
            if config['sheets']['Crédit Moyen Client'] and cmc_top_5_jours_dataframes and config['sheets']['TOP Client'] and len(cmc_top_5_jours_dataframes) >= 1:
                for year,_ in cmc_top_5_jours_dataframes.items():
                    if cmc_top_5_jours_dataframes.get(year):
                        # Only concatenate if the list is not empty
                        cmc_top_5_jours_dataframes[year] = pd.concat(cmc_top_5_jours_dataframes[year], axis=1)
                        cmc_top_5_jours_dataframes[year] = cmc_top_5_jours_dataframes[year].loc[:, ~cmc_top_5_jours_dataframes[year].columns.duplicated()]
                    else:
                        print(f"⚠️ Skipping top 5 credit moyen for {year}: no valid data.")
                        cmc_top_5_jours_dataframes[year] = pd.DataFrame()
                    
            
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
            
            if config['sheets']['Crédit Moyen Client'] and cmc_top_5_by_year_dataframes and config['sheets']['TOP Client'] and len(cmc_top_5_by_year_dataframes) > 0:
                add_df_to_sheet(wb, "Crédit Moyen CLT TOP 5", *cmc_top_5_by_year_dataframes[-1])
                if len(cmc_top_5_by_year_dataframes) > 1:
                    add_df_to_sheet(wb, "Crédit Moyen CLT TOP 5", *cmc_top_5_by_year_dataframes[-2],start_column=max_length_cmc+3)
                start_column = 2*(max_length_cmc + 3) if len(cmc_top_5_by_year_dataframes) > 1 else max_length_cmc + 3
                print(40*"-","\nDEBUG: CMC Top 5 Jours DataFrames:")
                print(cmc_top_5_jours_dataframes)
                print(40*"-")
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
