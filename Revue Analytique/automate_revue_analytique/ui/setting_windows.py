import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json 
import sys
from config.config_manager import load_config, save_config

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

