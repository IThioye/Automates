
from utils.utils import resource_path
import json
import os
from tkinter import messagebox

CONFIG_FILE = resource_path('assets/configuration.json')

LISTE_ENCODING = ["utf-8","UTF-16","MacRoman","SHIFT_JIS","CP932"]
JOURNAL_CODES_ACHATS = ["AC", "HA", "A10","60","ACH"]
JOURNAL_CODES_VENTES = ["VE", "VT", "VTE","70","VEN"]
FOURNISSEURS_CODE = ["401"]
CLIENTS_CODE = ["411"]
TAUX_CMC = ['0%','2.1%','5.5%','8.5%','10%','20%']

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