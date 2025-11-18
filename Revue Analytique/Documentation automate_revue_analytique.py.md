# Documentation : automate_revue_analytique.py

## Aperçu

Ce script Python propose une application graphique (GUI) permettant d'automatiser la revue analytique à partir de fichiers FEC (Fichier des Écritures Comptables). Il facilite l'import, le traitement et l'analyse de données comptables, puis génère des rapports Excel et Word personnalisés. L'utilisateur peut configurer les paramètres d'analyse, choisir les feuilles à inclure, exclure certains comptes, et personnaliser les calculs selon le type d'entreprise.

## Fonctionnement général

1. **Interface graphique (Tkinter)**  
   L'application s'ouvre avec une fenêtre principale permettant :
   - L'ajout de fichiers FEC (plusieurs années possibles)
   - La configuration des paramètres d'analyse (codes comptes, encodages, etc.)
   - La sélection des feuilles à générer selon le type d'entreprise
   - L'exclusion de certains comptes pour les analyses fines (salaires, clients, fournisseurs)
   - Le lancement du traitement et la génération du fichier Excel

2. **Traitement des fichiers FEC**
   - Lecture des fichiers FEC avec gestion automatique de l'encodage
   - Nettoyage et préparation des données (dates, montants, comptes)
   - Calculs analytiques : chiffre d'affaires, trésorerie, achats, salaires, crédit moyen client/fournisseur, top 10 clients/fournisseurs, investissements
   - Génération de tableaux et graphiques pour chaque analyse

3. **Personnalisation et configuration**
   - Fenêtre de paramètres pour modifier les codes, encodages, journaux, etc.
   - Import/export de configurations au format JSON
   - Fenêtres de configuration avancée pour exclure des comptes ou paramétrer les taux de TVA

4. **Export et reporting**
   - Génération d'un fichier Excel structuré avec tableaux et graphiques
   - Option de génération d'un reporting Word à partir d'un template, incluant les graphiques Excel

## Dépendances nécessaires

- **Python ≥ 3.7**
- **Tkinter** (interface graphique)
- **Pandas** (traitement des données)
- **Numpy** (calculs numériques)
- **Openpyxl** (création et manipulation de fichiers Excel)
- **Chardet** (détection d'encodage)
- **Unidecode** (normalisation de texte)
- **Pathlib** (gestion des chemins)
- **Json** (gestion des configurations)
- **Datetime** (gestion des dates)
- **Subprocess** (ouverture de fichiers)
- **Assets/automate_reporting.py** (module interne pour le reporting Word)

## Structure principale

- **FECProcessor** : classe principale de l'application, gère l'interface et le workflow utilisateur.
- **SettingsWindow** : fenêtre de configuration des paramètres par défaut.
- **Fonctions analytiques** : calculs pour chaque indicateur financier (CA, trésorerie, achats, etc.).
- **Fonctions de mise en forme** : création de tableaux Excel, ajustement des colonnes, coloration, génération de graphiques.
- **Fenêtres de configuration avancée** : pour la sélection/exclusion de comptes et la personnalisation des analyses.
- **Reporting Word** : génération optionnelle d'un rapport Word à partir du fichier Excel.

## Utilisation

1. Lancer le script :  
   `python automate_revue_analytique.py` ou directement l'application.

2. Ajouter les fichiers FEC à analyser.

3. Configurer les paramètres si besoin (codes, encodages, journaux).

4. Sélectionner les feuilles à générer et les options d'analyse.

5. Lancer le traitement et enregistrer le fichier Excel généré.

6. (Optionnel) Générer un reporting Word à partir du fichier Excel.

## Notes

- Les fichiers de configuration sont sauvegardés dans le dossier `assets` sous le nom `configuration.json`.
- Les templates Word pour le reporting doivent être placés dans le dossier `assets`.
- L'application gère les erreurs et propose des fenêtres d'information détaillées en cas de problèmes.