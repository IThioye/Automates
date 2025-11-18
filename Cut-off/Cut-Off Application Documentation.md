# Cut-Off Application Documentation

## Aperçu
Cette application est un outil de gestion des cut-offs pour les ventes et achats dans un contexte comptable. Elle permet de traiter les Fichiers des Écritures Comptables (FEC) et de générer des rapports détaillés pour l'analyse des cut-offs.

## Fonctionnalités principales
- Traitement des FEC pour deux années consécutives
- Analyse des cut-offs pour :
  - Ventes
  - Avoirs de vente
  - Achats
  - Avoirs d'achat
- Interface graphique intuitive
- Génération de rapports Excel détaillés
- Sélection flexible des périodes d'analyse

## Dépendances
```python
# Bibliothèques requises
pandas
numpy
tkinter
openpyxl
unidecode
```

## Process de fonctionnement

### 1. Sélection des fichiers
- L'utilisateur sélectionne les FEC (fichiers .txt) à analyser
- L'application accepte les fichiers pour deux années consécutives
- Détection automatique de l'année des fichiers basée sur leur nom

### 2. Configuration des périodes
Deux modes de sélection :
- **Mode automatique** : Spécification du nombre de mois avant/après la clôture
- **Mode manuel** : Sélection précise des périodes pour chaque type de document

### 3. Traitement des données
Pour chaque fichier :
1. Lecture et décodage du FEC (UTF-8 ou MacRoman)
2. Conversion des dates et formatage des valeurs numériques
3. Filtrage des écritures selon les critères de cut-off
4. Calcul des totaux et pourcentages

### 4. Génération du rapport
Création d'un fichier Excel avec :
- Une feuille par type de cut-off
- Tables formatées avec couleurs distinctives
- Sections pour les documents apportés par le client
- Calculs automatiques des totaux et pourcentages

## Structure du fichier de sortie

### Feuilles de calcul
1. **Cut-Off Ventes**
   - Données de l'année N-1
   - Période de fin d'exercice

2. **Cut-Off Avoirs de vente**
   - Données de l'année N
   - Période de début d'exercice suivant

3. **Cut-Off Achats**
   - Données de l'année N
   - Période de début d'exercice suivant

4. **Cut-Off Avoirs d'achat**
   - Données de l'année N-1
   - Période de fin d'exercice

### Format des données
Chaque feuille contient :
- Titre et période analysée
- Table des écritures sélectionnées
- Totaux et statistiques
- Espace pour les documents justificatifs du client

## Notes importantes
- L'application nécessite deux FEC pour un résultat complet
- Les fichiers doivent être au format texte (.txt) avec séparateur tabulation ou pipe (|)
- Le format de date attendu dans les FEC est YYYYMMDD
- L'interface permet la correction manuelle des années détectées
- Les seuils de sélection sont paramétrés (ex: 3000€ pour les avoirs)

## Codes couleur
- Ventes : Bleu (#BDD7EE)
- Avoirs de vente : Orange (#FFE699)
- Achats : Violet (#E5E5FF)
- Avoirs d'achat : Vert