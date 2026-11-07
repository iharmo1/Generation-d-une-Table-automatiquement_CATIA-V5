# Génération automatique d'une table dans CATIA V5

Automatisation du processus de création d'une pièce en 3D : une table (poste de travail en structure tubulaire avec plateaux) est générée automatiquement dans **CATIA V5**, sans la dessiner à la main.

Le même résultat est obtenu par **3 méthodes différentes**, pour comparer les approches.

## Aperçu

Table Génerée

<img width="959" height="465" alt="image" src="https://github.com/user-attachments/assets/74e1ffcd-ea19-48ca-8a92-fbb5e60bb6d2" />
<img width="960" height="498" alt="demo_catia" src="https://github.com/user-attachments/assets/359e9d2b-90c7-4117-ad3d-07f8b87e4093" />




*Résultat dans CATIA V5 : vue isométrique (en haut à gauche) et vues orthogonales de la table.*

## Les 3 méthodes

| Méthode | Fichier | Où l'exécuter |
|---|---|---|
| VBA dans Excel | `Table3_Excel_CATIA.bas` | Excel, qui pilote CATIA |
| CATScript (VBA de CATIA) | `Table3_PartDesign_v5.CATScript` | Directement dans CATIA V5 |
| Python | `Table3.py` | Un terminal ou VS Code, qui pilote CATIA |

Le fichier `Description.md` détaille ces méthodes de génération.

## Prérequis

- Windows avec **CATIA V5** installé (licence Part Design)
- **Microsoft Excel** (pour la méthode VBA Excel)
- **Python 3** (pour la méthode Python) et les bibliothèques utilisées dans `Table3.py` *(à compléter, par exemple `pywin32`)*

## Utilisation

### Méthode 1 : VBA Excel
1. Ouvrez Excel, puis l'éditeur VBA avec `Alt + F11`.
2. Faites Fichier > Importer un fichier et choisissez `Table3_Excel_CATIA.bas`.
3. Ouvrez CATIA V5, puis lancez la macro depuis Excel.

### Méthode 2 : CATScript
1. Dans CATIA V5, faites Outils > Macro > Macros (`Alt + F8`).
2. Cliquez sur « Bibliothèques de macros… » et ajoutez le dossier contenant `Table3_PartDesign_v5.CATScript`.
3. Sélectionnez la macro et cliquez sur Exécuter.

### Méthode 3 : Python
1. Ouvrez CATIA V5.
2. Dans un terminal, lancez :
   ```
   python Table3.py
   ```

## Résultat

Le fichier `Table __.CATPart` est un exemple de pièce générée par les scripts. Ouvrez-le dans CATIA V5 pour voir le résultat final.

## Contenu du dépôt

```
├── Description.md                    Détail des méthodes de génération
├── Table3.py                         Méthode Python
├── Table3_Excel_CATIA.bas            Méthode VBA Excel
├── Table3_PartDesign_v5.CATScript    Méthode CATScript
├── Table __.CATPart                  Pièce exemple générée
└── table_catia.png                   Capture d'écran du résultat
```

## Auteur

[iharmo1](https://github.com/iharmo1)
