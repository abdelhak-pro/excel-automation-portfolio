# Portfolio d'automatisation Excel
Projets Python pour automatiser des fichiers Excel : nettoyage, fusion et rapports.

## À propos
J'apprends Python pour automatiser les tâches Excel répétitives des petites entreprises. Ce dépôt suit ma progression : chaque exercice est testé et vérifié.

Ce dépôt montre comment préparer automatiquement un rapport Excel : fusion de fichiers, nettoyage des données (espaces, casse, doublons, valeurs manquantes) et génération d'un rapport formaté.

## Projet phare : `day13.py`
**Entrée :** `test_final.xlsx` (fichier de ventes désordonné)
**Sortie :** `rapport_test.xlsx` avec 3 feuilles : `Par product`, `Par ville`, `Données nettoyées`

![Rapport généré](rapport_test.PNG)

### Utilisation
1. Installer les bibliothèques : `pip install pandas openpyxl`
2. Placer `test_final.xlsx` dans le même dossier que `day13.py`
3. Lancer : `python day13.py`

### Règles de nettoyage appliquées
- Espaces en trop supprimés, première lettre en majuscule (produit et ville)
- Doublons supprimés après uniformisation
- Produit manquant : remplacé par `Vide` (le montant est conservé)
- Montant manquant : remplacé par `0`
- Contrôle : le total des deux rapports est identique (1 370)

## Contenu
| Fichier | Description |
|---------|-------------|
| `day1.py` | Somme d'une liste avec une boucle `for` |
| `day2.py` | Somme de prix, puis somme des ventes supérieures à 100 avec `if` |
| `day3.py` | Lecture d'un fichier Excel avec pandas : somme et valeurs uniques |
| `day4.py` | Révision : lecture Excel, somme et produits uniques |
| `day5.py` | Filtrer (montant > 100) et trier les données avec pandas |
| `day6.py` | Nettoyer un fichier Excel : espaces, doublons, cellules vides |
| `day7.py` | Fusionner plusieurs fichiers Excel en un seul avec pandas |
| `day8.py` | Fusionner et nettoyer plusieurs fichiers Excel en un seul |
| `day9.py` | Rapport du total des ventes par produit |
| `day10.py` | Rapport du total des ventes par ville |
| `day11.py` | Mettre en forme un rapport Excel avec openpyxl : en-têtes en couleur et colonnes ajustées |
| `day12.py` | Projet final : fusionner, nettoyer et générer un rapport Excel formaté en 2 feuilles |
| `day13.py` | Rapport Excel formaté : ventes par produit et par ville |

## Avant / Après
Fichier Excel désordonné, puis fichier nettoyé avec `day6.py` :

![Avant](before.png)
![Après](after.png)

## Rapport final
![Rapport final](rapport_final.PNG)


## Outils
- Python 3.11
- pandas et openpyxl

## Feuille de route
- [x] Lire un fichier Excel avec pandas
- [x] Filtrer et trier les données
- [x] Nettoyer les données (doublons, cellules vides, espaces)
- [x] Écrire un fichier Excel
- [x] Fusionner plusieurs fichiers Excel en un seul
- [x] Générer un rapport de ventes

## Exemple de message client
Message type accompagnant un rapport Excel : signaler les lignes à vérifier (produit vide, montant manquant) et confirmer les totaux.

## Contact
Disponible pour des petites missions d'automatisation Excel. Contactez-moi via GitHub.
