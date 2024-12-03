# Prototype de RNPV pour Pégase (PVs seulement)

## Description

Pégase ne permet actuellement que d'exporter un PV sous format CSV. Cela induit qu'à chaque génération d'un PV,
le gestionnaire doit effectuer manuellement un certain nombre d'opérations pour le remettre en forme avant
de l'imprimer pour le jury.

Ce développement est un prototype qui permet de fusionner 1/ un fichier de données CSV exporté de Pégase
et 2/ un fichier modèle XLSX.

## Pré-requis

- Python 3.11

## Image Docker

Une image Docker est disponible. Pour l'utiliser :

```
docker build -t pegase_rnpv .
docker run --rm -p 5000:5000 pegase_rnpv
```

## Installation

```
git clone https://github.com/kaisersly/pegase_rnpv
cd pegase_rnpv
pip install -r requirements.txt
```

## Lancement

```
flask run (en localhost:5000)
ou
flask run -h 0.0.0.0 -p 5001

Puis ouvrir l'URL indiquée
```

## Fonctionnement

Une fois sur la page web, deux options sont disponibles :

1. Lister les champs de fusion : charge un fichier CSV généré par Pégase et liste les balises utilisables dans un fichier modèle XLSX
2. Fusionner les fichiers : fusionne un fichier CSV généré par Pégase et un fichier modèle XLSX

### Fichiers de démonstration

Les fichiers de démonstration (CSV et XLSX) sont disponibles sur le forum du projet PC-SCOL. Chercher le sujet "Prototype Pégase RNPV".

### Balises

Deux types de balises existent :

- GLOBAL : ces balises peuvent être n'importe où dans le tableau SAUF dans la ligne étudiant. Il est possible d'ajouter du texte autour de la balise
  qui sera rendu tel quel après la fusion (par ex. "Type de session : #SESSION_LIBELLE#" sera fusionné en "Type de session : Session 1")
- LIGNE ETUDIANT : ces balises ne peuvent être utilisées QUE sur la ligne étudiant. Il n'est pas possible d'ajouter du texte autour de la balise.
  Si du texte apparaît en plus de la balise, il disparaît lors de la fusion (par ex. "Nom : #NOM#" n'affichera au final que le nom de l'étudiant)

### Format

Le format global du classeur est préservé lors de la fusion. Le format de la ligne étudiant et de la ligne en-dessous sont utilisés pour générer
des lignes paires et impaires en alternant le format.

#### En-tête et pied de page

L'en-tête et le pied de page du modèle ne sont pas conserver lors de la fusion (limitation de la libraire openpyxl utilisée pour lire le fichier modèle
et générer le fichier fusionné).
