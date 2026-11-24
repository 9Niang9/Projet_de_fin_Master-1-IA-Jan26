# Assistant de recherche cosmétique en pharmacie

Prototype académique de recherche locale de produits cosmétiques à partir de la base Open Beauty Facts. Le projet contient un notebook interactif qui propose des produits selon des mots-clés, un pays de vente facultatif et une liste d'ingrédients à éviter.

Ce projet n'est ni un diagnostic, ni une recommandation médicale personnalisée. Les résultats doivent être examinés par un pharmacien ou un professionnel de santé avant toute utilisation.

## Fonctionnalités

- Nettoyage et normalisation des données Open Beauty Facts.
- Recherche textuelle explicable dans le nom, la marque, la catégorie et la composition.
- Filtrage optionnel par pays.
- Exclusion optionnelle d'ingrédients.
- Affichage des allergènes déclarés et de marqueurs indicatifs.
- Utilisation entièrement locale du notebook, sans modèle de langage ni API externe.

## Structure du dépôt

```text
.
├── chatbot_cosmetique_pharmacie.ipynb
│   Notebook principal : chargement du catalogue, recherche et chatbot interactif.
├── prepare_obf.py
│   Script de téléchargement et de nettoyage des données Open Beauty Facts.
└── data_Open Beauty Facts/
    ├── obf_raw.csv.gz              Archive source téléchargée
    ├── obf_raw.csv                 Version CSV brute
    ├── products_clean.csv          Catalogue nettoyé utilisé par le notebook
    ├── products_clean.jsonl        Même catalogue au format JSON Lines
    ├── ingredients_freq.csv        Fréquence des ingrédients normalisés
    └── report.txt                  Statistiques de préparation
```

## Prérequis

- Python 3.10 ou supérieur recommandé.
- Jupyter Notebook ou JupyterLab.
- Une connexion Internet uniquement pour télécharger les données Open Beauty Facts.

## Installation

Depuis la racine du projet :

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pandas requests tqdm jupyter ipykernel
python -m ipykernel install --user --name projet-cosmetique --display-name "Python (Cosmetique)"
```

Sur Linux ou macOS :

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install pandas requests tqdm jupyter ipykernel
python -m ipykernel install --user --name projet-cosmetique --display-name "Python (Cosmetique)"
```

## Préparer les données

Le catalogue déjà nettoyé est inclus dans `data_Open Beauty Facts/`. Pour régénérer les fichiers :

```powershell
python prepare_obf.py
```

Cette commande télécharge plusieurs centaines de mégaoctets depuis Open Beauty Facts. Il est également possible d'utiliser un fichier déjà téléchargé :

```powershell
python prepare_obf.py --input "data_Open Beauty Facts\obf_raw.csv.gz"
```

Pour obtenir uniquement un échantillon limité pendant les tests :

```powershell
python prepare_obf.py --limit 10000
```

Pour filtrer les produits par pays :

```powershell
python prepare_obf.py --country france
```

Le script produit `products_clean.csv`, `products_clean.jsonl`, `ingredients_freq.csv` et `report.txt`.

## Lancer l'assistant

1. Ouvrir `chatbot_cosmetique_pharmacie.ipynb` dans VS Code ou Jupyter.
2. Sélectionner le noyau `Python (Cosmetique)`.
3. Exécuter les cellules dans l'ordre.
4. Exécuter la dernière cellule ou appeler la fonction `chat()` pour démarrer l'assistant.
5. Saisir `aide` pour afficher l'aide ou `quitter` pour terminer.

Exemples de recherches :

```text
crème hydratante visage
nettoyant doux
masque cheveux
```

Les exemples sont fictifs et ne constituent pas des validations d'adéquation ou de sécurité.

## Recherche et sortie

Le score attribué à chaque produit est un score lexical simple. Les correspondances dans le nom et la marque ont plus de poids que celles dans la catégorie ou les ingrédients.

La recherche ne comprend pas les synonymes, les fautes d'orthographe ou la morphologie. Plusieurs mots-clés peuvent être nécessaires.

Pour chaque candidat, le notebook affiche notamment :

- le nom et la marque ;
- la catégorie ;
- les ingrédients disponibles ;
- les allergènes détectés ;
- les marqueurs sensibles détectés ;
- le lien vers la fiche Open Beauty Facts.

## Limites

- Les données peuvent être incomplètes, obsolètes, bruitées ou disponibles dans plusieurs langues.
- Une absence de marqueur ne prouve pas l'absence d'un ingrédient.
- Une composition présente peut elle-même être incomplète ou incorrecte.
- Les marqueurs sensibles sont indicatifs et doivent être validés par un pharmacien.
- Le système ne pose pas de diagnostic et ne prend pas en compte l'historique médical d'une personne.
- Vérifier l'emballage actuel, les mentions réglementaires et les données professionnelles avant tout conseil.

## Source des données

Les données proviennent de [Open Beauty Facts](https://world.openbeautyfacts.org/data) et sont réutilisées sous licence ODbL. Toute redistribution ou publication doit conserver l'attribution de la source et respecter les conditions de la licence.
