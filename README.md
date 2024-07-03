# Aperçu

Cette application est une application web basée sur Flask conçue pour comparer des fichiers Excel contenant des noms et prénoms d'employés afin d'assurer une continuité dans la déclaration des congés. Elle permet de combiner les colonnes 'Nom' et 'Prénom', de récupérer les valeurs uniques de la colonne combinée 'Nom et prénom', de comparer ces valeurs à travers plusieurs fichiers Excel, et de produire un fichier Excel de résultats.

## Fonctionnalités

- **Chargement des fichiers Excel :** Permet de charger plusieurs fichiers Excel.
- **Combinaison des colonnes :** Combine les colonnes 'Nom' et 'Prénom' en 'Nom et prénom' si nécessaire.
- **Valeurs uniques :** Récupère les valeurs uniques de la colonne 'Nom et prénom'.
- **Comparaison des valeurs :** Compare les valeurs à travers plusieurs fichiers Excel.
- **Génération de fichiers Excel :** Génère un fichier Excel contenant les résultats de la comparaison.
- **Interface web :** Interface utilisateur simple et intuitive pour déposer les fichiers Excel et obtenir les résultats.

## Prérequis

- Python 3.x
- Flask
- Pandas
- openpyxl
- Bootstrap

## Documentation

Pour accéder à plus de documentation, cliquez sur le lien suivant: [https://krdto.github.io/outil_comparaison_maladie/](https://krdto.github.io/outil_comparaison_maladie/)

## Installation

1. Clonez le repository :
    ```bash
    git clone https://github.com/Krdto/outil_comparaison_maladie.git
    ```

2. Accédez au répertoire du projet :
    ```bash
    cd outil_comparaison_maladie
    ```

3. Installez les packages Python requis :
    ```bash
    pip install -r requirements.txt
    ```

## Utilisation

1. Exécutez l'application :
    ```bash
    python app.py
    ```

2. Accédez à l'interface web :
    Ouvrez un navigateur web et allez à [http://localhost:5000](http://localhost:5000).

3. Déposez les fichiers Excel :
    - Sélectionnez plusieurs fichiers Excel.
    - Cliquez sur le bouton **"Comparer les Fichiers"**.

4. Téléchargez le fichier de résultats :
    Le fichier Excel généré contenant les résultats de la comparaison sera disponible en téléchargement.

## Structure des fichiers

- `app.py` : Script principal de l'application.
- `templates/index.html` : Modèle HTML pour l'interface web.
- `static/` : Fichiers statiques (par exemple, images, feuilles de style).
- `requirements.txt` : Liste des dépendances Python.
