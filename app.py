from flask import Flask, request, render_template, flash, redirect, send_file
import pandas as pd
import io

# Initialisation de l'application Flask et configuration du journal de logs
app = Flask(__name__)
app.secret_key = 'infokey3'

def process_files(files):
    """Charge les fichiers Excel et retourne une liste de DataFrames.
    Concatène les colonnes 'Nom' et 'prénom' si nécessaire, puis récupère les valeurs uniques de la colonne 'Nom et prénom' 
    avant de comparer les valeurs entre les fichiers. Crée une dataframe de résultat, sauvegarde le résultat en objet BytesIO, 
    l'écris dans un fichier Excel et retourne le fichier.

    Args:
        files (str): Chemin des fichiers à charger.

    Returns:
        output (BytesIO): objet BytesIO contenant le fichier Excel de résultat.
    """
    # Chargement des données
    dataframes = []
    filenames = []
    for file in files:
        df = pd.read_excel(io.BytesIO(file.read()))
        dataframes.append(df)
        filenames.append(file.filename)

    # Combine 'Nom' et 'Prénom' en 'Nom et prénom'
    for df in dataframes:
        if 'Nom' in df.columns and 'Prénom' in df.columns:
            df['Nom et prénom'] = df['Nom'] + ' ' + df['Prénom']
            df.drop(columns=['Nom', 'Prénom'], inplace=True)

    # Obtenir des valeurs uniques
    unique_values = set()
    for df in dataframes:
        if 'Nom et prénom' in df.columns:
            unique_values.update(df['Nom et prénom'].dropna().unique())

    # Comparer les valeurs à travers les DataFrames
    comparison_results = []
    for value in unique_values:
        presence = [value]
        for df in dataframes:
            if 'Nom et prénom' in df.columns:
                presence.append(value in df['Nom et prénom'].values)
            else:
                presence.append(False)
        comparison_results.append(presence)

    # Créer une DataFrame de comparaison
    columns = ["Nom et prénom"] + [f'Présent dans {filename}' for filename in filenames]
    comparison_df = pd.DataFrame(comparison_results, columns=columns)

    # Sauvegarder le résultat dans un fichier en mémoire
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        comparison_df.to_excel(writer, index=False, sheet_name='Résultats')
    output.seek(0)
    
    return output

@app.route('/', methods=['GET', 'POST'])
def index():
    """Route pour la page d'accueil et pour traiter les fichiers déposés.

    Returns:
        str: Le rendu HTML de la page.
    """
    # Si utilisateur a soumis des fichiers, traiter les fichiers, sinon flash un message d'erreur
    if request.method == 'POST':
        files = request.files.getlist('files[]')
        if not files or len(files) == 0:
            flash('Aucun fichier téléchargé!', 'danger')
            return redirect(request.url)

        try:
            result_file = process_files(files)
            flash('Fichiers traités avec succès!', 'success')
            return send_file(result_file, as_attachment=True, download_name='resultat.xlsx')
        except Exception as e:
            flash(str(e), 'danger')
            return redirect(request.url)
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
