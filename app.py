import pandas as pd
import os
from flask import Flask, request, redirect, url_for, send_from_directory, render_template, flash

app = Flask(__name__)
app.secret_key = 'infokey2'
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    """Vérifie si le fichier à une extension autorisée."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def load_data(file_paths):
    """Charger les données des fichiers Excel dans des DataFrames et concaténer si nécessaire les colonnes 'Nom' et 'Prénom' en 'Nom et Prénom'."""
    dataframes = []
    for file in file_paths:
        df = pd.read_excel(file)
        print(f"Loaded DataFrame from {file}:\n{df.head()}")
        # Vérifie si les colonnes 'Nom' et 'Prénom' existent
        if 'Nom' in df.columns and 'Prénom' in df.columns:
            # Concatène les colonnes 'Nom' et 'Prénom'
            df['Nom et Prénom'] = df['Nom'].str.strip() + ' ' + df['Prénom'].str.strip()
            # Supprime les colonnes 'Nom' et 'Prénom'
            df = df.drop(['Nom', 'Prénom'], axis=1)
        dataframes.append(df)
    return dataframes

def get_unique_values(dataframes, column_name):
    """Élimination des doublons et récupération des valeurs uniques de 'Nom et Prénom'."""
    unique_values = set()
    for df in dataframes:
        # Vérification d'existence de la colonne avant extraction
        if column_name in df.columns:
            unique_values.update(df[column_name].dropna().unique())
    return unique_values

def compare_values(dataframes, unique_values, column_name):
    """Comparaison des valeurs uniques de 'Nom et Prénom' à travers les DataFrames et enregistrement de leur présence."""
    comparison_results = []
    for value in unique_values:
      # On initialise une liste presence avec la valeur unique actuelle.
         # Cette liste enregistrera si la valeur est présente dans chaque DataFrame.
        presence = [value]
        for df in dataframes:
            # Pour chaque DataFrame, on vérifie si la colonne column_name existe
            if column_name in df.columns:
              # Si la colonne existe, on vérifie si value est présente dans la colonne et on ajoute True ou False à la liste
                presence.append(value in df[column_name].values)
            else:
                # on ajoute False à la liste presence pour indiquer que la colonne n'est pas présente dans la DataFrame
                presence.append(False)
        # On ajoute la liste presence (contenant la valeur unique et sa présence dans chaque DataFrame) à la liste comparison_results
        comparison_results.append(presence)
    return comparison_results

def create_comparison_df(comparison_results, file_names):
    """Création d'un DataFrame contenant les résultats de la comparaison."""
    columns = ['Nom et Prénom'] + [f'Présent dans {os.path.basename(file)}' for file in file_names]
    comparison_df = pd.DataFrame(comparison_results, columns=columns)
    return comparison_df

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        files = request.files.getlist('files[]')
        file_paths = []
        file_names = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = file.filename
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                file_paths.append(filepath)
                file_names.append(filename)
            else:
                flash('Seuls les fichiers Excel sont autorisés', 'danger')
                return redirect(request.url)

        if len(file_paths) < 2:
            flash('Veuillez déposer au moins deux fichiers', 'danger')
            return redirect(request.url)

        dataframes = load_data(file_paths)
        unique_values = get_unique_values(dataframes, 'Nom et Prénom')
        comparison_results = compare_values(dataframes, unique_values, 'Nom et Prénom')
        comparison_df = create_comparison_df(comparison_results, file_names)

        output_file = 'résultat_comparaison.xlsx'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_file)
        comparison_df.to_excel(output_path, index=False)

        # Delete uploaded files after processing
        for file_path in file_paths:
            os.remove(file_path)

        return redirect(url_for('download_file', filename=output_file))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """Route pour le téléchargement du fichier."""
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        flash('Fichier introuvable', 'danger')
        return redirect(url_for('upload_files'))
    except Exception as e:
        flash(f"Erreur lors du téléchargement du fichier: {e}", 'danger')
        return redirect(url_for('upload_files'))

if __name__ == "__main__":
    app.run(debug=True)
