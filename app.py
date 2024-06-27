from flask import Flask, request, render_template, flash, redirect, url_for, send_file
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'results'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

if not os.path.exists(app.config['RESULT_FOLDER']):
    os.makedirs(app.config['RESULT_FOLDER'])

def process_files(file_paths):
    # Loading the data
    dataframes = [pd.read_excel(file) for file in file_paths]

    # Combine 'Nom' and 'Prénom' into 'Nom et prénom'
    for df in dataframes:
        if 'Nom' in df.columns and 'Prénom' in df.columns:
            df['Nom et prénom'] = df['Nom'] + ' ' + df['Prénom']
            df.drop(columns=['Nom', 'Prénom'], inplace=True)

    # Get unique values
    unique_values = set()
    for df in dataframes:
        if 'Nom et prénom' in df.columns:
            unique_values.update(df['Nom et prénom'].dropna().unique())

    # Compare values across DataFrames
    comparison_results = []
    for value in unique_values:
        presence = [value]
        for df in dataframes:
            if 'Nom et prénom' in df.columns:
                presence.append(value in df['Nom et prénom'].values)
            else:
                presence.append(False)
        comparison_results.append(presence)

    # Create comparison DataFrame
    columns = ["Nom et prénom"] + [f'Présent dans {os.path.basename(file)}' for file in file_paths]
    comparison_df = pd.DataFrame(comparison_results, columns=columns)

    # Save the result
    output_file = os.path.join(app.config['RESULT_FOLDER'], 'resultat.xlsx')
    comparison_df.to_excel(output_file, index=False)
    return output_file

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('files[]')
        if not files or len(files) == 0:
            flash('No files uploaded!', 'danger')
            return redirect(request.url)
        
        file_paths = []
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                file.save(filepath)
                file_paths.append(filepath)

        try:
            result_file = process_files(file_paths)
            flash('Files processed successfully!', 'success')
            return send_file(result_file, as_attachment=True)
        except Exception as e:
            flash(str(e), 'danger')
            return redirect(request.url)
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
