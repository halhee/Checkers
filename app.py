from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
from openpyxl import Workbook
from werkzeug.utils import secure_filename
import openai

# Configuration OpenAI
api_key = "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"  # Remplacez par votre clé API OpenAI
openai.api_key = api_key

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def load_element_types(file):
    element_types = {}
    df = pd.read_excel(file, sheet_name="Element_Types", engine='openpyxl')
    for index, row in df.iterrows():
        ifc_class = row["IFC_Class"]
        element_types[ifc_class] = ifc_class
    return element_types

def load_required_psets_and_params(file):
    required_psets_and_params = {}
    xls = pd.ExcelFile(file, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        if sheet_name != "Element_Types":
            df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
            for index, row in df.iterrows():
                ifc_class = row["IFC_Class"]
                param_name = row["Parametre"]
                param_type = row["Type"]
                if ifc_class not in required_psets_and_params:
                    required_psets_and_params[ifc_class] = {}
                if sheet_name not in required_psets_and_params[ifc_class]:
                    required_psets_and_params[ifc_class][sheet_name] = {}
                required_psets_and_params[ifc_class][sheet_name][param_name] = param_type
    return required_psets_and_params

def str_to_type(param_type_str):
    param_type_str = param_type_str.lower()
    if param_type_str.lower() == 'string':
        return str
    elif param_type_str.lower() == 'int':
        return int
    elif param_type_str.lower() == 'float':
        return float
    elif param_type_str == 'number':
        return float
    elif param_type_str.lower() == 'bool':
        return bool
    else:
        raise ValueError(f"Unsupported parameter type '{param_type_str}'")

def process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path):
    element_types = load_element_types(excel_file_path)
    required_psets_and_params = load_required_psets_and_params(excel_file_path)

    # Load the IFC file
    ifc_file = ifcopenshell.open(ifc_file_path)

    # Get the list of elements and filter by element type
    elements = ifc_file.by_type("IfcElement")
    filtered_elements = [e for e in elements if e.is_a() in element_types]

    # Creating an output Excel workbook
    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)  # Remove the default active sheet

    # Process elements and write results to sheets
    pset_sheets = {}
    for ifc_class, psets in required_psets_and_params.items():
        for pset_name, params in psets.items():
            if pset_name not in pset_sheets:
                pset_sheets[pset_name] = output_workbook.create_sheet(f"Results_{pset_name}")
                pset_sheets[pset_name].append(["IFC Class", "Element GlobalId", "Pset Name", "Parameter Name", "Value", "Status"])
            for element in filtered_elements:
                if element.is_a() == ifc_class:
                    pset = ifcopenshell.util.element.get_psets(element).get(pset_name, None)
                    for param_name, param_type in params.items():
                        value = pset.get(param_name) if pset else None
                        status = "OK" if value else "Missing"
                        pset_sheets[pset_name].append([ifc_class, element.GlobalId, pset_name, param_name, value, status])

    # Save the output workbook
    output_workbook.save(output_file_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        ifc_file = request.files['ifc_file']
        excel_file = request.files['excel_file']
        if ifc_file and excel_file:
            ifc_filename = secure_filename(ifc_file.filename)
            excel_filename = secure_filename(excel_file.filename)
            ifc_file_path = os.path.join(app.config['UPLOAD_FOLDER'], ifc_filename)
            excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            ifc_file.save(ifc_file_path)
            excel_file.save(excel_file_path)

            with tempfile.TemporaryDirectory() as temp_dir:
                output_file_path = os.path.join(temp_dir, f'output_{os.path.splitext(ifc_filename)[0]}.xlsx')
                process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path)

                # Convert Excel to CSV for question-answering
                csv_path = os.path.join(temp_dir, f'output_{os.path.splitext(ifc_filename)[0]}.csv')
                pd.read_excel(output_file_path, engine='openpyxl').to_csv(csv_path, index=False)

                return send_file(output_file_path, as_attachment=True, download_name=f'output_{os.path.splitext(ifc_filename)[0]}.xlsx')

    return redirect(url_for('index'))

@app.route('/ask-results', methods=['POST'])
def ask_results():
    try:
        question = request.json.get('question', '').strip()
        if not question:
            return jsonify({'error': 'Aucune question fournie.'}), 400

        # Chemin du fichier Excel généré
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'resultats.xlsx')
        if not os.path.exists(excel_path):
            return jsonify({'error': 'Fichier Excel introuvable. Veuillez d\'abord effectuer une analyse.'}), 400

        # Charger toutes les feuilles de l'Excel
        sheets = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl')

        # Consolider les données des onglets contenant la colonne "Status"
        consolidated_data = pd.DataFrame()
        for sheet_name, sheet_data in sheets.items():
            if 'Status' in sheet_data.columns:
                consolidated_data = pd.concat([consolidated_data, sheet_data], ignore_index=True)

        # Vérifier si la colonne "Status" est trouvée après consolidation
        if consolidated_data.empty or 'Status' not in consolidated_data.columns:
            return jsonify({'error': 'La colonne "Status" est introuvable dans les onglets analysés.'}), 400

        # Calculer des statistiques globales
        total_elements = len(consolidated_data)
        missing_params = consolidated_data[consolidated_data['Status'] == 'Param missing'].shape[0]
        missing_psets = consolidated_data[consolidated_data['Status'] == 'PSet missing'].shape[0]
        correct_params = consolidated_data[consolidated_data['Status'] == 'OK'].shape[0]

        summary = (
            f"Total des éléments analysés : {total_elements}\n"
            f"Paramètres manquants : {missing_params}\n"
            f"Psets manquants : {missing_psets}\n"
            f"Paramètres corrects : {correct_params}\n"
        )

        # Construire le contexte pour OpenAI
        context = (
            f"Voici un résumé des résultats consolidés :\n\n{summary}\n\n"
            f"Extrait des données :\n{consolidated_data.head(10).to_string(index=False)}\n\n"
            f"Question : {question}\n\n"
            f"Répondez en fonction des données ci-dessus."
        )

        # Interroger OpenAI
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "Tu es un assistant expert en analyse de données BIM. Utilise les données fournies pour répondre précisément à la question posée."},
                {"role": "user", "content": context}
            ],
            max_tokens=300
        )
        ai_response = response['choices'][0]['message']['content']

        # Retourner la réponse
        return jsonify({'answer': ai_response})

    except Exception as e:
        print(f"Erreur lors de l'interrogation des résultats : {e}")
        return jsonify({'error': f"Erreur : {e}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
