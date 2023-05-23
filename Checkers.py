from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from werkzeug.utils import secure_filename

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

def gray_empty_cells(worksheet):
    gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is None or cell.value == "":
                cell.fill = gray_fill

def get_building_storey(ifc_file, element):
    for rel in element.ContainedInStructure:
        if rel.RelatingStructure.is_a("IfcBuildingStorey"):
            return rel.RelatingStructure
        elif rel.RelatingStructure.is_a("IfcBuilding") or rel.RelatingStructure.is_a("IfcSite"):
            return None
        else:
            return get_building_storey(ifc_file, rel.RelatingStructure)

def process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path):
    element_types = load_element_types(excel_file_path)
    required_psets_and_params = load_required_psets_and_params(excel_file_path)

    # Charger le fichier IFC
    ifc_file = ifcopenshell.open(ifc_file_path)

    # Obtenir la liste des éléments et filtrer par type d'élément
    elements = ifc_file.by_type("IfcElement")
    filtered_elements = [e for e in elements if e.is_a() in element_types]

    # Création d'un classeur Excel de sortie
    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)  # Supprimer la feuille active par défaut

    # Création de la feuille "Overview"
    overview_sheet = output_workbook.create_sheet("Overview")
    overview_sheet.append(["IFC Class", "Pset Name", "Percentage of Correct Properties"])

    for ifc_class, psets in required_psets_and_params.items():
        for pset_name in psets:
            # Créer une nouvelle feuille pour chaque PSet s'il n'existe pas encore
            pset_sheet = output_workbook.create_sheet(f"{ifc_class}_{pset_name}")
            pset_sheet.append(["IFC Class", "Element GlobalId", "Pset Name", "Parameter Name", "Floor", "Value", "Status"])

            num_elements = 0
            num_correct_properties = 0

            for element in filtered_elements:
                if element.is_a() == ifc_class:
                    pset = ifcopenshell.util.element.get_psets(element).get(pset_name, None)

                    # Obtenir l'étage de l'élément
                    floor = get_building_storey(ifc_file, element)
                    floor_name = floor.Name if floor else "N/A"

                    for param_name, param_type_str in required_psets_and_params[ifc_class][pset_name].items():
                        param_type = str_to_type(param_type_str)
                        if pset is None:
                            pset_sheet.append([ifc_class, element.GlobalId, pset_name, param_name, floor_name, "", "PSet missing"])
                        else:
                            value = pset.get(param_name, None)
                            if value is None:
                                pset_sheet.append([ifc_class, element.GlobalId, pset_name, param_name, floor_name, "", "Param missing"])
                            else:
                                # Vérifier si la propriété est correcte
                                expected_type = str_to_type(param_type_str)
                                try:
                                    # Tentative de conversion de la valeur à son type attendu
                                    converted_value = expected_type(value)
                                    pset_sheet.append([ifc_class, element.GlobalId, pset_name, param_name, floor_name, value, "OK"])
                                    num_correct_properties += 1
                                except ValueError:
                                    pset_sheet.append([ifc_class, element.GlobalId, pset_name, param_name, floor_name, value, "Incorrect"])

                    num_elements += 1

            if num_elements > 0:
                percentage_correct = (num_correct_properties / (len(required_psets_and_params[ifc_class][pset_name]) * num_elements)) * 100
            else:
                percentage_correct = 0

            overview_sheet.append([ifc_class, pset_name, f"{percentage_correct:.2f}"])

    # Appliquer la mise en forme
    for sheet in output_workbook:
        gray_empty_cells(sheet)

    # Enregistrer le classeur de sortie
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

            # Extraire le nom de base du fichier IFC sans extension
            ifc_base_name = os.path.splitext(ifc_filename)[0]

            with tempfile.TemporaryDirectory() as temp_dir:
                output_file_path = os.path.join(temp_dir, f'output_{ifc_base_name}.xlsx')
                process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path)
                return send_file(output_file_path, as_attachment=True, download_name=f'output_{ifc_base_name}.xlsx')

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
