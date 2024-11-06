from flask import Flask, render_template, request, redirect, url_for, send_file, abort
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)

# Configuration pour limiter la taille des fichiers à 16 Mo
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024  # Limite de 16 Mo
app.config['UPLOAD_FOLDER'] = os.path.join(tempfile.gettempdir(), 'uploads')

# Créer le dossier si nécessaire et s'assurer qu'il est sécurisé
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'], mode=0o700)

ALLOWED_EXTENSIONS = {'ifc', 'xlsx'}

def allowed_file(filename):
    """Vérifie si le fichier a une extension autorisée"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def load_element_types(file):
    element_types = {}
    df = pd.read_excel(file, sheet_name="Element_Types", engine='openpyxl')
    for _, row in df.iterrows():
        ifc_class = row["IFC_Class"]
        element_types[ifc_class] = ifc_class
    return element_types

def load_required_psets_and_params(file):
    required_psets_and_params = {}
    xls = pd.ExcelFile(file, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        if sheet_name != "Element_Types":
            df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
            for _, row in df.iterrows():
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
    if param_type_str == 'string':
        return str
    elif param_type_str == 'int':
        return int
    elif param_type_str == 'float':
        return float
    elif param_type_str == 'number':
        return float
    elif param_type_str == 'bool':
        return bool
    else:
        raise ValueError(f"Unsupported parameter type '{param_type_str}'")

def gray_empty_cells(worksheet):
    gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is None or cell.value == "":
                cell.fill = gray_fill

def get_building_storey(ifc_file, global_id):
    element = ifc_file.by_guid(global_id)
    if element:
        for rel in element.ContainedInStructure:
            if rel.RelatingStructure.is_a("IfcBuildingStorey"):
                return rel.RelatingStructure.Name or ""
            elif rel.RelatingStructure.is_a("IfcBuilding") or rel.RelatingStructure.is_a("IfcSite"):
                return ""
            else:
                return get_building_storey(ifc_file, rel.RelatingStructure)
    return ""

def process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path):
    element_types = load_element_types(excel_file_path)
    required_psets_and_params = load_required_psets_and_params(excel_file_path)
    ifc_file = ifcopenshell.open(ifc_file_path)
    elements = ifc_file.by_type("IfcElement")
    filtered_elements = [e for e in elements if e.is_a() in element_types]
    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)
    pset_sheets = {}
    model_name = os.path.basename(ifc_file_path)

    for ifc_class, psets in required_psets_and_params.items():
        for pset_name in psets:
            if pset_name not in pset_sheets:
                pset_sheets[pset_name] = output_workbook.create_sheet(f"Results_{pset_name}")
                pset_sheets[pset_name].append(["IFC Class", "3D Model", "Element GlobalId", "Pset Name", "Parameter Name", "Floor", "Value", "Status"])

            for element in filtered_elements:
                if element.is_a() == ifc_class:
                    pset = ifcopenshell.util.element.get_psets(element).get(pset_name, None)
                    for param_name, param_type_str in required_psets_and_params[ifc_class][pset_name].items():
                        param_type = str_to_type(param_type_str)
                        if pset is None:
                            pset_sheets[pset_name].append([element.is_a(), model_name, element.GlobalId, pset_name, param_name, get_building_storey(ifc_file, element.GlobalId), "", "PSet missing"])
                        else:
                            value = pset.get(param_name, None)
                            if value is None:
                                pset_sheets[pset_name].append([element.is_a(), model_name, element.GlobalId, pset_name, param_name, get_building_storey(ifc_file, element.GlobalId), "", "Param missing"])
                            else:
                                pset_sheets[pset_name].append([element.is_a(), model_name, element.GlobalId, pset_name, param_name, get_building_storey(ifc_file, element.GlobalId), value, "OK"])

    overview_sheet = output_workbook.create_sheet("Overview")
    overview_sheet.append(["IFC Class", "3D Model", "Pset Name", "Correct Property set", "Correct parameters", "Correct values", "Percentage of Correct values"])

    for sheet_name, sheet in pset_sheets.items():
        ok_params, missing_params, missing_ps, total_params = 0, 0, 0, 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[7] == "OK":
                ok_params += 1
            elif row[7] == "Param missing":
                missing_params += 1
            elif row[7] == "PSet missing":
                missing_ps += 1
            total_params += 1
        correct_ps = total_params - missing_ps
        correct_params = total_params - missing_params
        percentage_correct = (ok_params / total_params) * 100 if total_params > 0 else 0
        overview_sheet.append([row[0], model_name, sheet_name, correct_ps, correct_params, ok_params, f"{percentage_correct:.2f}"])

    output_workbook._sheets.sort(key=lambda ws: ws.title != 'Overview')
    for sheet in output_workbook:
        gray_empty_cells(sheet)
    output_workbook.save(output_file_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        ifc_file = request.files.get('ifc_file')
        excel_file = request.files.get('excel_file')
        if not ifc_file or not excel_file:
            abort(400, "Les fichiers IFC et Excel sont requis.")

        if not (allowed_file(ifc_file.filename) and allowed_file(excel_file.filename)):
            abort(400, "Format de fichier non autorisé.")

        ifc_filename = secure_filename(f"{uuid.uuid4().hex}_{ifc_file.filename}")
        excel_filename = secure_filename(f"{uuid.uuid4().hex}_{excel_file.filename}")
        ifc_file_path = os.path.join(app.config['UPLOAD_FOLDER'], ifc_filename)
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        ifc_file.save(ifc_file_path)
        excel_file.save(excel_file_path)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_file_path = os.path.join(temp_dir, f'output_{os.path.splitext(ifc_filename)[0]}.xlsx')
            process_files(temp_dir, ifc_file_path, excel_file_path, output_file_path)
            return send_file(output_file_path, as_attachment=True, download_name=f'output_{os.path.splitext(ifc_filename)[0]}.xlsx')

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
