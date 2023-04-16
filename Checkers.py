#!/usr/bin/env python3

import ifcopenshell
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from io import BytesIO


def load_element_types(filepath):
    element_types_df = pd.read_excel(filepath, sheet_name='Element_Types', header=None)
    return element_types_df[0].tolist()

ifc_directory = "/Users/bouznira/Desktop/code/python/BIMBot/IFCAnalyser"
ppt_directory = "/Users/bouznira/Desktop/code/python/BIMBot/IFCAnalyser/PPT"

element_types_file = os.path.join(ppt_directory, 'parametres_requis.xlsx')
element_types = load_element_types(element_types_file)

def load_required_psets_and_params(filepath):
    df = pd.read_excel(filepath, sheet_name=None)
    required_psets_and_params = {}
    
    for sheet_name, sheet_data in df.items():
        if sheet_name != 'Element_Types':
            required_psets_and_params[sheet_name] = {}
            for _, row in sheet_data.iterrows():
                ifc_class = row['ifcclass']
                param_name = row['param']
                param_type = eval(row['type'])

                if ifc_class not in required_psets_and_params[sheet_name]:
                    required_psets_and_params[sheet_name][ifc_class] = {}

                required_psets_and_params[sheet_name][ifc_class][param_name] = param_type

    return required_psets_and_params

def gray_empty_cells(excel_filename):
    wb = load_workbook(excel_filename)
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    cell.fill = gray_fill

    wb.save(excel_filename)


required_psets_and_params = load_required_psets_and_params(os.path.join(ppt_directory, 'parametres_requis.xlsx'))

for filename in os.listdir(ifc_directory):
    if filename.endswith(".ifc"):
        file_path = os.path.join(ifc_directory, filename)
        ifc_file = ifcopenshell.open(file_path)

        ifc_name = os.path.splitext(filename)[0]
        excel_filename = f"Résultats_analyse_{ifc_name}.xlsx"

        sheet_data = {}

        for element_type in element_types:
            elements = ifc_file.by_type(element_type)

            for element in elements:
                properties = ifcopenshell.util.element.get_psets(element)

                pset_data = {
                    "Fichier IFC": filename,
                    "Classe Objet IFC 2x3": element.is_a(),
                }

                ifc_class = element.is_a()

                for pset_name, ifc_classes in required_psets_and_params.items():
                    if ifc_class in ifc_classes:
                        params = ifc_classes[ifc_class]
                        pset_present = pset_name in properties

                        for param_name, param_type in params.items():
                            key = f"{param_name}"

                            if not pset_present:
                                pset_data[key] = "PSet non conforme"
                            else:
                                value = properties[pset_name].get(param_name)
                                if value is None:
                                    pset_data[key] = "Paramètres non conforme"
                                elif not isinstance(value, param_type):
                                    pset_data[key] = "Type non conforme"
                                else:
                                    pset_data[key] = "Conforme"

                        if pset_name not in sheet_data:
                            sheet_data[pset_name] = []
                        sheet_data[pset_name].append(pset_data)

        with pd.ExcelWriter(excel_filename) as writer:
            for sheet_name, data in sheet_data.items():
                df = pd.DataFrame(data).groupby("Classe Objet IFC 2x3", group_keys=True).first().reset_index()
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        gray_empty_cells(excel_filename)
