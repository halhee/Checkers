from flask import Flask, render_template, request, redirect, url_for, send_file, abort, jsonify
import os
import tempfile
import ifcopenshell
import ifcopenshell.util.element
import uuid
import time
import pandas as pd
import numpy as np
from collections import defaultdict, Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter
from typing import Dict, List, Set, Tuple
import shutil
import re
import zipfile
from datetime import datetime
import threading
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

import carbon
import site_location
import corps_etat
import history
from excel_safety import excel_safe
from agents.recommendations import run_recommendations_pipeline, create_recommendations_sheet
from agents.lot_matching import match_lot_with_ai

load_dotenv()

app = Flask(__name__)


@app.after_request
def _set_security_headers(response):
    """En-têtes de sécurité standards, défense en profondeur -- n'affectent
    pas le fonctionnement de l'app (pas d'iframe, pas de contenu tiers à
    sniffer). N'importe quoi d'uploadé (IFC, Excel) n'est jamais servi
    directement comme une page HTML par Flask, mais ça coûte rien de le
    garantir explicitement plutôt que de compter sur ce fait implicite."""
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'DENY')
    response.headers.setdefault('Referrer-Policy', 'same-origin')
    return response


# Configuration
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024  # 300 Mo
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
TEMP_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')

# Créer les dossiers nécessaires. mode=0o750 (propriétaire : lecture/écriture/
# exécution ; groupe : lecture/exécution ; autres : rien) plutôt que le 0o777
# précédent -- ces dossiers contiennent les fichiers IFC/Excel uploadés et les
# rapports générés (potentiellement des données de projet sensibles), inutile
# qu'ils soient lisibles/écrivables par n'importe quel autre compte du poste
# en dehors de Docker (où un seul utilisateur non-root tourne de toute façon,
# voir Dockerfile). os.makedirs applique encore l'umask du process par-dessus,
# donc le chmod explicite ci-dessous garantit le résultat indépendamment de lui.
print(f"Creating folders: {UPLOAD_FOLDER}, {TEMP_FOLDER}")
for folder in [UPLOAD_FOLDER, TEMP_FOLDER]:
    if not os.path.exists(folder):
        try:
            os.makedirs(folder, mode=0o750, exist_ok=True)
            os.chmod(folder, 0o750)
            print(f"Created folder: {folder}")
        except Exception as e:
            print(f"Error creating folder {folder}: {str(e)}")

history.init_db()

ALLOWED_EXTENSIONS = {'ifc', 'xlsx'}

# Identifiant d'analyse = uuid4 utilisé à la fois comme nom de dossier sur
# disque (TEMP_FOLDER/<id>/...) et comme clé SQLite -- validé dans toutes les
# routes qui le prennent en paramètre d'URL pour écarter toute tentative de
# traversée de chemin (ex. "../../etc/passwd") avant de s'en servir.
_ANALYSIS_ID_RE = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')


def _valid_analysis_id(analysis_id: str) -> bool:
    return bool(_ANALYSIS_ID_RE.match(analysis_id))

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Un .xlsx est une archive ZIP de XML -- Python ne limite pas la taille
# décompressée par défaut ("zip bomb" : un petit fichier qui explose en
# mémoire à la décompression ; démontré en test manuel : 500 Mo de données
# compressent à moins de 500 Ko, largement sous la limite d'upload de 300 Mo).
# Seuils volontairement larges pour ne jamais rejeter un vrai classeur Excel
# (le plus gros fichier de règles de ce projet fait quelques dizaines de Ko) :
# un vrai .xlsx compresse rarement plus de 20-50x (XML répétitif), une bombe
# zip vise des ratios de plusieurs centaines à plusieurs milliers.
_XLSX_MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
_XLSX_MAX_COMPRESSION_RATIO = 100


def _reject_if_zip_bomb(file_path: str):
    """Lève ValueError si `file_path` n'est pas une archive ZIP valide, ou si
    sa taille décompressée totale / le ratio de compression d'un de ses
    membres dépasse un seuil raisonnable. À appeler avant tout parsing
    openpyxl d'un fichier Excel uploadé."""
    try:
        with zipfile.ZipFile(file_path) as zf:
            total_uncompressed = 0
            for info in zf.infolist():
                total_uncompressed += info.file_size
                if info.compress_size > 0 and info.file_size / info.compress_size > _XLSX_MAX_COMPRESSION_RATIO:
                    raise ValueError("Fichier Excel suspect (taux de compression anormal)")
                if total_uncompressed > _XLSX_MAX_UNCOMPRESSED_BYTES:
                    raise ValueError("Fichier Excel suspect (taille décompressée excessive)")
    except zipfile.BadZipFile:
        raise ValueError("Fichier Excel invalide (pas une archive .xlsx valide)")

def _row_matches_corps_etat(value, code) -> bool:
    """True si la cellule "Corps_Etat" d'une ligne inclut `code`. Une valeur
    vide/absente s'applique à tous les corps d'état -- l'absence de
    restriction n'en est pas une, ça permet de ne remplir cette colonne
    que là où une règle est vraiment propre à une discipline. La cellule
    peut lister plusieurs codes séparés par une virgule ou un point-virgule
    (ex. "ARC, STR") pour une règle partagée entre plusieurs corps d'état."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    codes = [c.strip().upper() for c in str(value).replace(";", ",").split(",") if c.strip()]
    return not codes or code.upper() in codes

def load_element_types(file, code=None):
    """On raisonne par PSet, pas par corps d'état -- un seul onglet
    "Element_Types" comme avant, avec une colonne "Corps_Etat" optionnelle
    sur chaque ligne (voir _row_matches_corps_etat) plutôt que des onglets
    dupliqués par discipline. Sans code (défaut) ou si la colonne est
    absente, toutes les lignes sont retenues : comportement historique
    inchangé pour les fichiers à un seul jeu de règles."""
    element_types = {}
    df = pd.read_excel(file, sheet_name="Element_Types", engine='openpyxl')
    has_corps_etat_col = "Corps_Etat" in df.columns
    for index, row in df.iterrows():
        if code and has_corps_etat_col and not _row_matches_corps_etat(row.get("Corps_Etat"), code):
            continue
        ifc_class = row["IFC_Class"]
        element_types[ifc_class] = ifc_class
    return element_types

def load_required_psets_and_params(file, code=None):
    """Voir load_element_types pour le rôle de `code` et de la colonne
    "Corps_Etat" optionnelle. Les onglets de PSet gardent leur nom réel
    (ex. "BaseQuantities", "Pset_WallCommon") -- pas de suffixe, puisqu'il
    n'y a plus qu'un seul onglet par PSet, partagé par tous les corps
    d'état et filtré ligne par ligne."""
    required_psets_and_params = {}
    xls = pd.ExcelFile(file, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        if sheet_name in ("Element_Types", "Corps_Etat"):
            continue
        df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
        if not {"IFC_Class", "Parametre", "Type"}.issubset(df.columns):
            # Pas un onglet de règles PSet (ex. une note/README ajoutée par
            # l'utilisateur) -- ignoré plutôt que de faire planter l'analyse.
            print(f"Onglet '{sheet_name}' ignoré (ne ressemble pas à un onglet de PSet).")
            continue
        has_corps_etat_col = "Corps_Etat" in df.columns
        for index, row in df.iterrows():
            if code and has_corps_etat_col and not _row_matches_corps_etat(row.get("Corps_Etat"), code):
                continue
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
    try:
        if element is not None:
            if hasattr(element, 'ContainedInStructure') and element.ContainedInStructure:
                for rel in element.ContainedInStructure:
                    if rel.RelatingStructure.is_a("IfcBuildingStorey"):
                        return rel.RelatingStructure.Name if rel.RelatingStructure.Name else "Sans étage"
                    elif rel.RelatingStructure.is_a("IfcBuilding") or rel.RelatingStructure.is_a("IfcSite"):
                        return "Sans étage"
                    else:
                        return get_building_storey(ifc_file, rel.RelatingStructure)
        return "Sans étage"
    except Exception as e:
        print(f"Error getting building storey: {str(e)}")
        return "Sans étage"

def sort_floor_name(floor_name):
    if floor_name == "Sans étage":
        return "ZZZ"
    elif floor_name == "Sous-sol":
        return "AAA_-1"
    elif floor_name.startswith("R+"):
        try:
            return f"CCC_{int(floor_name[2:]):02d}"
        except ValueError:
            return f"DDD_{floor_name}"
    elif floor_name == "Rez-de-chaussée":
        return "BBB_00"
    elif floor_name.startswith("-"):
        try:
            return f"AAA_{int(floor_name.split('.')[0]):02d}"
        except ValueError:
            return f"DDD_{floor_name}"
    else:
        return f"DDD_{floor_name}"

def create_summary_sheet(workbook, total_elements, valid_elements, missing_elements, missing_psets, missing_params, floor_stats, elements_by_class):
    # Créer l'onglet de résumé
    summary = workbook.create_sheet("Résumé", 0)
    
    # Styles
    title_font = Font(size=14, bold=True, color="FFFFFF")
    header_font = Font(size=12, bold=True)
    normal_font = Font(size=11)
    
    title_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ko_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    percent_format = '0.0%'
    
    def apply_style_to_range(start_cell, end_cell, font=None, fill=None, alignment=None, border=None, value=None, number_format=None):
        start_col = ord(start_cell[0].upper()) - ord('A') + 1
        end_col = ord(end_cell[0].upper()) - ord('A') + 1
        start_row = int(start_cell[1:])
        end_row = int(end_cell[1:])
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = summary.cell(row=row, column=col)
                if font:
                    cell.font = font
                if fill:
                    cell.fill = fill
                if alignment:
                    cell.alignment = alignment
                if border:
                    cell.border = border
                if number_format:
                    cell.number_format = number_format
                if value and col == start_col and row == start_row:
                    cell.value = excel_safe(value)
    
    # Titre principal
    apply_style_to_range('A1', 'E1', font=title_font, fill=title_fill, alignment=center_alignment, border=border, value="Rapport de validation IFC")
    summary.merge_cells('A1:E1')
    
    # Statistiques globales
    apply_style_to_range('A3', 'A3', font=header_font, fill=header_fill, value="Statistiques globales")
    
    stats = [
        ("Total d'éléments", total_elements),
        ("Éléments valides", valid_elements, valid_elements/total_elements if total_elements > 0 else 0),
        ("Éléments invalides", missing_elements, missing_elements/total_elements if total_elements > 0 else 0),
        ("PSet manquants", missing_psets),
        ("Paramètres manquants", missing_params)
    ]
    
    for i, stat in enumerate(stats, 4):
        label = stat[0]
        value = stat[1]
        percent = stat[2] if len(stat) > 2 else None
        
        apply_style_to_range(f'A{i}', f'A{i}', font=normal_font, alignment=left_alignment, border=border, value=label)
        apply_style_to_range(f'B{i}', f'B{i}', font=normal_font, alignment=center_alignment, border=border, value=value)
        if percent is not None:
            apply_style_to_range(f'C{i}', f'C{i}', font=normal_font, alignment=center_alignment, border=border, value=percent, number_format=percent_format)
    
    # Statistiques par étage
    apply_style_to_range('A7', 'A7', font=header_font, fill=header_fill, value="Répartition par étage")
    
    headers = ["Étage", "Total", "Valides", "Invalides", "Taux de validité"]
    for col, header in enumerate(headers, 1):
        apply_style_to_range(f'{chr(ord("A")+col-1)}8', f'{chr(ord("A")+col-1)}8', 
                           font=header_font, fill=header_fill, alignment=center_alignment, border=border, value=header)
    
    row = 9
    sorted_floors = sorted(floor_stats.items(), key=lambda x: sort_floor_name(x[0]))
    for floor, stats in sorted_floors:
        total = stats["valid"] + stats["invalid"]
        valid_rate = stats["valid"] / total if total > 0 else 0
        
        cells = [
            (floor, left_alignment),
            (total, center_alignment),
            (stats["valid"], center_alignment),
            (stats["invalid"], center_alignment),
            (valid_rate, center_alignment)
        ]
        
        for col, (value, alignment) in enumerate(cells, 1):
            cell_ref = f'{chr(ord("A")+col-1)}{row}'
            fill = None
            number_format = None
            
            if col == 5:  # Dernière colonne (taux)
                number_format = percent_format
                if valid_rate >= 0.8:
                    fill = ok_fill
                elif valid_rate < 0.5:
                    fill = ko_fill
            
            apply_style_to_range(cell_ref, cell_ref, 
                               font=normal_font, 
                               alignment=alignment, 
                               border=border, 
                               fill=fill, 
                               value=value,
                               number_format=number_format)
        
        row += 1
    
    # Statistiques par type
    type_title_row = row + 2
    apply_style_to_range(f'A{type_title_row}', f'A{type_title_row}', font=header_font, fill=header_fill, value="Répartition par type")
    
    headers = ["Type", "Total", "Valides", "Invalides", "Taux de validité"]
    header_row = type_title_row + 1
    for col, header in enumerate(headers, 1):
        apply_style_to_range(f'{chr(ord("A")+col-1)}{header_row}', f'{chr(ord("A")+col-1)}{header_row}',
                           font=header_font, fill=header_fill, alignment=center_alignment, border=border, value=header)
    
    row = header_row + 1
    sorted_types = sorted(elements_by_class.items())
    for element_type, stats in sorted_types:
        total = stats["total"]
        valid_rate = stats["valid"] / total if total > 0 else 0
        
        cells = [
            (element_type, left_alignment),
            (total, center_alignment),
            (stats["valid"], center_alignment),
            (stats["invalid"], center_alignment),
            (valid_rate, center_alignment)
        ]
        
        for col, (value, alignment) in enumerate(cells, 1):
            cell_ref = f'{chr(ord("A")+col-1)}{row}'
            fill = None
            number_format = None
            
            if col == 5:  # Dernière colonne (taux)
                number_format = percent_format
                if valid_rate >= 0.8:
                    fill = ok_fill
                elif valid_rate < 0.5:
                    fill = ko_fill
            
            apply_style_to_range(cell_ref, cell_ref,
                               font=normal_font,
                               alignment=alignment,
                               border=border,
                               fill=fill,
                               value=value,
                               number_format=number_format)
        
        row += 1
    
    # Ajuster la largeur des colonnes
    for col in range(1, 6):
        letter = get_column_letter(col)
        max_length = 0
        for cell in summary[letter]:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        summary.column_dimensions[letter].width = adjusted_width
    
    # Figer les volets
    summary.freeze_panes = 'A2'
    
    # Activer l'onglet de résumé
    workbook.active = workbook.worksheets.index(summary)


def create_lot_comparison_sheet(workbook, elements_by_class_by_lot, carbon_facts, lot_descriptions):
    """Nouvel onglet Excel "Comparatif Lots" : une ligne par lot (discipline),
    avec sa conformité PSet et son empreinte carbone -- vue de synthèse pour
    comparer où en est chaque lot, sans détection de clash spatial (décision
    utilisateur : seulement des statistiques agrégées et comparées).
    `lot_descriptions` (voir corps_etat.py) indique, par lot, quel corps
    d'état a été retenu et comment (correspondance exacte/floue/IA), ou
    qu'aucune correspondance n'a été trouvée."""
    sheet = workbook.create_sheet("Comparatif Lots")

    title_font = Font(size=14, bold=True, color="FFFFFF")
    header_font = Font(size=12, bold=True)
    normal_font = Font(size=11)
    title_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ko_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    percent_format = '0.0%'

    sheet.merge_cells('A1:H1')
    cell = sheet.cell(row=1, column=1)
    cell.value = "COMPARATIF PAR LOT"
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = center_alignment
    cell.border = border
    sheet.row_dimensions[1].height = 30

    headers = ["Lot", "Corps d'état détecté", "Total éléments", "Valides", "Invalides", "Taux de validité",
               "Empreinte carbone (kg CO2e)", "% du total carbone"]
    for col, header in enumerate(headers, 1):
        cell = sheet.cell(row=3, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border

    total_carbon = carbon_facts["total"]
    by_lot_carbon = carbon_facts.get("by_lot", {})

    row = 4
    for lot_name, classes in elements_by_class_by_lot.items():
        has_rules = bool(classes)
        total = sum(s["total"] for s in classes.values())
        valid = sum(s["valid"] for s in classes.values())
        invalid = total - valid
        valid_rate = valid / total if total else 0
        carbon_kg = by_lot_carbon.get(lot_name, 0)
        carbon_pct = (carbon_kg / total_carbon * 100) if total_carbon else 0
        rate_fill = ok_fill if valid_rate >= 0.8 else (ko_fill if valid_rate < 0.5 else None)
        description = lot_descriptions.get(lot_name, "")

        cells = [
            (lot_name, left_alignment, None, None),
            (description, left_alignment, None, None if has_rules else ko_fill),
            (total if has_rules else "N/A", center_alignment, None, None),
            (valid if has_rules else "N/A", center_alignment, None, None),
            (invalid if has_rules else "N/A", center_alignment, None, None),
            (valid_rate if has_rules else "N/A", center_alignment, percent_format if has_rules else None, rate_fill if has_rules else None),
            (f"{carbon_kg:.2f}", center_alignment, None, None),
            (f"{carbon_pct:.1f}%", center_alignment, None, None),
        ]
        for col, (value, alignment, number_format, fill) in enumerate(cells, 1):
            cell = sheet.cell(row=row, column=col)
            cell.value = excel_safe(value)
            cell.font = normal_font
            cell.alignment = alignment
            cell.border = border
            if number_format:
                cell.number_format = number_format
            if fill:
                cell.fill = fill
        row += 1

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        max_length = max(
            (len(str(sheet.cell(row=r, column=col).value or "")) for r in range(1, row)),
            default=10,
        )
        sheet.column_dimensions[letter].width = max_length + 2

    sheet.freeze_panes = 'A4'


# Seuils de verdict "recette" sur le taux de validité d'un lot -- indicatifs,
# à ajuster selon l'exigence réelle du projet/du marché (pas une norme
# officielle). >= VERDICT_OK : validé sans réserve ; entre les deux : validé
# avec réserves (anomalies à solder, mais pas bloquant) ; en dessous : rejeté.
VERDICT_THRESHOLD_OK = 0.98
VERDICT_THRESHOLD_RESERVE = 0.90


def _lot_verdict(valid_rate: float) -> str:
    if valid_rate >= VERDICT_THRESHOLD_OK:
        return "Validé"
    if valid_rate >= VERDICT_THRESHOLD_RESERVE:
        return "Validé avec réserves"
    return "Rejeté"


def create_recette_sheet(workbook, label: str, timestamp_str: str, excel_filename: str,
                          lots_summary: list, anomalies_raw: list):
    """Onglet Excel "Fiche de Recette" : le livrable attendu d'une recette
    fonctionnelle (verdict par lot, liste des anomalies ouvertes regroupées
    par étage/type/PSet, zone de signature) -- mêmes faits que les autres
    onglets, mis en forme comme un document de recette que l'on peut
    imprimer/faire signer tel quel, plutôt qu'un export de données brutes."""
    sheet = workbook.create_sheet("Fiche de Recette")

    title_font = Font(size=14, bold=True, color="FFFFFF")
    section_font = Font(size=12, bold=True)
    header_font = Font(size=11, bold=True)
    normal_font = Font(size=11)
    italic_font = Font(size=10, italic=True, color="777777")
    title_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    reserve_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    ko_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    n_cols = 5

    def merged_title(row, text, font, fill, height=None):
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        cell = sheet.cell(row=row, column=1)
        cell.value = excel_safe(text)
        cell.font = font
        if fill:
            cell.fill = fill
        cell.alignment = left if font is not title_font else center
        if height:
            sheet.row_dimensions[row].height = height
        return row + 1

    r = 1
    r = merged_title(r, "FICHE DE RECETTE FONCTIONNELLE", title_font, title_fill, height=28)
    r += 1
    r = merged_title(r, f"Projet / repère : {label}", normal_font, None)
    r = merged_title(r, f"Date d'analyse : {timestamp_str}", normal_font, None)
    r = merged_title(r, f"Référentiel de règles (Excel) : {excel_filename}", normal_font, None)
    r += 1

    r = merged_title(r, "1. Verdict par lot (corps d'état)", section_font, section_fill)
    headers = ["Lot", "Corps d'état détecté", "Taux de validité", "Verdict"]
    for col, h in enumerate(headers, 1):
        cell = sheet.cell(row=r, column=col)
        cell.value = h
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    r += 1
    for lot in lots_summary:
        total = lot["total_elements"]
        valid = lot["valid_elements"]
        has_rules = lot["has_rules"]
        if has_rules and total:
            valid_rate = valid / total
            verdict = _lot_verdict(valid_rate)
            fill = {"Validé": ok_fill, "Validé avec réserves": reserve_fill, "Rejeté": ko_fill}[verdict]
            rate_display, rate_format = valid_rate, '0.0%'
        else:
            verdict = "Bilan carbone seul (pas de contrôle PSet)"
            fill = None
            rate_display, rate_format = "N/A", None
        row_cells = [
            (lot["name"], left, None, None),
            (lot["corps_etat"], left, None, None),
            (rate_display, center, rate_format, None),
            (verdict, center, None, fill),
        ]
        for col, (value, alignment, number_format, cell_fill) in enumerate(row_cells, 1):
            cell = sheet.cell(row=r, column=col)
            cell.value = excel_safe(value)
            cell.font = normal_font
            cell.alignment = alignment
            cell.border = border
            if number_format:
                cell.number_format = number_format
            if cell_fill:
                cell.fill = cell_fill
        r += 1
    r += 1

    r = merged_title(
        r, f"2. Anomalies ouvertes ({len(anomalies_raw)} -- regroupées par lot/étage/type/PSet, "
           f"pas par élément individuel)", section_font, section_fill,
    )
    headers = ["Lot", "Étage", "Classe IFC", "PSet", "Nb éléments concernés"]
    for col, h in enumerate(headers, 1):
        cell = sheet.cell(row=r, column=col)
        cell.value = h
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    r += 1
    for a in sorted(anomalies_raw, key=lambda a: -a["count"]):
        row_cells = [a["lot"], a["floor"], a["ifc_class"], a["pset"], a["count"]]
        for col, value in enumerate(row_cells, 1):
            cell = sheet.cell(row=r, column=col)
            cell.value = excel_safe(value)
            cell.font = normal_font
            cell.alignment = left if col <= 4 else center
            cell.border = border
        r += 1
    if not anomalies_raw:
        sheet.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
        sheet.cell(row=r, column=1, value="Aucune anomalie ouverte.").font = normal_font
        r += 1
    r += 1

    r = merged_title(r, "3. Validation", section_font, section_fill)
    r = merged_title(r, "Testé par : ________________________        Date : ____________        Visa : ____________", normal_font, None, height=22)
    r = merged_title(r, "Observations : ___________________________________________________________________", normal_font, None, height=22)
    r += 1
    r = merged_title(
        r,
        "Seuils de verdict indicatifs (ajustables selon l'exigence du projet) : "
        f"≥ {VERDICT_THRESHOLD_OK:.0%} validé -- ≥ {VERDICT_THRESHOLD_RESERVE:.0%} validé avec réserves -- en dessous rejeté.",
        italic_font, None,
    )

    sheet.column_dimensions['A'].width = 22
    sheet.column_dimensions['B'].width = 30
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 30
    sheet.column_dimensions['E'].width = 22


def _write_detail_row(details, row: int, cells, col_widths: list):
    """Écrit une ligne de l'onglet Détails. Applique le style via un
    `NamedStyle` déjà enregistré sur le classeur (`cell.style = <nom>`) plutôt
    que d'assigner font/border/alignment/fill séparément sur chaque cellule --
    openpyxl re-hashe l'objet de style à chaque assignation de propriété pour
    dédupliquer les styles au niveau du classeur, ce qui domine largement le
    temps de génération sur les gros modèles (profilé : >50% du temps total
    pour un fichier de ~2300 éléments). Passer par un style nommé unique par
    cellule ramène ça à un lookup par nom, nettement moins coûteux.
    Calcule aussi la largeur de colonne au passage pour éviter une seconde
    lecture complète de la feuille en fin de génération (voir process_files)."""
    for col, value, style_name in cells:
        cell = details.cell(row=row, column=col)
        cell.value = excel_safe(value)
        cell.style = style_name
        length = len(str(value)) if value is not None else 0
        if length > col_widths[col - 1]:
            col_widths[col - 1] = length


def process_files(temp_dir: str, lots: list, excel_file_path: str, output_file_path: str, label: str = None):
    """`lots` = liste de `(lot_name, ifc_file_path, explicit_corps_etat_code)`,
    au moins un élément. `explicit_corps_etat_code` est None si l'utilisateur
    a laissé "Auto" côté dashboard -- la correspondance automatique (exact,
    puis floue, puis IA) prend alors le relais, exactement comme avant
    l'ajout du sélecteur explicite. Un seul fichier Excel de paramètres
    requis, partagé par tous les lots. `label` est le repère libre affiché
    dans l'historique et la Fiche de Recette (voir history.py) -- par défaut
    le nom du fichier Excel si l'utilisateur n'en a pas saisi."""
    start_time = time.time()
    print(f"Starting analysis...")
    analysis_timestamp = datetime.now()
    excel_filename = os.path.basename(excel_file_path)
    label = label or excel_filename

    # Chargement des règles : un seul jeu partagé par tous les lots (mode
    # historique), ou un jeu par lot si l'Excel déclare un onglet
    # "Corps_Etat" -- voir corps_etat.py. Sans onglet "Corps_Etat", le
    # fichier est traité exactement comme avant ce chantier.
    corps_etat_registry = corps_etat.load_registry(excel_file_path)
    # {lot_name: (element_types | None, required_psets_and_params | None, description)}
    # element_types=None signifie "aucune règle pour ce lot" : bilan carbone
    # seul, sans contrôle PSet (décision utilisateur -- jamais bloquant).
    lot_rules = {}
    if corps_etat_registry is None:
        shared_element_types = load_element_types(excel_file_path)
        shared_required_psets = load_required_psets_and_params(excel_file_path)
        for lot_name, _, _ in lots:
            lot_rules[lot_name] = (shared_element_types, shared_required_psets, "règles partagées")
    else:
        for lot_name, _, explicit_code in lots:
            # Choix explicite de l'utilisateur (sélecteur du dashboard) :
            # prioritaire sur toute correspondance automatique par nom.
            if explicit_code is not None and explicit_code in corps_etat_registry:
                code, method = explicit_code, "choix manuel"
            else:
                code, method = corps_etat.match_lot_name(lot_name, corps_etat_registry)
                if code is None and method == "ambigu":
                    ai_code = match_lot_with_ai(lot_name, corps_etat_registry)
                    if ai_code is not None:
                        code, method = ai_code, "IA"
            if code is not None:
                lot_rules[lot_name] = (
                    load_element_types(excel_file_path, code),
                    load_required_psets_and_params(excel_file_path, code),
                    f"{corps_etat_registry[code]} ({method})",
                )
            else:
                lot_rules[lot_name] = (None, None, "aucune correspondance -- bilan carbone seul")
            print(f"Lot '{lot_name}' -> corps d'état : {lot_rules[lot_name][2]}")

    # Préfiltrage des éléments
    elements_by_class = {}          # combiné, tous lots confondus
    elements_by_class_by_lot = {}   # {lot_name: {ifc_class: {...}}}
    missing_psets = 0
    missing_params = 0
    # Compte, par (lot, étage, type, PSet), le nombre d'éléments concernés par
    # un manque. Les préconisations IA restent volontairement lot-agnostiques
    # (décision utilisateur) : la dimension lot est retirée avant de leur être
    # transmise (voir plus bas), elle ne sert ici qu'au rapport Excel.
    missing_combo_counts = Counter()

    # Statistiques par étage, combinées tous lots confondus
    floor_stats = defaultdict(lambda: {"valid": 0, "invalid": 0})

    site_coords = None  # premier lot avec des coordonnées de site renseignées

    # Préparation du rapport Excel
    workbook = Workbook()
    details = workbook.active
    details.title = "Détails"
    
    # Styles pour l'onglet de détails -- enregistrés une fois comme NamedStyle
    # et appliqués via `cell.style = <nom>` (voir _write_detail_row) pour
    # éviter de rehasher des objets Font/Border/Fill/Alignment sur chacune
    # des dizaines de milliers de cellules de cet onglet.
    normal_font = Font(size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')

    named_styles = [
        NamedStyle(name="details_header", font=Font(size=12, bold=True),
                   fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
                   alignment=center_alignment, border=border),
        NamedStyle(name="details_normal_left", font=normal_font, alignment=left_alignment, border=border),
        NamedStyle(name="details_normal_center", font=normal_font, alignment=center_alignment, border=border),
        NamedStyle(name="details_status_ok", font=normal_font,
                   fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                   alignment=center_alignment, border=border),
        NamedStyle(name="details_status_ko", font=normal_font,
                   fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                   alignment=center_alignment, border=border),
    ]
    for style in named_styles:
        workbook.add_named_style(style)

    # En-têtes de l'onglet de détails -- "Lot" en premier, c'est l'axe de tri
    # principal pour une analyse multilot.
    headers = ["Lot", "Type", "Étage", "ID", "Nom", "PSet", "Paramètre", "Valeur", "Statut"]
    details_col_widths = [len(h) for h in headers]
    for col, header in enumerate(headers, 1):
        cell = details.cell(row=1, column=col)
        cell.value = header
        cell.style = "details_header"

    row = 2
    carbon_results = []

    for lot_name, ifc_file_path, _ in lots:
        print(f"Analyzing lot '{lot_name}'... ({time.time() - start_time:.2f}s)")
        ifc_file = ifcopenshell.open(ifc_file_path)
        unit_scale = carbon.get_length_unit_scale(ifc_file)
        if site_coords is None:
            site_coords = site_location.get_site_location(ifc_file)
        lot_classes = elements_by_class_by_lot.setdefault(lot_name, {})
        lot_element_types, lot_required_psets, lot_match_info = lot_rules[lot_name]

        for element in ifc_file.by_type('IfcProduct'):
            if lot_element_types is not None and element.is_a() in lot_element_types:
                ifc_class = element.is_a()
                if ifc_class not in elements_by_class:
                    elements_by_class[ifc_class] = {"total": 0, "valid": 0, "invalid": 0}
                if ifc_class not in lot_classes:
                    lot_classes[ifc_class] = {"total": 0, "valid": 0, "invalid": 0}

                elements_by_class[ifc_class]["total"] += 1
                lot_classes[ifc_class]["total"] += 1
                floor = get_building_storey(ifc_file, element)
                element_psets = ifcopenshell.util.element.get_psets(element)
                element_valid = True
                has_missing_pset = False
                has_missing_param = False
                # Évite de compter deux fois le même (lot, étage, type, PSet)
                # si plusieurs paramètres du même PSet manquent sur cet élément.
                combos_seen_for_element = set()

                # Vérification des PSet et paramètres requis -- .get() et non
                # [] : une classe listée dans Element_Types peut n'avoir aucune
                # règle PSet propre (ex. toutes ses règles "Corps_Etat" sont
                # réservées à une autre discipline que celle de ce lot), ce
                # qui est "rien à vérifier", pas une erreur.
                for pset_name, params in lot_required_psets.get(ifc_class, {}).items():
                    if pset_name not in element_psets:
                        element_valid = False
                        has_missing_pset = True
                        combo = (lot_name, str(floor), ifc_class, pset_name)
                        if combo not in combos_seen_for_element:
                            missing_combo_counts[combo] += 1
                            combos_seen_for_element.add(combo)

                        cells = [
                            (1, lot_name, "details_normal_left"),
                            (2, ifc_class, "details_normal_left"),
                            (3, str(floor), "details_normal_left"),
                            (4, element.GlobalId, "details_normal_center"),
                            (5, getattr(element, 'Name', ''), "details_normal_left"),
                            (6, pset_name, "details_normal_left"),
                            (7, "TOUS", "details_normal_center"),
                            (8, "MANQUANT", "details_normal_center"),
                            (9, "KO", "details_status_ko"),
                        ]
                        _write_detail_row(details, row, cells, details_col_widths)
                        row += 1
                    else:
                        for param_name, param_type in params.items():
                            actual_value = element_psets[pset_name].get(param_name)

                            if actual_value is None:
                                element_valid = False
                                has_missing_param = True
                                combo = (lot_name, str(floor), ifc_class, pset_name)
                                if combo not in combos_seen_for_element:
                                    missing_combo_counts[combo] += 1
                                    combos_seen_for_element.add(combo)

                                cells = [
                                    (1, lot_name, "details_normal_left"),
                                    (2, ifc_class, "details_normal_left"),
                                    (3, str(floor), "details_normal_left"),
                                    (4, element.GlobalId, "details_normal_center"),
                                    (5, getattr(element, 'Name', ''), "details_normal_left"),
                                    (6, pset_name, "details_normal_left"),
                                    (7, param_name, "details_normal_left"),
                                    (8, "MANQUANT", "details_normal_center"),
                                    (9, "KO", "details_status_ko"),
                                ]
                                _write_detail_row(details, row, cells, details_col_widths)
                                row += 1
                            else:
                                try:
                                    param_type_class = str_to_type(param_type)
                                    param_type_class(actual_value)

                                    cells = [
                                        (1, lot_name, "details_normal_left"),
                                        (2, ifc_class, "details_normal_left"),
                                        (3, str(floor), "details_normal_left"),
                                        (4, element.GlobalId, "details_normal_center"),
                                        (5, getattr(element, 'Name', ''), "details_normal_left"),
                                        (6, pset_name, "details_normal_left"),
                                        (7, param_name, "details_normal_left"),
                                        (8, str(actual_value), "details_normal_left"),
                                        (9, "OK", "details_status_ok"),
                                    ]
                                    _write_detail_row(details, row, cells, details_col_widths)
                                    row += 1
                                except (ValueError, TypeError):
                                    element_valid = False
                                    has_missing_param = True
                                    combo = (lot_name, str(floor), ifc_class, pset_name)
                                    if combo not in combos_seen_for_element:
                                        missing_combo_counts[combo] += 1
                                        combos_seen_for_element.add(combo)

                                    cells = [
                                        (1, lot_name, "details_normal_left"),
                                        (2, ifc_class, "details_normal_left"),
                                        (3, str(floor), "details_normal_left"),
                                        (4, element.GlobalId, "details_normal_center"),
                                        (5, getattr(element, 'Name', ''), "details_normal_left"),
                                        (6, pset_name, "details_normal_left"),
                                        (7, param_name, "details_normal_left"),
                                        (8, f"{actual_value} (attendu: {param_type})", "details_normal_left"),
                                        (9, "KO", "details_status_ko"),
                                    ]
                                    _write_detail_row(details, row, cells, details_col_widths)
                                    row += 1

                # Mise à jour des statistiques
                if element_valid:
                    elements_by_class[ifc_class]["valid"] += 1
                    lot_classes[ifc_class]["valid"] += 1
                    floor_stats[floor]["valid"] += 1
                else:
                    elements_by_class[ifc_class]["invalid"] += 1
                    lot_classes[ifc_class]["invalid"] += 1
                    floor_stats[floor]["invalid"] += 1
                    if has_missing_pset:
                        missing_psets += 1
                    if has_missing_param:
                        missing_params += 1

                # Empreinte carbone (voir carbon.py) : un élément multicouche
                # produit plusieurs résultats, un par couche de matériau.
                carbon_results.extend(carbon.calculate_element_carbon(element, str(floor), unit_scale, lot=lot_name))
            elif lot_element_types is None:
                # Aucun corps d'état n'a pu être associé à ce lot : pas de
                # contrôle PSet possible, mais le bilan carbone reste calculé
                # sur tous les éléments (décision utilisateur -- jamais bloquant).
                floor = get_building_storey(ifc_file, element)
                carbon_results.extend(carbon.calculate_element_carbon(element, str(floor), unit_scale, lot=lot_name))

    # Statistiques globales
    total_elements = sum(stats["total"] for stats in elements_by_class.values())
    valid_elements = sum(stats["valid"] for stats in elements_by_class.values())
    invalid_elements = sum(stats["invalid"] for stats in elements_by_class.values())

    # Créer l'onglet de résumé
    create_summary_sheet(workbook, total_elements, valid_elements, invalid_elements, missing_psets, missing_params, floor_stats, elements_by_class)

    carbon_facts = carbon.aggregate_carbon_results(carbon_results)
    print(f"Bilan carbone : total {carbon_facts['total']:.2f} kg CO2e, "
          f"fiabilité {carbon_facts['reliability_summary']}")
    carbon.create_carbon_footprint_sheet(workbook, carbon_facts)
    lot_descriptions = {lot_name: info[2] for lot_name, info in lot_rules.items()}
    create_lot_comparison_sheet(workbook, elements_by_class_by_lot, carbon_facts, lot_descriptions)

    # Préconisations (agents IA) : jamais bloquant -- en cas d'échec/indisponibilité,
    # run_recommendations_pipeline renvoie un résultat "degraded" plutôt que de lever.
    # Les faits restent lot-agnostiques (décision utilisateur) : la dimension
    # lot de missing_combo_counts est retirée par sommation avant transmission.
    combined_missing_counts = Counter()
    for (lot_name, floor, ifc_class, pset), count in missing_combo_counts.items():
        combined_missing_counts[(floor, ifc_class, pset)] += count
    top_missing = sorted(combined_missing_counts.items(), key=lambda item: -item[1])[:15]
    validation_facts = {
        "total_elements": total_elements,
        "valid_elements": valid_elements,
        "missing_elements": invalid_elements,
        "elements_by_class": elements_by_class,
        "top_missing": [
            {"floor": floor, "ifc_class": ifc_class, "pset": pset, "count": count}
            for (floor, ifc_class, pset), count in top_missing
        ],
    }
    print("Génération des préconisations...")
    recommendations = run_recommendations_pipeline(validation_facts, carbon_facts, elements_by_class)
    print(f"Préconisations : degraded={recommendations['degraded']} tiers={recommendations['tiers_used']}")
    create_recommendations_sheet(workbook, recommendations)

    # Comparatif par lot (même contenu que l'onglet Excel "Comparatif Lots"
    # et utilisé aussi par la Fiche de Recette ci-dessous). Le champ "folder"
    # est ajouté par la route /upload (process_files ne connaît pas la
    # structure de dossiers sur disque).
    lots_summary = []
    for lot_name, _, _ in lots:
        lot_classes = elements_by_class_by_lot.get(lot_name, {})
        lot_total = sum(s["total"] for s in lot_classes.values())
        lot_valid = sum(s["valid"] for s in lot_classes.values())
        lots_summary.append({
            "name": lot_name,
            "total_elements": lot_total,
            "valid_elements": lot_valid,
            "carbon_total": carbon_facts["by_lot"].get(lot_name, 0),
            "corps_etat": lot_rules[lot_name][2],
            "has_rules": lot_rules[lot_name][0] is not None,
        })

    # Anomalies = non-conformités agrégées par (lot, étage, classe IFC, PSet)
    # -- voir history.py pour le choix de cette granularité. Alimente à la
    # fois la Fiche de Recette ci-dessous et l'historique persistant
    # (comparaison entre deux analyses -- voir /compare dans Checkers.py).
    anomalies_raw = [
        {"lot": lot_name, "floor": floor, "ifc_class": ifc_class, "pset": pset, "count": count}
        for (lot_name, floor, ifc_class, pset), count in missing_combo_counts.items()
    ]
    create_recette_sheet(workbook, label, analysis_timestamp.strftime("%d/%m/%Y %H:%M"),
                          excel_filename, lots_summary, anomalies_raw)

    # Largeurs de colonnes de l'onglet Détails : déjà calculées à la volée
    # pendant l'écriture (voir _write_detail_row) pour éviter de relire les
    # dizaines de milliers de cellules de cet onglet une seconde fois ici.
    for col, width in enumerate(details_col_widths, 1):
        details.column_dimensions[get_column_letter(col)].width = width + 2

    # Ajuster les largeurs des colonnes des autres onglets
    for ws in workbook.worksheets:
        if ws is details:
            continue
        for column in ws.columns:
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width
    
    # Figer les volets dans l'onglet de détails
    details.freeze_panes = 'A2'
    
    # Sauvegarder le fichier
    print(f"Saving file... ({time.time() - start_time:.2f}s)")
    workbook.save(output_file_path)
    print(f"Analysis completed in {time.time() - start_time:.2f}s")
    
    # Préparation des données pour le dashboard
    sorted_floors = sorted(floor_stats.items(), key=lambda x: sort_floor_name(x[0]))

    return {
        "label": label,
        "total_elements": total_elements,
        "valid_elements": valid_elements,
        "missing_elements": invalid_elements,
        "missing_psets": missing_psets,
        "missing_params": missing_params,
        "floors": [{"name": floor, "valid": stats["valid"], "invalid": stats["invalid"]} for floor, stats in sorted_floors],
        "lots": lots_summary,
        "anomalies": anomalies_raw,
        "carbon_footprint": {
            "total": carbon_facts["total"],
            "by_type": carbon_facts["by_type"],
            "by_floor": carbon_facts["by_floor"],
            "reliability_summary": carbon_facts["reliability_summary"],
        },
        "site_location": site_coords,
        "recommendations": {
            "carbon_expert": recommendations["carbon_expert"],
            "data_quality_expert": recommendations["data_quality_expert"],
            "business_expert": recommendations["business_expert"],
            "synthesis": recommendations["synthesis"],
            "tiers_used": recommendations["tiers_used"],
            "degraded": recommendations["degraded"],
            "degraded_reason": recommendations["degraded_reason"],
        },
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/corps_etat_registry', methods=['POST'])
def corps_etat_registry_route():
    """Aperçu léger : lit uniquement l'onglet "Corps_Etat" du fichier Excel
    envoyé (sans lancer d'analyse) pour peupler le sélecteur de corps d'état
    de chaque lot côté dashboard, dès que l'utilisateur choisit son fichier
    Excel. Retourne {"registry": null} si l'onglet n'existe pas (mode
    legacy, un seul jeu de règles -- pas de sélecteur à afficher)."""
    try:
        if 'excel_file' not in request.files:
            return jsonify({"error": "Missing Excel file"}), 400
        excel_file = request.files['excel_file']
        if excel_file.filename == '' or not allowed_file(excel_file.filename):
            return jsonify({"error": "Invalid Excel file"}), 400

        fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            excel_file.save(tmp_path)
            try:
                _reject_if_zip_bomb(tmp_path)
            except ValueError as e:
                return jsonify({"error": str(e)}), 400
            registry = corps_etat.load_registry(tmp_path)
        finally:
            os.remove(tmp_path)

        return jsonify({"registry": registry})
    except Exception as e:
        print(f"Error reading corps d'état registry: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload():
    try:
        print("Starting upload...")
        if 'excel_file' not in request.files:
            print("Missing Excel file in request")
            return jsonify({"error": "Missing Excel file"}), 400

        excel_file = request.files['excel_file']
        if excel_file.filename == '' or not allowed_file(excel_file.filename):
            print("Invalid Excel file")
            return jsonify({"error": "Invalid Excel file"}), 400

        # Collecte des lots : champs indexés lot_name_N/ifc_file_N -- au moins
        # un lot requis. Un seul fichier Excel de paramètres requis, partagé
        # par tous les lots. Les index sont scannés (pas supposés contigus à
        # partir de 0) pour que le frontend puisse retirer une ligne de lot
        # sans avoir à renuméroter les autres.
        lot_indices = sorted(
            int(match.group(1))
            for key in request.files
            for match in [re.match(r'ifc_file_(\d+)$', key)]
            if match
        )
        lot_uploads = []  # (lot_name, FileStorage, index, explicit_corps_etat_code|None)
        for i in lot_indices:
            ifc_file = request.files[f'ifc_file_{i}']
            lot_name = request.form.get(f'lot_name_{i}', '').strip() or f"Lot {i + 1}"
            # Choix explicite du corps d'état via le sélecteur du dashboard
            # (voir /corps_etat_registry) -- vide/absent = "Auto", le lot
            # passe par la correspondance automatique habituelle (exact,
            # puis floue, puis IA en dernier recours).
            explicit_code = request.form.get(f'corps_etat_{i}', '').strip() or None
            if ifc_file.filename != '':
                if not allowed_file(ifc_file.filename):
                    print(f"Invalid file type for lot '{lot_name}'")
                    return jsonify({"error": f"Invalid file type for lot '{lot_name}'"}), 400
                lot_uploads.append((lot_name, ifc_file, i, explicit_code))

        if not lot_uploads:
            print("No IFC file in request")
            return jsonify({"error": "No IFC file provided"}), 400

        print(f"Received {len(lot_uploads)} lot(s), Excel={excel_file.filename}")

        # Générer un ID unique pour cette analyse
        analysis_id = str(uuid.uuid4())
        analysis_dir = os.path.join(TEMP_FOLDER, analysis_id)
        print(f"Creating analysis directory: {analysis_dir}")
        try:
            os.makedirs(analysis_dir, exist_ok=True)
        except Exception as e:
            print(f"Error creating analysis directory: {str(e)}")
            return jsonify({"error": f"Could not create analysis directory: {str(e)}"}), 500

        excel_path = os.path.join(analysis_dir, secure_filename(excel_file.filename))
        try:
            excel_file.save(excel_path)
        except Exception as e:
            print(f"Error saving Excel file: {str(e)}")
            return jsonify({"error": f"Could not save Excel file: {str(e)}"}), 500

        try:
            _reject_if_zip_bomb(excel_path)
        except ValueError as e:
            print(f"Rejected Excel file: {str(e)}")
            return jsonify({"error": str(e)}), 400

        # Un sous-dossier par lot (évite les collisions de nom de fichier
        # entre lots) ; "folder" est renvoyé au frontend pour cibler la
        # visionneuse 3D sur le bon lot (/model/<analysis_id>/<folder>).
        lots = []         # (lot_name, ifc_path, explicit_corps_etat_code) pour process_files
        lot_folders = []  # folder correspondant, même ordre que `lots`
        for lot_name, ifc_file, i, explicit_code in lot_uploads:
            folder = f"{i:02d}_{secure_filename(lot_name) or f'lot{i}'}"
            lot_dir = os.path.join(analysis_dir, folder)
            os.makedirs(lot_dir, exist_ok=True)
            ifc_path = os.path.join(lot_dir, secure_filename(ifc_file.filename))
            try:
                ifc_file.save(ifc_path)
            except Exception as e:
                print(f"Error saving IFC file for lot '{lot_name}': {str(e)}")
                return jsonify({"error": f"Could not save IFC file for lot '{lot_name}': {str(e)}"}), 500
            lots.append((lot_name, ifc_path, explicit_code))
            lot_folders.append(folder)

        output_path = os.path.join(analysis_dir, 'output_rapport.xlsx')
        # Repère libre affiché dans l'historique et la Fiche de Recette --
        # par défaut le nom du fichier Excel si l'utilisateur n'en saisit pas.
        label = request.form.get('analysis_label', '').strip() or None

        # Analyser les fichiers
        print("Starting analysis...")
        try:
            results = process_files(analysis_dir, lots, excel_path, output_path, label=label)
            for lot_summary, folder in zip(results["lots"], lot_folders):
                lot_summary["folder"] = folder
            results["analysis_id"] = analysis_id
            print("Analysis completed successfully")
            try:
                history.record_analysis(analysis_id, results["label"], excel_file.filename, results, results["anomalies"])
            except Exception as e:
                # L'historique est un bonus (tableau de bord d'exploitation) --
                # un échec de persistance ne doit jamais invalider une analyse
                # par ailleurs réussie, dont le rapport Excel est déjà sur disque.
                print(f"Warning: could not record analysis history: {str(e)}")
            return jsonify(results)
        except Exception as e:
            print(f"Error during analysis: {str(e)}")
            return jsonify({"error": f"Analysis failed: {str(e)}"}), 500

    except Exception as e:
        print(f"Unexpected error during upload: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/download/<analysis_id>')
def download(analysis_id):
    try:
        if not _valid_analysis_id(analysis_id):
            abort(404)
        analysis_dir = os.path.join(TEMP_FOLDER, analysis_id)
        if not os.path.exists(analysis_dir):
            return jsonify({"error": "Analysis not found"}), 404

        # Trouver le fichier Excel de sortie
        output_files = [f for f in os.listdir(analysis_dir) if f.startswith('output_') and f.endswith('.xlsx')]
        if not output_files:
            return jsonify({"error": "Output file not found"}), 404

        output_path = os.path.join(analysis_dir, output_files[0])
        return send_file(output_path, as_attachment=True, download_name=output_files[0])

    except Exception as e:
        print(f"Error during download: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/model/<analysis_id>/<lot_folder>')
def model(analysis_id, lot_folder):
    """Sert le fichier IFC brut d'un lot donné, pour la visionneuse 3D côté
    client (le `folder` de chaque lot est renvoyé par /upload)."""
    try:
        # `analysis_id`/`lot_folder` viennent de l'URL -- on les valide avant
        # de les utiliser comme composants de chemin (sinon traversée de
        # répertoire possible, ex. lot_folder="../../.."). `lot_folder` est
        # toujours de la forme "<index>_<nom_secure_filename>" (voir /upload) :
        # secure_filename() doit donc le laisser inchangé.
        if not _valid_analysis_id(analysis_id) or secure_filename(lot_folder) != lot_folder:
            abort(404)
        lot_dir = os.path.join(TEMP_FOLDER, analysis_id, lot_folder)
        if not os.path.exists(lot_dir):
            return jsonify({"error": "Lot not found"}), 404

        ifc_files = [f for f in os.listdir(lot_dir) if f.lower().endswith('.ifc')]
        if not ifc_files:
            return jsonify({"error": "IFC file not found"}), 404

        ifc_path = os.path.join(lot_dir, ifc_files[0])
        return send_file(ifc_path, mimetype='application/octet-stream')

    except Exception as e:
        print(f"Error serving IFC model: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/history')
def history_route():
    """Historique persistant des analyses (voir history.py) -- alimente
    l'onglet "Historique" du dashboard : suivi dans le temps et point de
    départ d'une comparaison entre deux analyses."""
    try:
        return jsonify({"analyses": history.list_analyses()})
    except Exception as e:
        print(f"Error reading history: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/compare/<before_id>/<after_id>')
def compare_route(before_id, after_id):
    """Diff entre deux analyses déjà enregistrées : anomalies corrigées,
    nouvelles, persistantes -- le cœur du suivi de recette entre deux
    livraisons successives d'une même maquette (ou de deux lots à comparer)."""
    try:
        if not _valid_analysis_id(before_id) or not _valid_analysis_id(after_id):
            abort(404)
        before = history.get_analysis(before_id)
        after = history.get_analysis(after_id)
        if before is None or after is None:
            return jsonify({"error": "Analyse introuvable (a-t-elle été supprimée du dossier temp ?)"}), 404
        diff = history.compare_analyses(before_id, after_id)
        return jsonify({"before": before, "after": after, "diff": diff})
    except Exception as e:
        print(f"Error comparing analyses: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # Configuration du port via variable d'environnement pour Vercel
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
