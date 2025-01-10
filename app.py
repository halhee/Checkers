from flask import Flask, render_template, request, jsonify
import os
from werkzeug.utils import secure_filename
from flask_cors import CORS
import logging
import ifcopenshell
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
TEMP_FOLDER = 'temp'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload and temp directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

# Prix des matériaux (€/unité)
MATERIAL_PRICES = {
    # Gros œuvre
    'béton': {
        'price': 120,  # €/m³
        'unit': 'm³',
        'category': 'Gros œuvre',
        'description': 'Béton standard pour structure'
    },
    'béton_haute_performance': {
        'price': 180,  # €/m³
        'unit': 'm³',
        'category': 'Gros œuvre',
        'description': 'Béton haute résistance'
    },
    'béton_préfabriqué': {
        'price': 150,  # €/m³
        'unit': 'm³',
        'category': 'Gros œuvre',
        'description': 'Éléments en béton préfabriqués'
    },
    'acier_construction': {
        'price': 2.5,  # €/kg
        'unit': 'kg',
        'category': 'Gros œuvre',
        'description': 'Acier de construction standard'
    },
    'acier_inox': {
        'price': 8.0,  # €/kg
        'unit': 'kg',
        'category': 'Gros œuvre',
        'description': 'Acier inoxydable'
    },
    'bois_construction': {
        'price': 800,  # €/m³
        'unit': 'm³',
        'category': 'Gros œuvre',
        'description': 'Bois de construction standard'
    },
    'bois_lamellé_collé': {
        'price': 1200,  # €/m³
        'unit': 'm³',
        'category': 'Gros œuvre',
        'description': 'Bois lamellé-collé pour structure'
    },
    'pierre_naturelle': {
        'price': 250,  # €/m²
        'unit': 'm²',
        'category': 'Gros œuvre',
        'description': 'Pierre naturelle pour façade'
    },

    # Second œuvre
    'verre_simple': {
        'price': 50,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Vitrage simple'
    },
    'verre_double': {
        'price': 80,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Double vitrage standard'
    },
    'verre_triple': {
        'price': 120,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Triple vitrage haute performance'
    },
    'isolation_laine_verre': {
        'price': 25,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Isolation en laine de verre'
    },
    'isolation_laine_roche': {
        'price': 30,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Isolation en laine de roche'
    },
    'isolation_polyuréthane': {
        'price': 35,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Isolation en polyuréthane'
    },
    'peinture_standard': {
        'price': 15,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Peinture acrylique standard'
    },
    'peinture_haute_qualité': {
        'price': 25,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Peinture haut de gamme'
    },
    'carrelage_standard': {
        'price': 45,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Carrelage céramique standard'
    },
    'carrelage_grès': {
        'price': 65,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Carrelage en grès cérame'
    },
    'parquet_standard': {
        'price': 55,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Parquet en bois standard'
    },
    'parquet_massif': {
        'price': 85,  # €/m²
        'unit': 'm²',
        'category': 'Second œuvre',
        'description': 'Parquet en bois massif'
    },

    # Menuiseries
    'menuiserie_alu': {
        'price': 350,  # €/m²
        'unit': 'm²',
        'category': 'Menuiseries',
        'description': 'Menuiserie aluminium standard'
    },
    'menuiserie_alu_rupture': {
        'price': 450,  # €/m²
        'unit': 'm²',
        'category': 'Menuiseries',
        'description': 'Menuiserie aluminium à rupture de pont thermique'
    },
    'menuiserie_bois': {
        'price': 250,  # €/m²
        'unit': 'm²',
        'category': 'Menuiseries',
        'description': 'Menuiserie en bois'
    },
    'menuiserie_pvc': {
        'price': 200,  # €/m²
        'unit': 'm²',
        'category': 'Menuiseries',
        'description': 'Menuiserie en PVC'
    },

    # Toiture
    'tuiles_terre_cuite': {
        'price': 35,  # €/m²
        'unit': 'm²',
        'category': 'Toiture',
        'description': 'Tuiles en terre cuite'
    },
    'tuiles_béton': {
        'price': 25,  # €/m²
        'unit': 'm²',
        'category': 'Toiture',
        'description': 'Tuiles en béton'
    },
    'ardoise_naturelle': {
        'price': 60,  # €/m²
        'unit': 'm²',
        'category': 'Toiture',
        'description': 'Ardoise naturelle'
    },
    'zinc': {
        'price': 75,  # €/m²
        'unit': 'm²',
        'category': 'Toiture',
        'description': 'Couverture en zinc'
    },
    'membrane_epdm': {
        'price': 45,  # €/m²
        'unit': 'm²',
        'category': 'Toiture',
        'description': 'Membrane d\'étanchéité EPDM'
    },

    # Équipements techniques
    'climatisation': {
        'price': 250,  # €/m²
        'unit': 'm²',
        'category': 'Équipements',
        'description': 'Système de climatisation'
    },
    'ventilation_simple': {
        'price': 35,  # €/m²
        'unit': 'm²',
        'category': 'Équipements',
        'description': 'Ventilation simple flux'
    },
    'ventilation_double': {
        'price': 65,  # €/m²
        'unit': 'm²',
        'category': 'Équipements',
        'description': 'Ventilation double flux'
    },
    'panneaux_solaires': {
        'price': 350,  # €/m²
        'unit': 'm²',
        'category': 'Équipements',
        'description': 'Panneaux solaires photovoltaïques'
    },
    'pompe_chaleur': {
        'price': 12000,  # €/unité
        'unit': 'unité',
        'category': 'Équipements',
        'description': 'Pompe à chaleur'
    },

    # VRD (Voirie et Réseaux Divers)
    'enrobé': {
        'price': 45,  # €/m²
        'unit': 'm²',
        'category': 'VRD',
        'description': 'Revêtement en enrobé'
    },
    'pavés': {
        'price': 65,  # €/m²
        'unit': 'm²',
        'category': 'VRD',
        'description': 'Pavage extérieur'
    },
    'bordures': {
        'price': 35,  # €/ml
        'unit': 'ml',
        'category': 'VRD',
        'description': 'Bordures de trottoir'
    },
    'réseaux_eau': {
        'price': 120,  # €/ml
        'unit': 'ml',
        'category': 'VRD',
        'description': 'Réseaux d\'eau potable et usée'
    }
}

def get_material_from_element(element):
    """Déterminer le matériau d'un élément IFC."""
    try:
        material_select = element.HasAssociations
        for rel in material_select:
            if rel.is_a('IfcRelAssociatesMaterial'):
                material = rel.RelatingMaterial
                if material.is_a('IfcMaterial'):
                    return material.Name.lower()
                elif material.is_a('IfcMaterialLayer'):
                    return material.Material.Name.lower()
                elif material.is_a('IfcMaterialLayerSet'):
                    return material.MaterialLayers[0].Material.Name.lower()
    except:
        pass
    return None

def calculate_element_quantity(element):
    """Calculer la quantité d'un élément selon son type."""
    try:
        if element.is_a('IfcWall'):
            # Calculer le volume pour les murs
            length = element.get_attribute('Length', 0)
            height = element.get_attribute('Height', 0)
            thickness = element.get_attribute('Width', 0)
            return length * height * thickness
        elif element.is_a('IfcSlab') or element.is_a('IfcRoof'):
            # Calculer la surface pour les dalles et toits
            length = element.get_attribute('Length', 0)
            width = element.get_attribute('Width', 0)
            return length * width
        elif element.is_a('IfcWindow') or element.is_a('IfcDoor'):
            # Calculer la surface pour les fenêtres et portes
            height = element.get_attribute('Height', 0)
            width = element.get_attribute('Width', 0)
            return height * width
    except:
        pass
    return 1  # Valeur par défaut si impossible de calculer

def calculate_construction_costs(ifc_file):
    """Calculer les coûts de construction à partir du fichier IFC."""
    try:
        # Charger le fichier IFC
        ifc = ifcopenshell.open(ifc_file)
        
        # Initialiser les variables de coût
        total_cost = 0
        costs_by_category = {
            'Gros œuvre': 0,
            'Second œuvre': 0,
            'Équipements': 0,
            'VRD': 0,
            'Études et honoraires': 0
        }
        
        # Initialiser les coûts par matériau
        material_costs = {material: 0 for material in MATERIAL_PRICES}
        
        # Calculer la surface totale
        spaces = ifc.by_type('IfcSpace')
        total_area = sum(space.get_attribute('NetFloorArea', 0) for space in spaces)
        
        # Parcourir les éléments et calculer les coûts
        for element in ifc.by_type('IfcElement'):
            material = get_material_from_element(element)
            quantity = calculate_element_quantity(element)
            
            if material and material in MATERIAL_PRICES:
                # Calculer le coût selon le matériau
                unit_cost = MATERIAL_PRICES[material]['price']
                element_cost = quantity * unit_cost
                material_costs[material] += element_cost
                costs_by_category[MATERIAL_PRICES[material]['category']] += element_cost
                total_cost += element_cost
            else:
                # Utiliser les coûts par défaut si le matériau n'est pas reconnu
                element_type = element.is_a()
                if 'Wall' in element_type or 'Beam' in element_type or 'Column' in element_type:
                    element_cost = quantity * 500
                    costs_by_category['Gros œuvre'] += element_cost
                elif 'Door' in element_type or 'Window' in element_type:
                    element_cost = quantity * 300
                    costs_by_category['Second œuvre'] += element_cost
                elif 'Equipment' in element_type:
                    element_cost = quantity * 1000
                    costs_by_category['Équipements'] += element_cost
                else:
                    element_cost = quantity * 200
                    costs_by_category['VRD'] += element_cost
                total_cost += element_cost
        
        # Ajouter les coûts d'études (15% du total)
        costs_by_category['Études et honoraires'] = total_cost * 0.15
        total_cost *= 1.15
        
        # Générer des données temporelles
        timeline_data = generate_cost_timeline(total_cost)
        
        return {
            'total': round(total_cost, 2),
            'perSquareMeter': round(total_cost / total_area if total_area > 0 else 0, 2),
            'breakdown': {
                'categories': list(costs_by_category.keys()),
                'values': list(costs_by_category.values())
            },
            'materials': {
                'names': list(material_costs.keys()),
                'costs': list(material_costs.values()),
                'units': [MATERIAL_PRICES[m]['unit'] for m in material_costs.keys()]
            },
            'timeline': timeline_data
        }
    except Exception as e:
        logger.error(f"Erreur lors du calcul des coûts: {str(e)}")
        return None

def generate_cost_timeline(total_cost):
    """Générer une simulation de l'évolution des coûts dans le temps."""
    start_date = datetime.now()
    dates = []
    values = []
    cumulative_cost = 0
    
    # Simuler 12 mois de progression
    for i in range(12):
        current_date = start_date + timedelta(days=30*i)
        # Distribution non linéaire des coûts
        if i < 3:
            monthly_cost = total_cost * 0.1  # 10% par mois au début
        elif i < 8:
            monthly_cost = total_cost * 0.15  # 15% par mois pendant la phase principale
        else:
            monthly_cost = total_cost * 0.05  # 5% par mois à la fin
        
        cumulative_cost += monthly_cost
        dates.append(current_date.strftime('%Y-%m-%d'))
        values.append(round(cumulative_cost, 2))
    
    return {
        'dates': dates,
        'values': values
    }

@app.route('/')
def index():
    logger.info('Accessing index page')
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    logger.info('Upload request received')
    logger.debug(f'Files in request: {request.files}')
    
    if 'file' not in request.files:
        logger.error('No file part in request')
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        logger.error('No selected file')
        return jsonify({'error': 'No selected file'}), 400

    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        logger.info(f'Saving file to: {filepath}')
        try:
            file.save(filepath)
            logger.info('File saved successfully')
            
            # Calculer les coûts
            costs = calculate_construction_costs(filepath)
            if costs:
                return jsonify({
                    'message': 'File uploaded and analyzed successfully',
                    'filename': filename,
                    'costs': costs
                }), 200
            else:
                return jsonify({
                    'error': 'Error analyzing file costs'
                }), 500
                
        except Exception as e:
            logger.error(f'Error saving or analyzing file: {str(e)}')
            return jsonify({'error': f'Error processing file: {str(e)}'}), 500

if __name__ == '__main__':
    logger.info('Starting Flask application...')
    app.run(host='0.0.0.0', port=8080, debug=True)
