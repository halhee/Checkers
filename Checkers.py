from flask import Flask, render_template, request, redirect, url_for, send_file, abort, jsonify
import os
import tempfile
import uuid
import time
import shutil
import re
import zipfile
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

import corps_etat
import history
import ids_check

load_dotenv()

# Le moteur d'analyse (process_files, règles Excel, onglets du rapport) vit
# dans checkers_core/analysis.py, sans dépendance à Flask, pour être
# testable et réutilisable hors de l'application web.
# Importé après load_dotenv() : les agents IA lisent leur configuration
# (.env) au moment où ils sont appelés.
from checkers_core.analysis import process_files
from checkers_core.bcf import write_bcf

BCF_FILENAME = 'output_anomalies.bcf'

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

ALLOWED_EXTENSIONS = {'ifc', 'xlsx', 'ids'}

# Identifiant d'analyse = uuid4 utilisé à la fois comme nom de dossier sur
# disque (TEMP_FOLDER/<id>/...) et comme clé SQLite -- validé dans toutes les
# routes qui le prennent en paramètre d'URL pour écarter toute tentative de
# traversée de chemin (ex. "../../etc/passwd") avant de s'en servir.
_ANALYSIS_ID_RE = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')


def _valid_analysis_id(analysis_id: str) -> bool:
    return bool(_ANALYSIS_ID_RE.match(analysis_id))

def allowed_file(filename: str, extensions=None) -> bool:
    """`extensions` : jeu d'extensions autorisées pour CE champ (ex. {'ifc'}) ;
    par défaut tout ce que l'app sait lire. Chaque champ d'upload doit passer
    le sien -- un .xlsx n'a rien à faire dans un champ IFC."""
    allowed = ALLOWED_EXTENSIONS if extensions is None else extensions
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed

# Durée de conservation des dossiers d'analyse sur le serveur (fichiers
# envoyés + rapport généré) -- les fichiers IFC/Excel uploadés peuvent
# contenir des données de projet réelles, inutile de les garder
# indéfiniment dans temp/ une fois l'analyse terminée.
ANALYSIS_RETENTION_SECONDS = int(os.environ.get('ANALYSIS_RETENTION_SECONDS', 3600))


def cleanup_old_analyses(max_age_seconds: int = ANALYSIS_RETENTION_SECONDS) -> None:
    """Supprime les dossiers d'analyse de TEMP_FOLDER plus vieux que
    `max_age_seconds`. Appelé à chaque upload plutôt que via un cron/thread
    séparé -- suffisant pour une app mono-instance, pas de dépendance
    supplémentaire à faire tourner."""
    now = time.time()
    try:
        entries = os.listdir(TEMP_FOLDER)
    except OSError:
        return
    for name in entries:
        path = os.path.join(TEMP_FOLDER, name)
        try:
            if os.path.isdir(path) and now - os.path.getmtime(path) > max_age_seconds:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass

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
        if excel_file.filename == '' or not allowed_file(excel_file.filename, {'xlsx'}):
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
        # Référentiel de règles : Excel (PSet) et/ou IDS -- au moins un des deux.
        excel_file = request.files.get('excel_file')
        if excel_file is not None and excel_file.filename == '':
            excel_file = None
        ids_file = request.files.get('ids_file')
        if ids_file is not None and ids_file.filename == '':
            ids_file = None

        if excel_file is None and ids_file is None:
            print("No rules file in request")
            return jsonify({"error": "Fournir au moins un fichier de règles (Excel et/ou IDS)"}), 400
        if excel_file is not None and not allowed_file(excel_file.filename, {'xlsx'}):
            print("Invalid Excel file")
            return jsonify({"error": "Invalid Excel file"}), 400
        if ids_file is not None and not allowed_file(ids_file.filename, {'ids'}):
            print("Invalid IDS file")
            return jsonify({"error": "Invalid IDS file (extension .ids attendue)"}), 400

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
                if not allowed_file(ifc_file.filename, {'ifc'}):
                    print(f"Invalid file type for lot '{lot_name}'")
                    return jsonify({"error": f"Invalid file type for lot '{lot_name}'"}), 400
                lot_uploads.append((lot_name, ifc_file, i, explicit_code))

        if not lot_uploads:
            print("No IFC file in request")
            return jsonify({"error": "No IFC file provided"}), 400

        print(f"Received {len(lot_uploads)} lot(s), Excel={excel_file.filename if excel_file else None}, IDS={ids_file.filename if ids_file else None}")

        # Purge les analyses précédentes plus vieilles que la rétention avant
        # d'en créer une nouvelle -- évite d'accumuler indéfiniment des
        # fichiers IFC/Excel uploadés (données de projet potentiellement
        # sensibles) dans temp/.
        cleanup_old_analyses()

        # Générer un ID unique pour cette analyse
        analysis_id = str(uuid.uuid4())
        analysis_dir = os.path.join(TEMP_FOLDER, analysis_id)
        print(f"Creating analysis directory: {analysis_dir}")
        try:
            os.makedirs(analysis_dir, exist_ok=True)
        except Exception as e:
            print(f"Error creating analysis directory: {str(e)}")
            return jsonify({"error": f"Could not create analysis directory: {str(e)}"}), 500

        excel_path = None
        if excel_file is not None:
            excel_path = os.path.join(analysis_dir, secure_filename(excel_file.filename))
            try:
                excel_file.save(excel_path)
            except Exception as e:
                print(f"Error saving Excel file: {str(e)}")
                shutil.rmtree(analysis_dir, ignore_errors=True)
                return jsonify({"error": f"Could not save Excel file: {str(e)}"}), 500

            try:
                _reject_if_zip_bomb(excel_path)
            except ValueError as e:
                print(f"Rejected Excel file: {str(e)}")
                shutil.rmtree(analysis_dir, ignore_errors=True)
                return jsonify({"error": str(e)}), 400

        ids_path = None
        if ids_file is not None:
            # Nom fixe côté disque : le nom d'origine n'est conservé que pour
            # l'affichage (secure_filename ne protège pas d'un nom vide).
            ids_path = os.path.join(analysis_dir, 'regles.ids')
            try:
                ids_file.save(ids_path)
            except Exception as e:
                print(f"Error saving IDS file: {str(e)}")
                shutil.rmtree(analysis_dir, ignore_errors=True)
                return jsonify({"error": f"Could not save IDS file: {str(e)}"}), 500

            # Validation anticipée (sécurité + schéma) : un IDS refusé doit
            # l'être avant de lancer une analyse longue, avec un message clair.
            try:
                ids_check.load_ids(ids_path)
            except ids_check.IdsFileError as e:
                print(f"Rejected IDS file: {str(e)}")
                shutil.rmtree(analysis_dir, ignore_errors=True)
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
                shutil.rmtree(analysis_dir, ignore_errors=True)
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
            ids_display_name = secure_filename(ids_file.filename) if ids_file is not None else None
            # collect_elements : GlobalId par anomalie, nécessaires à l'export
            # BCF ci-dessous ; retirés ensuite du résultat pour ne pas alourdir
            # la réponse JSON ni l'historique.
            results = process_files(analysis_dir, lots, excel_path, output_path, label=label,
                                    ids_file_path=ids_path, ids_filename=ids_display_name,
                                    collect_elements=True)
            anomalies_with_guids = results["anomalies"]
            results["anomalies"] = [{k: v for k, v in a.items() if k != "guids"} for a in anomalies_with_guids]
            results.pop("elements", None)
            results["bcf_available"] = False
            try:
                rules_description = " + ".join(
                    f"{kind} « {secure_filename(f.filename)} »"
                    for kind, f in (("Excel", excel_file), ("IDS", ids_file)) if f is not None)
                write_bcf(os.path.join(analysis_dir, BCF_FILENAME), anomalies_with_guids,
                          project_name=results["label"], rules_description=rules_description)
                results["bcf_available"] = True
            except Exception as e:
                # Export complémentaire : le rapport Excel reste la référence,
                # un échec ici ne doit pas faire échouer l'analyse.
                print(f"Warning: could not write BCF export: {str(e)}")
            for lot_summary, folder in zip(results["lots"], lot_folders):
                lot_summary["folder"] = folder
            results["analysis_id"] = analysis_id
            print("Analysis completed successfully")
            try:
                rules_filename = excel_file.filename if excel_file is not None else ids_file.filename
                history.record_analysis(analysis_id, results["label"], rules_filename, results, results["anomalies"])
            except Exception as e:
                # L'historique est un bonus (tableau de bord d'exploitation) --
                # un échec de persistance ne doit jamais invalider une analyse
                # par ailleurs réussie, dont le rapport Excel est déjà sur disque.
                print(f"Warning: could not record analysis history: {str(e)}")
            # Le fichier Excel de règles n'est plus utile une fois l'analyse
            # terminée (contrairement aux IFC par lot, encore nécessaires à la
            # visionneuse 3D via /model/<analysis_id>/<lot_folder>) -- inutile
            # de le garder sur disque.
            for rules_path in (excel_path, ids_path):
                if rules_path:
                    try:
                        os.remove(rules_path)
                    except OSError:
                        pass
            return jsonify(results)
        except Exception as e:
            print(f"Error during analysis: {str(e)}")
            shutil.rmtree(analysis_dir, ignore_errors=True)
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

@app.route('/download_bcf/<analysis_id>')
def download_bcf(analysis_id):
    """Anomalies au format BCF 2.1 (voir checkers_core/bcf.py) : un sujet
    par anomalie, éléments concernés sélectionnés -- à transmettre aux
    maîtres d'œuvre / entreprises, qui l'ouvrent dans leur outil BIM."""
    if not _valid_analysis_id(analysis_id):
        abort(404)
    bcf_path = os.path.join(TEMP_FOLDER, analysis_id, BCF_FILENAME)
    if not os.path.isfile(bcf_path):
        return jsonify({"error": "Export BCF introuvable"}), 404
    return send_file(bcf_path, as_attachment=True, download_name=f"anomalies_{analysis_id[:8]}.bcf",
                     mimetype="application/octet-stream")


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
