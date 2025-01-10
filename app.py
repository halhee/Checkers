from flask import Flask, render_template, request, jsonify
import os
from werkzeug.utils import secure_filename
from flask_cors import CORS
import logging

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

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
            return jsonify({'message': 'File uploaded successfully',
                           'filename': filename}), 200
        except Exception as e:
            logger.error(f'Error saving file: {str(e)}')
            return jsonify({'error': f'Error saving file: {str(e)}'}), 500

if __name__ == '__main__':
    logger.info('Starting Flask application...')
    app.run(host='0.0.0.0', port=8080, debug=True)
