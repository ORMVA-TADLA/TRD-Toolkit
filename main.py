# main.py
from functions import read_mdb_file, table_to_xlsx, groupe_dict, mv1_to_mv2, mv2_to_xlsx, mv1_to_dict
from flask import Flask, render_template, request, jsonify
from flaskwebgui import FlaskUI
import pandas as pd
import json
import os

app = Flask(__name__, template_folder='templates', static_folder='static')

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
FILE_NAME = 'Distribution.mdb'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(app.static_folder):
    os.makedirs(app.static_folder)

# clean old files
for f in os.listdir(UPLOAD_FOLDER):
    os.remove(os.path.join(UPLOAD_FOLDER, f))

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64MB max file size


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/grantt-chart')
def grantt_chart():
    return render_template('grantt-chart.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check if file is in request
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']

        # Check if file is selected
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Check file extension
        if not file.filename.endswith('.mdb'):
            return jsonify({'error': 'Invalid file format. Please upload .mdb'}), 400

        # save file to upload folder
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], FILE_NAME)
        file.save(file_path)

        return jsonify({'message': 'File uploaded successfully', 'filename': file.filename}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/create-grantt-chart')
def create_grantt_chart():
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], FILE_NAME)

    if not os.path.exists(file_path):
        return jsonify({'error': 'No uploaded file found. Please upload a .mdb file first.'}), 400

    # Read MDB file
    db = read_mdb_file(file_path)

    # converting MV1 table to grouped dict
    mv1_dict = mv1_to_dict(db)
    grouped = groupe_dict(mv1_dict)

    # Save dictionary to json
    static_folder = os.path.join(app.static_folder)
    json_path = os.path.join(static_folder, "mv1.json")
    with open(json_path, "w") as f:
        json.dump(grouped, f, indent=4)

    return jsonify({'message': 'successfully'}), 200


@app.route('/extract-mv1', methods=['GET'])
def extract_mv1():
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], FILE_NAME)

        if not os.path.exists(file_path):
            return jsonify({'error': 'No uploaded file found. Please upload a .mdb file first.'}), 400

        # Read MDB file
        db = read_mdb_file(file_path)

        if not db:
            return jsonify({'error': 'Failed to read MDB file.'}), 400

        # Save MV1 table to csv file inside static folder
        table = 'MV1'
        if table not in db:
            return jsonify({'error': f"Table '{table}' not found in the database."}), 400

        static_folder = os.path.join(app.static_folder)

        # add random query to filename to prevent caching
        xlsx_path = os.path.join(
            static_folder, f'MV1 {pd.Timestamp.now().strftime("%Y-%m-%d %H-%M")}.xlsx')
        table_to_xlsx(db, table, xlsx_path)

        # Provide a download URL for the file
        download_url = f"/static/{os.path.basename(xlsx_path)}"

        return jsonify({'message': 'MV1 data extracted successfully', 'download_url': download_url}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/extract-mv2', methods=['GET'])
def extract_mv2():
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], FILE_NAME)

        if not os.path.exists(file_path):
            return jsonify({'error': 'No uploaded file found. Please upload a .mdb file first.'}), 400

        # Read MDB file
        db = read_mdb_file(file_path)

        if not db:
            return jsonify({'error': 'Failed to read MDB file.'}), 400

        table = 'MV1'
        if table not in db:
            return jsonify({'error': f"Table '{table}' not found in the database."}), 400

        # converting MV1 table to grouped dict
        mv1_dict = mv1_to_dict(db)
        grouped = groupe_dict(mv1_dict)

        static_folder = os.path.join(app.static_folder)
        types = ["ALL", "A-N", "C-C"]
        download_urls = []

        for a_type in types:
            # converting grouped dict to mv2
            mv2, TRD_start_hour, TRD_end_hour = mv1_to_mv2(grouped, a_type)
            xlsx_path = os.path.join(
                static_folder, f'MV2 {TRD_start_hour.strftime("%Y-%m-%d")} {TRD_end_hour.strftime("%Y-%m-%d")} - {a_type}.xlsx')
            # saving mv2 to excel file
            mv2_to_xlsx(mv2, TRD_start_hour, TRD_end_hour, xlsx_path)
            # Provide a download URL for the file
            download_url = f"/static/{os.path.basename(xlsx_path)}"
            download_urls.append(download_url)

        return jsonify({'message': 'MV2 Doc created successfully', 'download_urls': download_urls}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    # For development, use Flask's built-in server
    # app.run(debug=True)

    # For desktop app with FlaskWebGUI
    FlaskUI(
        app=app,
        server="flask",
        width=1000,
        height=700,
    ).run()
