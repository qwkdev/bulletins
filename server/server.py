from bulletin import build

from flask import Flask, request, abort, send_file
from flask_cors import CORS
from pathlib import Path
from docx2pdf import convert
import json
import os
import re

cwd = Path(__file__).parent.resolve()
app = Flask('Bulletins', template_folder=cwd/'templates', static_folder=cwd/'static')
CORS(app)
app.secret_key = os.getenv('key')

API_KEY = os.getenv('API_KEY')

@app.route('/')
def index():
    abort(403)

@app.route('/ping')
def ping():
    return {'success': True}, 200

@app.route('/check')
def check():
    return {'valid': request.args.get('key') == API_KEY}, 200

@app.route('/build', methods=['POST'])
def build_file():
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    data = request.get_json()
    build(
        OUTPUT_PATH='output.docx',
        front_page_margins=(data['front']['top-margin'], data['front']['left-margin']),
        info_data=data['front']['latest-info'],
        info_size=data['front']['latest-info-size'],
        title=data['front']['title'],
        title_size=data['front']['title-size'],
        church_title=data['front']['church-title'],
        church_title_size=data['front']['church-title-size'],
        church_info=data['front']['church-info'],
        church_info_size=data['front']['church-info-size'],
        mass_info=data['front']['mass-info'],
        mass_info_size=data['front']['mass-info-size'],
        data=data['back'],
        readings=data['readings']['readings'],
        reading_margins=(data['readings']['options']['top-margin'], data['readings']['options']['left-margin']),
        reading_heading_spacing=data['readings']['options']['heading-spacing'],
        reading_heading_size=data['readings']['options']['heading-size'],
        copyright_size=data['readings']['options']['copyright-size'],
        copyright_spacing=data['readings']['options']['copyright-spacing'],
        copyright_page=data['readings']['options']['copyright-page'],
        dpa_page=data['readings']['options']['dpa-page'],
    )

    with open(cwd/'latest.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)

    return {'success': True}, 200

@app.route('/get/<file>')
def download(file):
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    if file.endswith('docx'):
        path = cwd/'output.docx'
    elif file.endswith('pdf'):
        convert(cwd/'output.docx', cwd/'output.pdf')
        path = cwd/'output.pdf'

    if not path.exists():
        abort(404)

    return send_file(path, as_attachment=True)

@app.route('/latest')
def latest():
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    with open(cwd/'latest.json', encoding='utf-8') as f:
        data = json.load(f)

    return data

@app.route('/templates')
def templates():
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    filenames = [f.name.split('.json')[0] for f in (cwd/'json').glob('*.json')]

    return {'files': filenames}, 200

def safe_filename(filename: str) -> str:
    return re.sub(r'[^a-zA-Z0-9\-_ ]', '', filename)

@app.route('/template/get/<file>')
def get_template(file):
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    with open(cwd/f'json/{safe_filename(file)}.json', encoding='utf-8') as f:
        data = json.load(f)

    return data

@app.route('/template/save/<file>', methods=['POST'])
def save_template(file):
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    data = request.get_json()
    with open(cwd/f'json/{safe_filename(file)}.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)

    return {'success': True}, 200

@app.route('/template/delete/<file>')
def delete_template(file):
    key = request.args.get('key')
    if key != API_KEY:
        abort(403)

    try:
        os.remove(cwd/f'json/{safe_filename(file)}.json')
    except FileNotFoundError:
        return {'success': False, 'error': 'File Not Found'}, 500

    return {'success': True}, 200