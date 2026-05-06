import os
import uuid
from flask import Flask, request, jsonify, session, send_file, send_from_directory
from flask_cors import CORS
from pathlib import Path
from werkzeug.utils import secure_filename
import shutil

from app.services import (
    initialize_database, authenticate, change_password, verify_current_password,
    list_users, create_user, update_user_status, update_user_role, reset_user_password,
    create_record, update_record, delete_record, update_record_status, list_records, get_record, export_records,
    import_excel_data, get_fields, get_visible_fields,
    create_field, update_field, delete_field,
    get_fuid_header_config, save_fuid_header_config,
    get_fuid_detail_mapping, update_fuid_detail_mapping,
    generate_fuid,
    get_rotulo_carpeta_config, save_rotulo_carpeta_config,
    get_rotulo_carpeta_mapping, update_rotulo_carpeta_mapping,
    generate_rotulo_carpeta,
    get_rotulo_caja_config, save_rotulo_caja_config,
    get_rotulo_caja_mapping, update_rotulo_caja_mapping,
    generate_rotulo_caja,
    get_audit_rows, get_dashboard_stats,
    list_records_paged, save_attachment, get_attachments, delete_attachment,
    db_cursor
)

app = Flask(__name__, static_folder='static')
app.secret_key = 'archivo-documental-secret-key-123' # In production, use a secure key
CORS(app)

# Initialize database on startup
initialize_database()

UPLOAD_FOLDER = Path('uploads')
UPLOAD_FOLDER.mkdir(exist_ok=True)

SALIDAS_FOLDER = Path('salidas')
SALIDAS_FOLDER.mkdir(exist_ok=True)

# Helper to check login
def is_logged_in():
    return 'user' in session

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    username = data.get('username')
    password = data.get('password')
    user, err = authenticate(username, password)
    if err:
        return jsonify({'error': err}), 401
    session['user'] = user
    return jsonify({'user': user})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.pop('user', None)
    return jsonify({'success': True})

@app.route('/api/me', methods=['GET'])
def get_me():
    if not is_logged_in():
        return jsonify({'error': 'Not logged in'}), 401
    return jsonify({'user': session['user']})

@app.route('/api/records', methods=['GET'])
def get_records():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    search = request.args.get('search', '')
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    return jsonify(list_records_paged(search, page, per_page))

@app.route('/api/records', methods=['POST'])
def add_record():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    payload = request.json
    record_id = create_record(session['user']['username'], payload)
    return jsonify({'id': record_id})

@app.route('/api/records/<int:record_id>', methods=['GET'])
def get_single_record(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_record(record_id))

@app.route('/api/records/<int:record_id>', methods=['PUT'])
def edit_record(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    payload = request.json
    update_record(session['user']['username'], record_id, payload)
    return jsonify({'success': True})

@app.route('/api/records/<int:record_id>', methods=['DELETE'])
def remove_record(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    delete_record(session['user']['username'], record_id)
    return jsonify({'success': True})

@app.route('/api/records/<int:record_id>/status', methods=['PUT'])
def edit_record_status(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_record_status(session['user']['username'], record_id, data['active'])
    return jsonify({'success': True})

@app.route('/api/fields', methods=['GET'])
def get_all_fields():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_fields())

@app.route('/api/fields', methods=['POST'])
def add_field():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    try:
        create_field(session['user']['username'], data['column_name'], data['display_name'], data['visible'], data['display_order'], data['default_value'])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/fields/<int:field_id>', methods=['PUT'])
def edit_field(field_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    try:
        update_field(session['user']['username'], field_id, data['column_name'], data['display_name'], data['visible'], data['display_order'], data['default_value'])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/fields/<int:field_id>', methods=['DELETE'])
def remove_field(field_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    delete_field(session['user']['username'], field_id)
    return jsonify({'success': True})

@app.route('/api/import', methods=['POST'])
def import_excel():
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    if 'file' not in request.files: return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No selected file'}), 400
    
    filename = secure_filename(file.filename)
    filepath = UPLOAD_FOLDER / filename
    file.save(str(filepath))
    
    try:
        result = import_excel_data(session['user']['username'], str(filepath))
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if filepath.exists():
            os.remove(str(filepath))

@app.route('/api/export', methods=['GET'])
def export_all():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    fmt = request.args.get('format', 'xlsx')
    output_path = SALIDAS_FOLDER / f"export_{uuid.uuid4().hex}.{fmt}"
    try:
        final_path = export_records(str(output_path))
        return send_file(os.path.abspath(final_path), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/fuid/header', methods=['GET', 'POST'])
def config_fuid_header():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    if request.method == 'POST':
        save_fuid_header_config(session['user']['username'], request.json)
        return jsonify({'success': True})
    return jsonify(get_fuid_header_config())

@app.route('/api/config/fuid/mapping', methods=['GET'])
def config_fuid_mapping():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_fuid_detail_mapping())

@app.route('/api/config/fuid/mapping/<int:mapping_id>', methods=['PUT'])
def edit_fuid_mapping(mapping_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_fuid_detail_mapping(session['user']['username'], mapping_id, data['mapping_type'], data['mapping_value'])
    return jsonify({'success': True})

@app.route('/api/generate/fuid', methods=['GET'])
def generate_fuid_api():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    search = request.args.get('search', '')
    fmt = request.args.get('format', 'Excel')
    try:
        final_path = generate_fuid(session['user']['username'], search, fmt, str(SALIDAS_FOLDER))
        return send_file(os.path.abspath(final_path), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/rotulo-carpeta', methods=['GET', 'POST'])
def config_rotulo_carpeta():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    if request.method == 'POST':
        save_rotulo_carpeta_config(session['user']['username'], request.json)
        return jsonify({'success': True})
    return jsonify(get_rotulo_carpeta_config())

@app.route('/api/config/rotulo-carpeta/mapping', methods=['GET'])
def config_rotulo_carpeta_mapping():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_rotulo_carpeta_mapping())

@app.route('/api/config/rotulo-carpeta/mapping/<int:mapping_id>', methods=['PUT'])
def edit_rotulo_carpeta_mapping(mapping_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_rotulo_carpeta_mapping(session['user']['username'], mapping_id, data['mapping_type'], data['mapping_value'])
    return jsonify({'success': True})

@app.route('/api/generate/rotulo-carpeta', methods=['GET'])
def generate_rotulo_carpeta_api():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    desde = request.args.get('desde', '')
    hasta = request.args.get('hasta', '')
    fmt = request.args.get('format', 'Word')
    try:
        final_path = generate_rotulo_carpeta(session['user']['username'], desde, hasta, fmt, str(SALIDAS_FOLDER))
        return send_file(os.path.abspath(final_path), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/rotulo-caja', methods=['GET', 'POST'])
def config_rotulo_caja():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    if request.method == 'POST':
        save_rotulo_caja_config(session['user']['username'], request.json)
        return jsonify({'success': True})
    return jsonify(get_rotulo_caja_config())

@app.route('/api/config/rotulo-caja/mapping', methods=['GET'])
def config_rotulo_caja_mapping():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_rotulo_caja_mapping())

@app.route('/api/config/rotulo-caja/mapping/<int:mapping_id>', methods=['PUT'])
def edit_rotulo_caja_mapping_api(mapping_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_rotulo_caja_mapping(session['user']['username'], mapping_id, data['mapping_type'], data['mapping_value'])
    return jsonify({'success': True})

@app.route('/api/generate/rotulo-caja', methods=['GET'])
def generate_rotulo_caja_api():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    caja = request.args.get('caja', '')
    desde = request.args.get('desde', '')
    hasta = request.args.get('hasta', '')
    fmt = request.args.get('format', 'Word')
    try:
        final_path = generate_rotulo_caja(session['user']['username'], caja, desde, hasta, fmt, str(SALIDAS_FOLDER))
        return send_file(os.path.abspath(final_path), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/users', methods=['GET'])
def get_users_list():
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(list_users())

@app.route('/api/users', methods=['POST'])
def add_user():
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    try:
        create_user(session['user']['username'], data['username'], data['full_name'], data['password'], data['role'])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/users/<int:user_id>/status', methods=['PUT'])
def edit_user_status(user_id):
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_user_status(session['user']['username'], user_id, data['active'])
    return jsonify({'success': True})
@app.route('/api/users/<int:user_id>/role', methods=['PUT'])
def edit_user_role(user_id):
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    update_user_role(session['user']['username'], user_id, data['role'])
    return jsonify({'success': True})

@app.route('/api/users/<int:user_id>/reset-password', methods=['POST'])
def reset_password_api(user_id):
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    reset_user_password(session['user']['username'], user_id, data['password'])
    return jsonify({'success': True})

@app.route('/api/change-password', methods=['POST'])
def change_pwd():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    data = request.json
    if not verify_current_password(session['user']['id'], data['current_password']):
        return jsonify({'error': 'Clave actual incorrecta'}), 400
    change_password(session['user']['id'], data['new_password'])
    return jsonify({'success': True})

@app.route('/api/stats', methods=['GET'])
def get_stats():
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_dashboard_stats())

@app.route('/api/audit-logs', methods=['GET'])
def get_audits():
    if not is_logged_in() or session['user']['role'] != 'admin': return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_audit_rows())

@app.route('/api/records/<int:record_id>/attachments', methods=['GET'])
def get_record_attachments(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    return jsonify(get_attachments(record_id))

@app.route('/api/records/<int:record_id>/attachments', methods=['POST'])
def add_attachment(record_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No selected file'}), 400
    
    filename = secure_filename(file.filename)
    record_upload_dir = UPLOAD_FOLDER / str(record_id)
    record_upload_dir.mkdir(exist_ok=True)
    filepath = record_upload_dir / filename
    file.save(str(filepath))
    
    save_attachment(session['user']['username'], record_id, filename, filepath)
    return jsonify({'success': True})

@app.route('/api/attachments/<int:attachment_id>', methods=['DELETE'])
def remove_attachment(attachment_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    delete_attachment(session['user']['username'], attachment_id)
    return jsonify({'success': True})

@app.route('/api/attachments/<int:attachment_id>/download')
def download_attachment(attachment_id):
    if not is_logged_in(): return jsonify({'error': 'Unauthorized'}), 401
    with db_cursor() as (_, cur):
        cur.execute("SELECT filename, filepath FROM attachments WHERE id = ?", (attachment_id,))
        row = cur.fetchone()
        if not row: return "Not found", 404
        return send_file(os.path.abspath(row["filepath"]), as_attachment=True, download_name=row["filename"])

if __name__ == '__main__':
    app.run(debug=True, port=5000)
