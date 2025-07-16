from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, session
import pandas as pd
import numpy as np
import os
import sys
import tempfile
import shutil
import json
import pickle
import glob
import io
from pathlib import Path
import hashlib
from datetime import datetime, timedelta
from werkzeug.utils import secure_filename
from database import DatabaseManager
from sqlalchemy import text
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-here')

# Konfiguration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Erstelle Upload-Ordner
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Initialisiere Datenbank
db_manager = DatabaseManager()

# Passwort für Zugriffsschutz
APP_PASSWORD = 'Q*%(1v87q"cI'

# Decorator für Passwortschutz

def passwort_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('auth_ok') and request.endpoint != 'passwort':
            return redirect(url_for('passwort'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/passwort', methods=['GET', 'POST'])
def passwort():
    if request.method == 'POST':
        eingabe = request.form.get('passwort', '')
        if eingabe == APP_PASSWORD:
            session['auth_ok'] = True
            return redirect(url_for('index'))
        else:
            flash('Falsches Passwort!', 'error')
    return render_template('passwort.html')

# Alle Routen (außer passwort) schützen
# for rule in list(app.url_map.iter_rules()):
#     if rule.endpoint not in ('static', 'passwort'):
#         view_func = app.view_functions[rule.endpoint]
#         app.view_functions[rule.endpoint] = passwort_required(view_func)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_file_hash(file_content):
    """Berechnet einen Hash für die Datei"""
    return hashlib.md5(file_content).hexdigest()

def load_excel_data(file_path):
    """Lädt Excel-Datei und gibt DataFrame zurück"""
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        sheets_data = {}
        for sheet in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            sheets_data[sheet] = df
        
        return sheets_data, sheet_names
    except Exception as e:
        return None, []

def clean_value(value):
    """Bereinigt Werte für bessere Anzeige"""
    if pd.isna(value):
        return ""
    elif isinstance(value, (int, float)):
        if pd.isna(value):
            return ""
        return str(value)
    else:
        return str(value).strip()

def sortiere_nach_praxis(df):
    """Verteilt komplette Zeilen basierend auf Spalte B (Praxis) zu jeweiligen Praxis-Sheets"""
    try:
        if len(df.columns) < 2:
            return df, "Nicht genügend Spalten! Spalte B (Praxis-Spalte) nicht gefunden."

        praxis_col = df.columns[1]  # Spalte B (zweite Spalte)

        # Leere Werte in Spalte B prüfen
        non_empty_rows = df[praxis_col].notna() & (df[praxis_col].astype(str).str.strip() != '')

        if not non_empty_rows.any():
            return df, f"Spalte B '{praxis_col}' enthält keine gültigen Praxis-Namen!"

        # Nur Zeilen mit gültigen Praxis-Namen verarbeiten
        valid_df = df[non_empty_rows].copy()

        # Nach Praxis gruppieren
        praxis_groups = valid_df.groupby(praxis_col, dropna=False)

        # Bestehende Daten beibehalten
        updated_data = {}

        # Für jede Praxis komplette Zeilen zu eigenem Sheet hinzufügen
        sheets_created = 0
        total_rows_added = 0

        for praxis_name, praxis_rows in praxis_groups:
            praxis_name_clean = str(praxis_name).strip()

            if not praxis_name_clean:
                continue

            # Sheet-Name erstellen
            sheet_name = f"{praxis_name_clean}"
            sheet_name = sheet_name.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("[", "_").replace("]", "_")

            # Komplette Zeilen übernehmen
            complete_rows = praxis_rows.copy()

            if sheet_name in updated_data:
                # Zu bestehendem Sheet hinzufügen
                existing_data = updated_data[sheet_name]

                # Duplikate vermeiden
                existing_hashes = set()
                for _, row in existing_data.iterrows():
                    row_hash = hash(tuple(row.astype(str)))
                    existing_hashes.add(row_hash)

                new_rows = []
                duplicates_found = 0

                for _, row in complete_rows.iterrows():
                    row_hash = hash(tuple(row.astype(str)))
                    if row_hash not in existing_hashes:
                        new_rows.append(row)
                        existing_hashes.add(row_hash)
                    else:
                        duplicates_found += 1

                if new_rows:
                    new_df = pd.DataFrame(new_rows)
                    updated_data[sheet_name] = pd.concat([existing_data, new_df], ignore_index=True)
                    total_rows_added += len(new_rows)

                if duplicates_found > 0:
                    print(f"Sheet '{sheet_name}': {duplicates_found} Duplikate übersprungen")
            else:
                # Neues Sheet erstellen
                updated_data[sheet_name] = complete_rows
                sheets_created += 1
                total_rows_added += len(complete_rows)

        success_message = f"Verteilung abgeschlossen: {sheets_created} neue Sheets erstellt, {total_rows_added} Zeilen hinzugefügt"
        return updated_data, success_message

    except Exception as e:
        return df, f"Fehler bei der Verteilung: {str(e)}"

@app.route('/')
@passwort_required
def index():
    """Hauptseite"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@passwort_required
def upload_file():
    """Datei-Upload-Handler"""
    if 'file' not in request.files:
        flash('Keine Datei ausgewählt', 'error')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('Keine Datei ausgewählt', 'error')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        try:
            # Datei temporär speichern
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Dateiinhalt lesen für Hash-Berechnung
            with open(file_path, 'rb') as f:
                file_content = f.read()
            
            file_hash = calculate_file_hash(file_content)
            
            # Excel-Daten laden
            sheets_data, sheet_names = load_excel_data(file_path)
            
            if sheets_data is None:
                flash('Fehler beim Laden der Excel-Datei', 'error')
                return redirect(url_for('index'))
            
            # In Datenbank speichern
            metadata = {
                'sheet_count': len(sheet_names),
                'upload_date': datetime.now().isoformat()
            }
            
            file_id = db_manager.save_excel_file(filename, file_hash, sheets_data, metadata)
            
            if file_id:
                # Session-Daten setzen
                session['current_file_id'] = file_id
                session['current_filename'] = filename
                session['sheet_names'] = sheet_names
                session['current_sheet'] = sheet_names[0] if sheet_names else None
                
                flash(f'Datei "{filename}" erfolgreich hochgeladen', 'success')
                return redirect(url_for('viewer'))
            else:
                flash('Fehler beim Speichern in der Datenbank - Tabellen werden erstellt...', 'warning')
                # Versuche Tabellen zu erstellen und nochmal zu speichern
                try:
                    db_manager.create_tables()
                    file_id = db_manager.save_excel_file(filename, file_hash, sheets_data, metadata)
                    if file_id:
                        session['current_file_id'] = file_id
                        session['current_filename'] = filename
                        session['sheet_names'] = sheet_names
                        session['current_sheet'] = sheet_names[0] if sheet_names else None
                        flash(f'Datei "{filename}" erfolgreich hochgeladen (Tabellen erstellt)', 'success')
                        return redirect(url_for('viewer'))
                    else:
                        flash('Fehler beim Speichern in der Datenbank nach Tabellenerstellung', 'error')
                        return redirect(url_for('index'))
                except Exception as e:
                    flash(f'Datenbankfehler: {str(e)}', 'error')
                    return redirect(url_for('index'))
            
        except Exception as e:
            flash(f'Fehler beim Verarbeiten der Datei: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    flash('Ungültiger Dateityp. Nur Excel-Dateien (.xlsx, .xls) sind erlaubt.', 'error')
    return redirect(url_for('index'))

@app.route('/viewer')
@passwort_required
def viewer():
    if 'current_file_id' not in session:
        flash('Keine Datei geladen', 'error')
        return redirect(url_for('index'))

    file_id = session['current_file_id']
    filename = session.get('current_filename', 'Unbekannte Datei')
    # Sheet aus GET-Parameter oder Session
    current_sheet = request.args.get('sheet') or session.get('current_sheet')

    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or not sheets_data:
        flash('Fehler beim Laden der Datei', 'error')
        return redirect(url_for('index'))

    # Nur Original-Sheets anzeigen (wie in der Excel-Datei)
    original_sheet_names = session.get('sheet_names', [])
    sheet_names = [s for s in original_sheet_names if s in sheets_data]

    if not current_sheet or current_sheet not in sheet_names:
        current_sheet = sheet_names[0] if sheet_names else None
    session['current_sheet'] = current_sheet

    current_data = sheets_data.get(current_sheet, pd.DataFrame()) if current_sheet else pd.DataFrame()

    return render_template('viewer.html',
        filename=filename,
        sheet_names=sheet_names,
        current_sheet=current_sheet,
        data=current_data.fillna('').to_dict('records') if not current_data.empty else [],
        columns=current_data.columns.tolist() if not current_data.empty else []
    )

@app.route('/change_sheet/<sheet_name>')
@passwort_required
def change_sheet(sheet_name):
    print("Sheet-Wechsel zu:", sheet_name)
    print("Sheet-Wechsel zu (repr):", repr(sheet_name))
    if 'current_file_id' not in session:
        return jsonify({'error': 'Keine Datei geladen'})

    session['current_sheet'] = sheet_name
    filename = session.get('current_filename', '')
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    print("Verfügbare Sheets:", list(sheets_data.keys()))
    # Sheet-Namen robust vergleichen (strip, lower)
    sheet_map = {s.strip().lower(): s for s in sheets_data.keys()}
    key = sheet_name.strip().lower()
    if key in sheet_map:
        real_name = sheet_map[key]
        data = sheets_data[real_name]
        print("Daten für Sheet", real_name, ":", data.head() if hasattr(data, 'head') else data)
        return jsonify({
            'success': True,
            'data': data.to_dict('records'),
            'columns': data.columns.tolist()
        })
    print("Sheet nicht gefunden:", sheet_name)
    return jsonify({'error': 'Sheet nicht gefunden'})

@app.route('/sort_praxis')
@passwort_required
def sort_praxis():
    if 'current_file_id' not in session:
        return jsonify({'error': 'Keine Datei geladen'})

    file_id = session['current_file_id']
    filename = session.get('current_filename', '')
    current_sheet = session.get('current_sheet')

    # Lade aktuelle Daten
    sheets_data, sheet_names = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None:
        return jsonify({'error': 'Fehler beim Laden der Daten'})

    # Immer das Sheet 'Dashboard' als Quelle für die Sortierung verwenden, falls vorhanden
    original_sheet_names = session.get('sheet_names', [])
    dashboard_sheet = None
    for s in original_sheet_names:
        if s.strip().lower() == 'dashboard':
            dashboard_sheet = s
            break
    if dashboard_sheet and dashboard_sheet in sheets_data:
        df = sheets_data[dashboard_sheet]
        current_sheet = dashboard_sheet
    elif original_sheet_names and original_sheet_names[0] in sheets_data:
        df = sheets_data[original_sheet_names[0]]
        current_sheet = original_sheet_names[0]
    else:
        return jsonify({'error': 'Kein gültiges Sheet zum Sortieren gefunden'})
    print(f"[DEBUG] Spalten Dashboard: {list(df.columns)}")
    # Sortiere nach Praxis (Spalte B)
    sorted_data, message = sortiere_nach_praxis(df)
    if isinstance(sorted_data, dict):
        print(f"[DEBUG] sortiere_nach_praxis: { {k: v.shape for k, v in sorted_data.items()} }")
        for praxis_sheet, praxis_df in sorted_data.items():
            print(f"[DEBUG] Spalten {praxis_sheet}: {list(praxis_df.columns)}")
        # Bestehende Sheets aus der DB laden
        all_sheets_data, all_sheet_names = db_manager.load_excel_file(file_hash=None, filename=filename)
        # Original-Sheet immer übernehmen
        sheets_to_save = {current_sheet: df}
        # Praxis-Sheets ergänzen statt überschreiben
        for praxis_sheet, praxis_df in sorted_data.items():
            if praxis_sheet in all_sheets_data:
                praxis_df = praxis_df.reset_index(drop=True)
                alt_df = all_sheets_data[praxis_sheet].reset_index(drop=True)
                columns = list(praxis_df.columns)
                alt_df = alt_df.reindex(columns=columns)
                # Robuste Duplikat-Prüfung: Hash aus allen Spaltenwerten (stripped, NaN als '')
                def row_hash(row):
                    return '|'.join([(str(x).strip() if pd.notna(x) else '') for x in row])
                alt_hashes = set(row_hash(row) for _, row in alt_df.iterrows())
                new_rows = [row for _, row in praxis_df.iterrows() if row_hash(row) not in alt_hashes]
                if new_rows:
                    new_df = pd.DataFrame(new_rows, columns=columns)
                    merged = pd.concat([new_df, alt_df], ignore_index=True)
                else:
                    merged = alt_df
                sheets_to_save[praxis_sheet] = merged
            else:
                sheets_to_save[praxis_sheet] = praxis_df.reset_index(drop=True)
        for k, v in sheets_to_save.items():
            print(f"[DEBUG] Sheet to save: {k}, rows: {v.shape}")
        db_manager.save_sheets(file_id, sheets_to_save)
        # Nach dem Speichern: vollständige Sheet-Liste neu laden
        all_sheets_data, all_sheet_names = db_manager.load_excel_file(file_hash=None, filename=filename)
        # session['sheet_names'] NICHT überschreiben, damit die Original-Reihenfolge erhalten bleibt
        if current_sheet in all_sheet_names:
            session['current_sheet'] = current_sheet
        else:
            session['current_sheet'] = all_sheet_names[0] if all_sheet_names else None
        return jsonify({
            'success': True,
            'message': message,
            'new_sheets': list(sorted_data.keys()),
            'all_sheets': all_sheet_names
        })
    else:
        return jsonify({'error': message})

@app.route('/files')
@passwort_required
def list_files():
    """Listet alle gespeicherten Dateien"""
    files = db_manager.list_saved_files()
    return render_template('files.html', files=files)

@app.route('/load_file/<int:file_id>')
@passwort_required
def load_file(file_id):
    """Lädt eine gespeicherte Datei und setzt die Session-Variablen"""
    # Datei-Metadaten und Name aus der Datenbank holen
    file_info = next((f for f in db_manager.list_saved_files() if f['id'] == file_id), None)
    if not file_info:
        flash('Datei nicht gefunden', 'error')
        return redirect(url_for('list_files'))
    filename = file_info['filename'] if isinstance(file_info, dict) else getattr(file_info, 'filename', None)
    if not filename:
        flash('Dateiname nicht gefunden', 'error')
        return redirect(url_for('list_files'))
    # Excel-Daten laden
    sheets_data, sheet_names = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or not sheet_names:
        flash('Fehler beim Laden der Datei', 'error')
        return redirect(url_for('list_files'))
    # Session-Variablen setzen
    session['current_file_id'] = file_id
    session['current_filename'] = filename
    session['sheet_names'] = sheet_names
    session['current_sheet'] = sheet_names[0] if sheet_names else None
    flash(f'Datei "{filename}" geladen', 'success')
    return redirect(url_for('viewer'))

@app.route('/delete_file/<int:file_id>')
@passwort_required
def delete_file(file_id):
    """Löscht eine Datei"""
    if db_manager.delete_file(file_id):
        flash('Datei erfolgreich gelöscht', 'success')
    else:
        flash('Fehler beim Löschen der Datei', 'error')
    
    return redirect(url_for('list_files'))

@app.route('/backup/<int:file_id>')
@passwort_required
def create_backup(file_id):
    """Erstellt ein Backup einer Datei"""
    filename = session.get('current_filename', '')
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    
    if sheets_data:
        backup_id = db_manager.create_backup(file_id, sheets_data, 'manual')
        if backup_id:
            flash(f'Backup erfolgreich erstellt (ID: {backup_id})', 'success')
        else:
            flash('Fehler beim Erstellen des Backups', 'error')
    else:
        flash('Keine Daten zum Sichern gefunden', 'error')
    
    return redirect(url_for('viewer'))

@app.route('/setup-database')
@passwort_required
def setup_database():
    """Erstellt die Datenbank-Tabellen"""
    try:
        success = db_manager.create_tables()
        if success:
            flash('Datenbank-Tabellen erfolgreich erstellt!', 'success')
        else:
            flash('Fehler beim Erstellen der Tabellen', 'error')
    except Exception as e:
        flash(f'Datenbankfehler: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/test-db')
@passwort_required
def test_database():
    """Testet die Datenbankverbindung"""
    try:
        if db_manager.engine:
            with db_manager.engine.connect() as conn:
                result = conn.execute(text("SELECT 1 as test"))
                return jsonify({
                    'status': 'success',
                    'message': 'Datenbankverbindung erfolgreich',
                    'test_result': result.fetchone()[0]
                    })
        else:
            return jsonify({
                'status': 'error',
                'message': 'Keine Datenbankverbindung verfügbar'
                })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Datenbankfehler: {str(e)}',
            'error_type': type(e).__name__
            })

@app.route('/update_cell', methods=['POST'])
@passwort_required
def update_cell():
    data = request.get_json()
    sheet = data.get('sheet')
    row = data.get('row')
    column = data.get('column')
    value = data.get('value')
    filename = session.get('current_filename', '')
    file_id = session.get('current_file_id')
    if not (sheet and column and filename and file_id is not None):
        return jsonify({'success': False, 'error': 'Ungültige Anfrage'}), 400
    # Lade aktuelle Daten
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or sheet not in sheets_data:
        return jsonify({'success': False, 'error': 'Sheet nicht gefunden'}), 404
    df = sheets_data[sheet]
    try:
        row = int(row)
        if row < 0 or row >= len(df):
            return jsonify({'success': False, 'error': 'Ungültige Zeilennummer'}), 400
        df.at[row, column] = value
        sheets_data[sheet] = df
        db_manager.save_sheets(file_id, {sheet: df})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/delete_row', methods=['POST'])
@passwort_required
def delete_row():
    data = request.get_json()
    sheet = data.get('sheet')
    row = data.get('row')
    filename = session.get('current_filename', '')
    file_id = session.get('current_file_id')
    if not (sheet and filename and file_id is not None):
        return jsonify({'success': False, 'error': 'Ungültige Anfrage'}), 400
    # Lade aktuelle Daten
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or sheet not in sheets_data:
        return jsonify({'success': False, 'error': 'Sheet nicht gefunden'}), 404
    df = sheets_data[sheet]
    try:
        row = int(row)
        if row < 0 or row >= len(df):
            return jsonify({'success': False, 'error': 'Ungültige Zeilennummer'}), 400
        df = df.drop(df.index[row]).reset_index(drop=True)
        sheets_data[sheet] = df
        db_manager.save_sheets(file_id, {sheet: df})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/color_row', methods=['POST'])
@passwort_required
def color_row():
    data = request.get_json()
    sheet = data.get('sheet')
    row = data.get('row')
    color = data.get('color')
    filename = session.get('current_filename', '')
    file_id = session.get('current_file_id')
    if not (sheet and filename and file_id is not None):
        return jsonify({'success': False, 'error': 'Ungültige Anfrage'}), 400
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or sheet not in sheets_data:
        return jsonify({'success': False, 'error': 'Sheet nicht gefunden'}), 404
    df = sheets_data[sheet]
    try:
        row = int(row)
        if row < 0 or row >= len(df):
            return jsonify({'success': False, 'error': 'Ungültige Zeilennummer'}), 400
        if '_row_color' not in df.columns:
            df['_row_color'] = ''
        if color == 'none':
            df.at[row, '_row_color'] = ''
        else:
            df.at[row, '_row_color'] = color
        sheets_data[sheet] = df
        db_manager.save_sheets(file_id, {sheet: df})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/comment_row', methods=['POST'])
@passwort_required
def comment_row():
    data = request.get_json()
    sheet = data.get('sheet')
    row = data.get('row')
    comment = data.get('comment', '')
    filename = session.get('current_filename', '')
    file_id = session.get('current_file_id')
    if not (sheet and filename and file_id is not None):
        return jsonify({'success': False, 'error': 'Ungültige Anfrage'}), 400
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    if sheets_data is None or sheet not in sheets_data:
        return jsonify({'success': False, 'error': 'Sheet nicht gefunden'}), 404
    df = sheets_data[sheet]
    try:
        row = int(row)
        if row < 0 or row >= len(df):
            return jsonify({'success': False, 'error': 'Ungültige Zeilennummer'}), 400
        if '_row_comment' not in df.columns:
            df['_row_comment'] = ''
        df.at[row, '_row_comment'] = comment
        sheets_data[sheet] = df
        db_manager.save_sheets(file_id, {sheet: df})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    # Erstelle Datenbank-Tabellen
    db_manager.create_tables()
    
    # Starte Flask-App
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True) 