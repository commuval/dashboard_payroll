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
def index():
    """Hauptseite"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
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
                flash('Fehler beim Speichern in der Datenbank', 'error')
                return redirect(url_for('index'))
            
        except Exception as e:
            flash(f'Fehler beim Verarbeiten der Datei: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    flash('Ungültiger Dateityp. Nur Excel-Dateien (.xlsx, .xls) sind erlaubt.', 'error')
    return redirect(url_for('index'))

@app.route('/viewer')
def viewer():
    """Excel-Viewer Seite"""
    if 'current_file_id' not in session:
        flash('Keine Datei geladen', 'error')
        return redirect(url_for('index'))
    
    file_id = session['current_file_id']
    filename = session.get('current_filename', 'Unbekannte Datei')
    sheet_names = session.get('sheet_names', [])
    current_sheet = session.get('current_sheet')
    
    # Lade aktuelle Daten
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    
    if sheets_data is None:
        flash('Fehler beim Laden der Datei', 'error')
        return redirect(url_for('index'))
    
    current_data = sheets_data.get(current_sheet, pd.DataFrame()) if current_sheet else pd.DataFrame()
    
    return render_template('viewer.html', 
                         filename=filename,
                         sheet_names=sheet_names,
                         current_sheet=current_sheet,
                         data=current_data.to_dict('records') if not current_data.empty else [],
                         columns=current_data.columns.tolist() if not current_data.empty else [])

@app.route('/change_sheet/<sheet_name>')
def change_sheet(sheet_name):
    """Wechselt zu einem anderen Sheet"""
    if 'current_file_id' not in session:
        return jsonify({'error': 'Keine Datei geladen'})
    
    session['current_sheet'] = sheet_name
    
    # Lade Daten für das neue Sheet
    filename = session.get('current_filename', '')
    sheets_data, _ = db_manager.load_excel_file(file_hash=None, filename=filename)
    
    if sheets_data and sheet_name in sheets_data:
        data = sheets_data[sheet_name]
        return jsonify({
            'success': True,
            'data': data.to_dict('records'),
            'columns': data.columns.tolist()
        })
    
    return jsonify({'error': 'Sheet nicht gefunden'})

@app.route('/sort_praxis')
def sort_praxis():
    """Sortiert Daten nach Praxis"""
    if 'current_file_id' not in session:
        return jsonify({'error': 'Keine Datei geladen'})
    
    file_id = session['current_file_id']
    filename = session.get('current_filename', '')
    
    # Lade aktuelle Daten
    sheets_data, sheet_names = db_manager.load_excel_file(file_hash=None, filename=filename)
    
    if sheets_data is None:
        return jsonify({'error': 'Fehler beim Laden der Daten'})
    
    # Verwende das erste Sheet für die Sortierung
    if sheet_names:
        first_sheet = sheet_names[0]
        df = sheets_data[first_sheet]
        
        # Sortiere nach Praxis
        sorted_data, message = sortiere_nach_praxis(df)
        
        if isinstance(sorted_data, dict):
            # Neue Daten in Datenbank speichern
            db_manager.save_sheets(file_id, sorted_data)
            
            # Session aktualisieren
            session['sheet_names'] = list(sorted_data.keys())
            session['current_sheet'] = list(sorted_data.keys())[0] if sorted_data else None
            
            return jsonify({
                'success': True,
                'message': message,
                'new_sheets': list(sorted_data.keys())
            })
        else:
            return jsonify({'error': message})
    
    return jsonify({'error': 'Keine Daten zum Sortieren gefunden'})

@app.route('/files')
def list_files():
    """Listet alle gespeicherten Dateien"""
    files = db_manager.list_saved_files()
    return render_template('files.html', files=files)

@app.route('/load_file/<int:file_id>')
def load_file(file_id):
    """Lädt eine gespeicherte Datei"""
    # Hier würde die Logik zum Laden einer spezifischen Datei implementiert
    flash('Funktion noch nicht implementiert', 'info')
    return redirect(url_for('files'))

@app.route('/delete_file/<int:file_id>')
def delete_file(file_id):
    """Löscht eine Datei"""
    if db_manager.delete_file(file_id):
        flash('Datei erfolgreich gelöscht', 'success')
    else:
        flash('Fehler beim Löschen der Datei', 'error')
    
    return redirect(url_for('files'))

@app.route('/backup/<int:file_id>')
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

if __name__ == '__main__':
    # Erstelle Datenbank-Tabellen
    db_manager.create_tables()
    
    # Starte Flask-App
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False) 