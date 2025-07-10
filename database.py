import os
import pandas as pd
import sqlalchemy as sa
from sqlalchemy import create_engine, text, MetaData, Table, Column, String, DateTime, Integer, Float, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import json
from dotenv import load_dotenv

# Lade Umgebungsvariablen
load_dotenv()

Base = declarative_base()

class DatabaseManager:
    def __init__(self):
        self.engine = None
        self.Session = None
        self.metadata = MetaData()
        self.setup_connection()
    
    def setup_connection(self):
        """Stellt die Datenbankverbindung her"""
        try:
            # PostgreSQL-Verbindungsdaten aus Umgebungsvariablen
            db_host = os.getenv('DB_HOST', 'localhost')
            db_port = os.getenv('DB_PORT', '5432')
            db_name = os.getenv('DB_NAME', 'excel_viewer')
            db_user = os.getenv('DB_USER', 'postgres')
            db_password = os.getenv('DB_PASSWORD', '')
            
            print(f"Verbinde zu Datenbank: {db_host}:{db_port}/{db_name} als {db_user}")
            
            # Prüfe ob alle erforderlichen Variablen gesetzt sind
            if not db_password:
                print("WARNUNG: DB_PASSWORD ist nicht gesetzt!")
                self.engine = None
                return
            
            connection_string = f"postgresql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"
            
            self.engine = create_engine(connection_string)
            self.Session = sessionmaker(bind=self.engine)
            
            # Teste Verbindung
            with self.engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            
            print("PostgreSQL-Verbindung erfolgreich hergestellt")
            
        except Exception as e:
            print(f"Fehler bei der Datenbankverbindung: {str(e)}")
            print(f"Fehler-Typ: {type(e).__name__}")
            print("Stellen Sie sicher, dass PostgreSQL läuft und die Verbindungsdaten korrekt sind.")
            self.engine = None
    
    def create_tables(self):
        """Erstellt die notwendigen Tabellen"""
        if not self.engine:
            print("Keine Datenbankverbindung verfügbar")
            return False
        
        try:
            print("Erstelle Tabellen...")
            
            # Tabelle für Excel-Dateien
            excel_files = Table('excel_files', self.metadata,
                Column('id', Integer, primary_key=True),
                Column('filename', String(255), nullable=False),
                Column('upload_date', DateTime, default=datetime.utcnow),
                Column('file_hash', String(64), unique=True),
                Column('metadata', String(1000))  # JSON für zusätzliche Metadaten
            )
            
            # Tabelle für Sheets
            sheets = Table('sheets', self.metadata,
                Column('id', Integer, primary_key=True),
                Column('excel_file_id', Integer, sa.ForeignKey('excel_files.id')),
                Column('sheet_name', String(255), nullable=False),
                Column('data_json', String(10000000)),  # JSON für Sheet-Daten
                Column('last_modified', DateTime, default=datetime.utcnow)
            )
            
            # Tabelle für Backups
            backups = Table('backups', self.metadata,
                Column('id', Integer, primary_key=True),
                Column('excel_file_id', Integer, sa.ForeignKey('excel_files.id')),
                Column('backup_date', DateTime, default=datetime.utcnow),
                Column('backup_data_json', String(10000000)),
                Column('backup_type', String(50))  # 'manual', 'auto'
            )
            
            print("Tabellen-Definitionen erstellt, erstelle in Datenbank...")
            self.metadata.create_all(self.engine)
            print("Tabellen erfolgreich erstellt!")
            return True
            
        except Exception as e:
            print(f"Fehler beim Erstellen der Tabellen: {str(e)}")
            print(f"Fehler-Typ: {type(e).__name__}")
            return False
    
    def save_excel_file(self, filename, file_hash, sheets_data, metadata=None):
        """Speichert eine Excel-Datei in der Datenbank"""
        if not self.engine:
            print("Keine Datenbankverbindung verfügbar")
            return None
        
        try:
            print(f"Versuche Datei '{filename}' zu speichern...")
            with self.Session() as session:
                # Prüfe ob Datei bereits existiert
                existing_file = session.execute(
                    text("SELECT id FROM excel_files WHERE file_hash = :hash"),
                    {"hash": file_hash}
                ).fetchone()
                
                if existing_file:
                    file_id = existing_file[0]
                    print(f"Datei '{filename}' bereits in Datenbank vorhanden")
                else:
                    # Neue Datei einfügen
                    print("Füge neue Datei in Datenbank ein...")
                    result = session.execute(
                        text("""
                            INSERT INTO excel_files (filename, file_hash, metadata)
                            VALUES (:filename, :hash, :metadata)
                            RETURNING id
                        """),
                        {
                            "filename": filename,
                            "hash": file_hash,
                            "metadata": json.dumps(metadata) if metadata else None
                        }
                    )
                    file_id = result.fetchone()[0]
                    session.commit()
                    print(f"Datei '{filename}' erfolgreich in Datenbank gespeichert")
                
                # Sheets speichern
                print("Speichere Sheets...")
                self.save_sheets(file_id, sheets_data)
                
                return file_id
                
        except Exception as e:
            print(f"Fehler beim Speichern der Datei: {str(e)}")
            print(f"Fehler-Typ: {type(e).__name__}")
            return None
    
    def save_sheets(self, file_id, sheets_data):
        """Speichert alle Sheets einer Excel-Datei"""
        if not self.engine:
            return
        
        try:
            with self.Session() as session:
                for sheet_name, df in sheets_data.items():
                    # DataFrame zu JSON konvertieren
                    data_json = df.to_json(orient='records', date_format='iso')
                    
                    # Prüfe ob Sheet bereits existiert
                    existing_sheet = session.execute(
                        text("SELECT id FROM sheets WHERE excel_file_id = :file_id AND sheet_name = :sheet_name"),
                        {"file_id": file_id, "sheet_name": sheet_name}
                    ).fetchone()
                    
                    if existing_sheet:
                        # Sheet aktualisieren
                        session.execute(
                            text("""
                                UPDATE sheets 
                                SET data_json = :data_json, last_modified = :now
                                WHERE excel_file_id = :file_id AND sheet_name = :sheet_name
                            """),
                            {
                                "data_json": data_json,
                                "now": datetime.utcnow(),
                                "file_id": file_id,
                                "sheet_name": sheet_name
                            }
                        )
                    else:
                        # Neues Sheet einfügen
                        session.execute(
                            text("""
                                INSERT INTO sheets (excel_file_id, sheet_name, data_json)
                                VALUES (:file_id, :sheet_name, :data_json)
                            """),
                            {
                                "file_id": file_id,
                                "sheet_name": sheet_name,
                                "data_json": data_json
                            }
                        )
                
                session.commit()
                print(f"Alle Sheets erfolgreich in Datenbank gespeichert")
                
        except Exception as e:
            print(f"Fehler beim Speichern der Sheets: {str(e)}")
    
    def load_excel_file(self, file_hash=None, filename=None):
        """Lädt eine Excel-Datei aus der Datenbank"""
        if not self.engine:
            return None, []
        
        try:
            with self.Session() as session:
                if file_hash:
                    result = session.execute(
                        text("SELECT id, filename, metadata FROM excel_files WHERE file_hash = :hash"),
                        {"hash": file_hash}
                    ).fetchone()
                elif filename:
                    result = session.execute(
                        text("SELECT id, filename, metadata FROM excel_files WHERE filename = :filename"),
                        {"filename": filename}
                    ).fetchone()
                else:
                    return None, []
                
                if not result:
                    return None, []
                
                file_id, filename, metadata = result
                
                # Sheets laden
                sheets_result = session.execute(
                    text("SELECT sheet_name, data_json FROM sheets WHERE excel_file_id = :file_id"),
                    {"file_id": file_id}
                ).fetchall()
                
                sheets_data = {}
                sheet_names = []
                
                for sheet_name, data_json in sheets_result:
                    df = pd.read_json(data_json, orient='records')
                    sheets_data[sheet_name] = df
                    sheet_names.append(sheet_name)
                
                return sheets_data, sheet_names
                
        except Exception as e:
            print(f"Fehler beim Laden der Datei: {str(e)}")
            return None, []
    
    def list_saved_files(self):
        """Listet alle gespeicherten Excel-Dateien"""
        if not self.engine:
            return []
        
        try:
            with self.Session() as session:
                result = session.execute(
                    text("SELECT id, filename, upload_date, metadata FROM excel_files ORDER BY upload_date DESC")
                ).fetchall()
                
                files = []
                for file_id, filename, upload_date, metadata in result:
                    files.append({
                        'id': file_id,
                        'filename': filename,
                        'upload_date': upload_date,
                        'metadata': json.loads(metadata) if metadata else {}
                    })
                
                return files
                
        except Exception as e:
            print(f"Fehler beim Laden der Dateiliste: {str(e)}")
            return []
    
    def create_backup(self, file_id, sheets_data, backup_type='manual'):
        """Erstellt ein Backup der Excel-Datei"""
        if not self.engine:
            return None
        
        try:
            with self.Session() as session:
                backup_data = {}
                for sheet_name, df in sheets_data.items():
                    backup_data[sheet_name] = df.to_json(orient='records', date_format='iso')
                
                backup_json = json.dumps(backup_data)
                
                result = session.execute(
                    text("""
                        INSERT INTO backups (excel_file_id, backup_data_json, backup_type)
                        VALUES (:file_id, :backup_data, :backup_type)
                        RETURNING id
                    """),
                    {
                        "file_id": file_id,
                        "backup_data": backup_json,
                        "backup_type": backup_type
                    }
                )
                
                backup_id = result.fetchone()[0]
                session.commit()
                
                print(f"Backup erfolgreich erstellt (ID: {backup_id})")
                return backup_id
                
        except Exception as e:
            print(f"Fehler beim Erstellen des Backups: {str(e)}")
            return None
    
    def load_backup(self, backup_id):
        """Lädt ein Backup aus der Datenbank"""
        if not self.engine:
            return None
        
        try:
            with self.Session() as session:
                result = session.execute(
                    text("SELECT backup_data_json FROM backups WHERE id = :backup_id"),
                    {"backup_id": backup_id}
                ).fetchone()
                
                if not result:
                    return None
                
                backup_data = json.loads(result[0])
                sheets_data = {}
                
                for sheet_name, data_json in backup_data.items():
                    df = pd.read_json(data_json, orient='records')
                    sheets_data[sheet_name] = df
                
                return sheets_data
                
        except Exception as e:
            print(f"Fehler beim Laden des Backups: {str(e)}")
            return None
    
    def list_backups(self, file_id):
        """Listet alle Backups einer Datei"""
        if not self.engine:
            return []
        
        try:
            with self.Session() as session:
                result = session.execute(
                    text("SELECT id, backup_date, backup_type FROM backups WHERE excel_file_id = :file_id ORDER BY backup_date DESC"),
                    {"file_id": file_id}
                ).fetchall()
                
                backups = []
                for backup_id, backup_date, backup_type in result:
                    backups.append({
                        'id': backup_id,
                        'backup_date': backup_date,
                        'backup_type': backup_type
                    })
                
                return backups
                
        except Exception as e:
            print(f"Fehler beim Laden der Backups: {str(e)}")
            return []
    
    def delete_file(self, file_id):
        """Löscht eine Excel-Datei und alle zugehörigen Daten"""
        if not self.engine:
            return False
        
        try:
            with self.Session() as session:
                # Lösche zuerst alle Backups
                session.execute(
                    text("DELETE FROM backups WHERE excel_file_id = :file_id"),
                    {"file_id": file_id}
                )
                
                # Lösche alle Sheets
                session.execute(
                    text("DELETE FROM sheets WHERE excel_file_id = :file_id"),
                    {"file_id": file_id}
                )
                
                # Lösche die Datei
                session.execute(
                    text("DELETE FROM excel_files WHERE id = :file_id"),
                    {"file_id": file_id}
                )
                
                session.commit()
                print(f"Datei mit ID {file_id} erfolgreich gelöscht")
                return True
                
        except Exception as e:
            print(f"Fehler beim Löschen der Datei: {str(e)}")
            return False 