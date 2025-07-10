# Dashboard & Payroll System

Ein Flask-basiertes Dashboard für die Verwaltung und Analyse von Excel-Dateien mit PostgreSQL-Datenbank (Supabase).

## Features

- Excel-Datei-Upload und -Verarbeitung
- Automatische Sortierung nach Praxis
- Datenbank-Speicherung mit Supabase
- Backup-Funktionalität
- Web-basierte Benutzeroberfläche

## Supabase Setup

### 1. Supabase-Projekt erstellen

1. Gehen Sie zu [supabase.com](https://supabase.com)
2. Erstellen Sie ein neues Projekt
3. Notieren Sie sich die Projekt-Referenz und das Datenbankpasswort

### 2. Datenbank-Schema

Das System verwendet folgende Tabellen:

```sql
CREATE TABLE public.excel_files (
  id integer NOT NULL DEFAULT nextval('excel_files_id_seq'::regclass),
  filename character varying NOT NULL,
  upload_date timestamp without time zone DEFAULT CURRENT_TIMESTAMP,
  file_hash character varying UNIQUE,
  metadata text,
  CONSTRAINT excel_files_pkey PRIMARY KEY (id)
);

CREATE TABLE public.sheets (
  id integer NOT NULL DEFAULT nextval('sheets_id_seq'::regclass),
  excel_file_id integer,
  sheet_name character varying NOT NULL,
  data_json text,
  last_modified timestamp without time zone DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT sheets_pkey PRIMARY KEY (id),
  CONSTRAINT sheets_excel_file_id_fkey FOREIGN KEY (excel_file_id) REFERENCES public.excel_files(id)
);

CREATE TABLE public.backups (
  id integer NOT NULL DEFAULT nextval('backups_id_seq'::regclass),
  excel_file_id integer,
  backup_date timestamp without time zone DEFAULT CURRENT_TIMESTAMP,
  backup_data_json text,
  backup_type character varying,
  CONSTRAINT backups_pkey PRIMARY KEY (id),
  CONSTRAINT backups_excel_file_id_fkey FOREIGN KEY (excel_file_id) REFERENCES public.excel_files(id)
);
```

### 3. Umgebungsvariablen konfigurieren

Erstellen Sie eine `.env` Datei im Projektverzeichnis:

```env
# Flask Configuration
SECRET_KEY=your-secret-key-here
FLASK_ENV=production

# Supabase Database Configuration
DATABASE_URL=postgresql://postgres:[YOUR-PASSWORD]@db.[YOUR-PROJECT-REF].supabase.co:5432/postgres?sslmode=require

# Upload Configuration
MAX_CONTENT_LENGTH=16777216
```

Ersetzen Sie:
- `[YOUR-PASSWORD]` mit Ihrem Supabase-Datenbankpasswort
- `[YOUR-PROJECT-REF]` mit Ihrer Supabase-Projekt-Referenz

## Installation

1. **Abhängigkeiten installieren:**
```bash
pip install -r requirements.txt
```

2. **Datenbank-Tabellen erstellen:**
```bash
# Starten Sie die Anwendung und besuchen Sie:
# http://localhost:5000/setup-database
```

3. **Anwendung starten:**
```bash
python app.py
```

## Verwendung

1. **Datei hochladen:** Besuchen Sie die Hauptseite und laden Sie eine Excel-Datei hoch
2. **Daten anzeigen:** Die Daten werden automatisch in der Datenbank gespeichert und angezeigt
3. **Nach Praxis sortieren:** Verwenden Sie die Sortier-Funktion, um Daten nach Praxis zu gruppieren
4. **Backups erstellen:** Erstellen Sie Backups Ihrer Daten über die Backup-Funktion

## Datenbank-Test

Testen Sie die Datenbankverbindung unter:
```
http://localhost:5000/test-db
```

## Deployment

Das System ist für Deployment auf Heroku vorbereitet:

1. **Heroku-App erstellen**
2. **Umgebungsvariablen setzen:**
```bash
heroku config:set DATABASE_URL=postgresql://postgres:[YOUR-PASSWORD]@db.[YOUR-PROJECT-REF].supabase.co:5432/postgres?sslmode=require
heroku config:set SECRET_KEY=your-secret-key-here
```

3. **Deployen:**
```bash
git push heroku main
```

## Sicherheit

- Alle Datenbankverbindungen verwenden SSL
- Datei-Hashes verhindern Duplikate
- Sichere Datei-Upload-Validierung
- Session-basierte Authentifizierung 