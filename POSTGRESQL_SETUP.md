# PostgreSQL-Einrichtung für Excel Viewer Pro

## 1. PostgreSQL installieren

### Windows
1. Laden Sie PostgreSQL von https://www.postgresql.org/download/windows/ herunter
2. Führen Sie den Installer aus
3. Notieren Sie sich das Passwort für den `postgres` Benutzer
4. Standard-Port: 5432

### macOS
```bash
# Mit Homebrew
brew install postgresql
brew services start postgresql

# Oder mit Postgres.app
# Laden Sie Postgres.app von https://postgresapp.com/ herunter
```

### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install postgresql postgresql-contrib
sudo systemctl start postgresql
sudo systemctl enable postgresql
```

## 2. Datenbank erstellen

### Mit pgAdmin (GUI)
1. Öffnen Sie pgAdmin
2. Verbinden Sie sich mit dem Server
3. Rechtsklick auf "Databases" → "Create" → "Database"
4. Name: `excel_viewer`

### Mit Kommandozeile
```bash
# Als postgres Benutzer anmelden
sudo -u postgres psql

# Datenbank erstellen
CREATE DATABASE excel_viewer;

# Benutzer erstellen (optional)
CREATE USER excel_user WITH PASSWORD 'your_password';
GRANT ALL PRIVILEGES ON DATABASE excel_viewer TO excel_user;

# Beenden
\q
```

## 3. Umgebungsvariablen konfigurieren

### Lokale Entwicklung
Erstellen Sie eine `.env` Datei im Projektverzeichnis:

```env
DB_HOST=localhost
DB_PORT=5432
DB_NAME=excel_viewer
DB_USER=postgres
DB_PASSWORD=your_postgres_password
```

### Für Streamlit Cloud/Deployment
Fügen Sie die Secrets in der Streamlit Cloud hinzu:

```toml
[postgres]
host = "your-db-host"
port = 5432
database = "excel_viewer"
user = "your-username"
password = "your-password"
```

## 4. Dependencies installieren

```bash
pip install -r requirements.txt
```

## 5. App starten

```bash
streamlit run app.py
```

## 6. Datenbank-Tabellen

Die App erstellt automatisch folgende Tabellen:

### excel_files
- `id`: Primärschlüssel
- `filename`: Name der Excel-Datei
- `upload_date`: Upload-Datum
- `file_hash`: MD5-Hash der Datei (für Duplikatserkennung)
- `metadata`: JSON mit zusätzlichen Metadaten

### sheets
- `id`: Primärschlüssel
- `excel_file_id`: Fremdschlüssel zu excel_files
- `sheet_name`: Name des Excel-Sheets
- `data_json`: JSON-Daten des Sheets
- `last_modified`: Letzte Änderung

### backups
- `id`: Primärschlüssel
- `excel_file_id`: Fremdschlüssel zu excel_files
- `backup_date`: Backup-Datum
- `backup_data_json`: JSON-Backup-Daten
- `backup_type`: Typ des Backups ('manual', 'auto')

## 7. Features

### Automatische Speicherung
- Excel-Dateien werden automatisch in der Datenbank gespeichert
- Duplikatserkennung basierend auf Datei-Hash
- Metadaten-Speicherung (Upload-Datum, Dateigröße, Anzahl Sheets)

### Backup-System
- Manuelle und automatische Backups
- Backup-Historie pro Datei
- Wiederherstellung von Backups

### Datenbank-Verwaltung
- Anzeige aller gespeicherten Dateien
- Backup-Verwaltung
- Datei-Löschung mit allen zugehörigen Daten

## 8. Troubleshooting

### Verbindungsfehler
- Prüfen Sie, ob PostgreSQL läuft
- Überprüfen Sie die Verbindungsdaten in `.env`
- Stellen Sie sicher, dass die Datenbank existiert

### Berechtigungsfehler
- Prüfen Sie die Benutzerrechte
- Stellen Sie sicher, dass der Benutzer auf die Datenbank zugreifen kann

### Speicherfehler
- Prüfen Sie den verfügbaren Speicherplatz
- Große Excel-Dateien können viel Speicher benötigen

## 9. Performance-Optimierung

### Für große Dateien
- Erhöhen Sie die PostgreSQL-Speicher-Einstellungen
- Verwenden Sie Connection Pooling
- Optimieren Sie die JSON-Speicherung

### Backup-Strategie
- Regelmäßige automatische Backups
- Archivierung alter Backups
- Monitoring der Datenbankgröße 