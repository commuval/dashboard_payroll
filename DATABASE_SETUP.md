# PostgreSQL Datenbank-Einrichtung

## 1. PostgreSQL Installation

### Windows:
1. Laden Sie PostgreSQL von https://www.postgresql.org/download/windows/ herunter
2. Installieren Sie PostgreSQL mit dem Standard-Setup
3. Merken Sie sich Benutzername und Passwort für den `postgres` Benutzer

### macOS:
```bash
brew install postgresql
brew services start postgresql
```

### Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install postgresql postgresql-contrib
sudo systemctl start postgresql
sudo systemctl enable postgresql
```

## 2. Datenbank erstellen

```sql
-- Als postgres Benutzer anmelden
sudo -u postgres psql

-- Neue Datenbank erstellen
CREATE DATABASE excel_viewer;

-- Neuen Benutzer erstellen (optional)
CREATE USER excel_user WITH PASSWORD 'ihr_passwort';

-- Berechtigungen vergeben
GRANT ALL PRIVILEGES ON DATABASE excel_viewer TO excel_user;

-- Beenden
\q
```

## 3. Umgebungsvariablen konfigurieren

Erstellen Sie eine `.env` Datei im Projektverzeichnis:

```env
# PostgreSQL Database Configuration
DATABASE_URL=postgresql://excel_user:ihr_passwort@localhost:5432/excel_viewer

# Oder einzeln:
DB_HOST=localhost
DB_PORT=5432
DB_NAME=excel_viewer
DB_USER=excel_user
DB_PASSWORD=ihr_passwort
```

## 4. Verbindung in der App

1. Starten Sie die Excel Viewer App
2. Aktivieren Sie in der Sidebar "PostgreSQL verwenden"
3. Geben Sie die Datenbank-URL ein oder lassen Sie das Feld leer für Umgebungsvariablen
4. Klicken Sie auf "Verbinden"

## 5. Funktionen

### Automatische Tabellen-Erstellung
Die App erstellt automatisch folgende Tabellen:
- `excel_files`: Speichert hochgeladene Excel-Daten
- `backups`: Speichert Backups der Daten

### Daten speichern
- Laden Sie eine Excel-Datei hoch
- Klicken Sie auf "In DB speichern" um die Daten in PostgreSQL zu speichern

### Daten laden
- Gespeicherte Dateien werden in der Sidebar unter "Gespeicherte Dateien" angezeigt
- Klicken Sie auf "Laden" um Daten aus der Datenbank zu laden

### Backups
- Manuelle Backups werden sowohl als Datei als auch in der Datenbank gespeichert
- Automatische Backups bei Änderungen (wenn aktiviert)

## 6. Produktions-Deployment

### Heroku PostgreSQL:
```env
DATABASE_URL=postgres://user:password@hostname:port/dbname
```

### AWS RDS:
```env
DATABASE_URL=postgresql://username:password@endpoint:5432/dbname
```

### Google Cloud SQL:
```env
DATABASE_URL=postgresql://username:password@ip-address:5432/dbname
```

## Fehlerbehebung

### Verbindungsfehler:
1. Prüfen Sie PostgreSQL-Service: `sudo systemctl status postgresql`
2. Prüfen Sie Firewall-Einstellungen
3. Prüfen Sie Datenbank-URL Format

### Berechtigungsfehler:
```sql
-- Als postgres Benutzer
GRANT ALL PRIVILEGES ON ALL TABLES IN SCHEMA public TO excel_user;
GRANT ALL PRIVILEGES ON ALL SEQUENCES IN SCHEMA public TO excel_user;
```

### Port bereits in Verwendung:
```sql
-- Anderen Port verwenden
ALTER SYSTEM SET port = 5433;
SELECT pg_reload_conf();
``` 