# Excel Viewer Pro - Payroll Dashboard

Eine professionelle Webanwendung zur Analyse und Sortierung von Excel-Dateien, speziell entwickelt für Payroll-Daten und Praxis-Management.

## Features

- **Excel-Datei Upload**: Unterstützt .xlsx und .xls Dateien bis 16MB
- **Automatische Praxis-Sortierung**: Verteilt Daten basierend auf Spalte B (Praxis-Namen)
- **Datenbank-Speicherung**: Sichere PostgreSQL-Integration
- **Backup-System**: Automatische und manuelle Backups
- **Responsive Design**: Moderne, professionelle Benutzeroberfläche
- **Multi-Sheet Support**: Arbeitet mit mehreren Arbeitsblättern

## Technologie-Stack

- **Backend**: Flask (Python)
- **Frontend**: Bootstrap 5, jQuery
- **Datenbank**: PostgreSQL
- **Deployment**: Gunicorn, Digital Ocean

## Installation

### Lokale Entwicklung

1. **Repository klonen**:
   ```bash
   git clone <your-repo-url>
   cd Dashboard-Payroll
   ```

2. **Python-Abhängigkeiten installieren**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Umgebungsvariablen konfigurieren**:
   ```bash
   cp env_example.txt .env
   # Bearbeiten Sie .env mit Ihren Datenbank-Einstellungen
   ```

4. **PostgreSQL-Datenbank einrichten**:
   - Erstellen Sie eine PostgreSQL-Datenbank
   - Konfigurieren Sie die Verbindungsdaten in `.env`

5. **Anwendung starten**:
   ```bash
   python app.py
   ```

### Digital Ocean Deployment

1. **App Platform konfigurieren**:
   - Erstellen Sie eine neue App in Digital Ocean
   - Verbinden Sie Ihr Git-Repository
   - Wählen Sie Python als Runtime

2. **Umgebungsvariablen setzen**:
   - `SECRET_KEY`: Ein sicherer Schlüssel für Flask
   - `DB_HOST`: Ihre PostgreSQL-Datenbank-Host
   - `DB_PORT`: Datenbank-Port (standardmäßig 5432)
   - `DB_NAME`: Datenbank-Name
   - `DB_USER`: Datenbank-Benutzer
   - `DB_PASSWORD`: Datenbank-Passwort

3. **Datenbank einrichten**:
   - Erstellen Sie eine PostgreSQL-Datenbank in Digital Ocean
   - Verwenden Sie die bereitgestellten Verbindungsdaten

4. **Deployment**:
   - Die App wird automatisch deployed, wenn Sie zu Git pushen
   - Gunicorn startet die Anwendung im Production-Modus

## Verwendung

1. **Datei hochladen**: Navigieren Sie zur Startseite und laden Sie eine Excel-Datei hoch
2. **Daten anzeigen**: Die Datei wird automatisch geladen und angezeigt
3. **Nach Praxen sortieren**: Klicken Sie auf "Nach Praxen sortieren" um Daten zu verteilen
4. **Backup erstellen**: Erstellen Sie Backups Ihrer Daten
5. **Dateien verwalten**: Über die "Dateien"-Seite können Sie alle hochgeladenen Dateien verwalten

## Datenbank-Schema

Die Anwendung erstellt automatisch folgende Tabellen:

- `excel_files`: Speichert Metadaten zu hochgeladenen Dateien
- `sheets`: Speichert die eigentlichen Excel-Daten als JSON
- `backups`: Speichert Backup-Versionen der Daten

## Sicherheit

- Sichere Datei-Upload-Validierung
- SQL-Injection-Schutz durch SQLAlchemy
- CSRF-Schutz durch Flask
- Sichere Session-Verwaltung

## Support

Bei Fragen oder Problemen:
1. Überprüfen Sie die Logs in Digital Ocean
2. Stellen Sie sicher, dass die Datenbankverbindung korrekt ist
3. Überprüfen Sie die Umgebungsvariablen

## Lizenz

Dieses Projekt ist für den internen Gebrauch bestimmt. 