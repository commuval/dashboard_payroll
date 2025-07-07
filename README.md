# Excel Viewer Pro - Web Version

Eine webbasierte Excel-Viewer-Anwendung, die speziell für Cloud-Deployment (DigitalOcean) entwickelt wurde.

## Features

- 📊 **Excel-Dateien hochladen und anzeigen**: Unterstützt .xlsx und .xls Dateien
- 🔄 **Multi-Sheet Support**: Wechseln zwischen verschiedenen Excel-Sheets
- ⚡ **Sortierung nach Praxis**: Automatische Sortierung nach Praxis-Spalten
- ✏️ **Inline-Bearbeitung**: Direkte Bearbeitung von Zellen im Browser
- 💾 **Backup-System**: Automatische Backups mit Zeitstempel
- ⬇️ **Download-Funktion**: Export der bearbeiteten Daten als Excel
- 🎨 **Moderne Web-UI**: Responsive Design mit Streamlit

## Installation & Lokale Entwicklung

```bash
# Dependencies installieren
pip install -r requirements.txt

# App lokal starten
streamlit run app.py
```

Die App ist dann unter `http://localhost:8501` verfügbar.

## DigitalOcean Deployment

### 1. App Platform Setup

1. **Repository verbinden**: Verknüpfen Sie Ihr GitHub-Repository mit DigitalOcean App Platform
2. **Build-Einstellungen**:
   - **Run Command**: `./start.sh` oder `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true`
   - **Environment**: Python
   - **Source Directory**: `/` (Root)

### 2. Umgebungsvariablen

Keine speziellen Umgebungsvariablen erforderlich.

### 3. Port-Konfiguration

Die App ist für Port 8080 konfiguriert, der automatisch von DigitalOcean zugewiesen wird.

## Unterschiede zur Desktop-Version

- ❌ **Entfernt**: tkinter (Desktop-GUI)
- ✅ **Neu**: Streamlit Web-Interface
- ✅ **Verbessert**: Cloud-native Architektur
- ✅ **Hinzugefügt**: Web-Download-Funktionalität

## Datei-Struktur

```
Dashboard & Payroll/
├── app.py                 # Haupt-Streamlit-Anwendung
├── requirements.txt       # Python-Dependencies
├── runtime.txt           # Python-Version für DigitalOcean
├── start.sh              # Startup-Script für DigitalOcean
├── .streamlit/
│   └── config.toml       # Streamlit-Konfiguration
├── backups/              # Automatische Backups
└── README.md             # Diese Datei
```

## Verwendung

1. **Excel-Datei hochladen**: Nutzen Sie die Seitenleiste zum Upload
2. **Sheet auswählen**: Bei Multi-Sheet-Dateien können Sie das gewünschte Sheet wählen
3. **Daten sortieren**: Klicken Sie auf "Nach Praxis sortieren"
4. **Bearbeiten**: Nutzen Sie den integrierten Dateneditor
5. **Speichern**: Erstellen Sie Backups oder laden Sie die Datei herunter

## Technische Details

- **Framework**: Streamlit 1.28+
- **Backend**: Python 3.13
- **Excel-Handling**: pandas + openpyxl
- **Cloud-Platform**: DigitalOcean App Platform
- **Backup-Format**: Pickle (.pkl) + Excel (.xlsx)

## Troubleshooting

### Deployment-Probleme

- Stellen Sie sicher, dass `requirements.txt` alle Dependencies enthält
- Port 8080 muss von der App verwendet werden
- `start.sh` muss im Root-Verzeichnis liegen

### Performance

- Große Excel-Dateien (>50MB) können länger laden
- Backups werden automatisch nach 10 Tagen gelöscht

## Support

Bei Problemen prüfen Sie die DigitalOcean App-Logs für detaillierte Fehlermeldungen. 