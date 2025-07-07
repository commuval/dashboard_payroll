import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import subprocess
from datetime import datetime, timedelta
import tempfile
import shutil
import json
import pickle
import glob
import io
from pathlib import Path

# Prüfe ob die App über streamlit läuft
def is_streamlit():
    """Prüft ob das Script über 'streamlit run' ausgeführt wird"""
    try:
        import streamlit.runtime.scriptrunner
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except:
        return False

# Wenn nicht über streamlit gestartet, starte mit streamlit
if __name__ == "__main__" and not is_streamlit():
    # Hole Port aus Umgebungsvariable
    port = os.environ.get('PORT', '8080')
    
    # Starte mit streamlit
    cmd = [
        sys.executable, '-m', 'streamlit', 'run', __file__,
        '--server.port', port,
        '--server.address', '0.0.0.0',
        '--server.headless', 'true',
        '--server.enableCORS', 'false',
        '--server.enableXsrfProtection', 'false'
    ]
    
    try:
        subprocess.run(cmd, check=True)
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"Fehler beim Starten von Streamlit: {e}")
        sys.exit(1)
    sys.exit(0)

# Streamlit Konfiguration
st.set_page_config(
    page_title="Excel Viewer Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExcelViewerWeb:
    def __init__(self):
        self.setup_directories()
    
    def setup_session_state(self):
        """Initialisiert Session State Variablen"""
        if 'data' not in st.session_state:
            st.session_state.data = None
        if 'sheet_names' not in st.session_state:
            st.session_state.sheet_names = []
        if 'current_sheet' not in st.session_state:
            st.session_state.current_sheet = None
        if 'sorted_data' not in st.session_state:
            st.session_state.sorted_data = None
        if 'file_name' not in st.session_state:
            st.session_state.file_name = None
        if 'backup_enabled' not in st.session_state:
            st.session_state.backup_enabled = True
    
    def setup_directories(self):
        """Erstellt notwendige Verzeichnisse"""
        self.backup_dir = Path("backups")
        self.backup_dir.mkdir(exist_ok=True)
    
    def load_excel_data(self, uploaded_file):
        """Lädt Excel-Datei und gibt DataFrame zurück"""
        try:
            # Excel-Datei laden
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            # Alle Sheets laden
            sheets_data = {}
            for sheet in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                sheets_data[sheet] = df
            
            return sheets_data, sheet_names
        except Exception as e:
            st.error(f"Fehler beim Laden der Excel-Datei: {str(e)}")
            return None, []
    
    def clean_value(self, value):
        """Bereinigt Werte für bessere Anzeige"""
        if pd.isna(value):
            return ""
        elif isinstance(value, (int, float)):
            if pd.isna(value):
                return ""
            return str(value)
        else:
            return str(value).strip()
    
    def sortiere_nach_praxis(self, df):
        """Sortiert DataFrame nach Praxis-Spalte"""
        try:
            # Suche nach Praxis-Spalte (verschiedene mögliche Namen)
            praxis_cols = [col for col in df.columns if 'praxis' in str(col).lower()]
            
            if not praxis_cols:
                st.warning("Keine 'Praxis'-Spalte gefunden. Verfügbare Spalten: " + ", ".join(df.columns))
                return df
            
            praxis_col = praxis_cols[0]
            
            # Sortieren nach Praxis
            sorted_df = df.sort_values(by=praxis_col, na_position='last')
            
            st.success(f"Daten erfolgreich nach '{praxis_col}' sortiert!")
            return sorted_df
            
        except Exception as e:
            st.error(f"Fehler beim Sortieren: {str(e)}")
            return df
    
    def create_backup(self, df, filename):
        """Erstellt Backup der Daten"""
        if not st.session_state.backup_enabled:
            return
        
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            
            # Pickle Backup
            pickle_filename = f"{filename}_backup_{timestamp}.pkl"
            pickle_path = self.backup_dir / pickle_filename
            
            with open(pickle_path, 'wb') as f:
                pickle.dump(df, f)
            
            # Excel Backup
            excel_filename = f"{filename}_backup_{timestamp}.xlsx"
            excel_path = self.backup_dir / excel_filename
            
            df.to_excel(excel_path, index=False)
            
            st.success(f"Backup erstellt: {pickle_filename} und {excel_filename}")
            
        except Exception as e:
            st.error(f"Fehler beim Erstellen des Backups: {str(e)}")
    
    def load_sorted_data(self):
        """Lädt gespeicherte sortierte Daten"""
        try:
            sorted_files = list(self.backup_dir.glob("*_sorted_data.pkl"))
            if sorted_files:
                # Neueste Datei laden
                latest_file = max(sorted_files, key=os.path.getctime)
                with open(latest_file, 'rb') as f:
                    data = pickle.load(f)
                return data
        except Exception as e:
            st.error(f"Fehler beim Laden der sortierten Daten: {str(e)}")
        return None
    
    def save_sorted_data(self, df, filename):
        """Speichert sortierte Daten"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Pickle speichern
            pickle_filename = f"{filename}_sorted_data.pkl"
            pickle_path = self.backup_dir / pickle_filename
            
            with open(pickle_path, 'wb') as f:
                pickle.dump(df, f)
            
            st.success(f"Sortierte Daten gespeichert: {pickle_filename}")
            
            # Auch als Excel zum Download anbieten
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sorted_Data')
            
            return output.getvalue()
            
        except Exception as e:
            st.error(f"Fehler beim Speichern: {str(e)}")
            return None
    
    def display_data_editor(self, df):
        """Zeigt bearbeitbaren Dateneditor an"""
        if df is None or df.empty:
            st.warning("Keine Daten zum Anzeigen vorhanden.")
            return df
        
        # Konfiguration für den Data Editor
        column_config = {}
        for col in df.columns:
            column_config[col] = st.column_config.TextColumn(
                col,
                help=f"Bearbeiten Sie die Werte in der Spalte '{col}'",
                max_chars=100,
            )
        
        # Data Editor anzeigen
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",  # Ermöglicht Hinzufügen/Entfernen von Zeilen
            column_config=column_config,
            hide_index=True,
            key="data_editor"
        )
        
        return edited_df
    
    def main(self):
        """Hauptfunktion der Streamlit App"""
        
        # Session State initialisieren (ERST hier, innerhalb des Streamlit-Kontexts)
        self.setup_session_state()
        
        # Header
        st.title("📊 Excel Viewer Pro")
        st.markdown("---")
        
        # Sidebar für Datei-Upload und Optionen
        with st.sidebar:
            st.header("🔧 Optionen")
            
            # Backup-Einstellungen
            st.session_state.backup_enabled = st.checkbox(
                "Automatische Backups", 
                value=st.session_state.backup_enabled
            )
            
            # Datei-Upload
            st.header("📁 Datei hochladen")
            uploaded_file = st.file_uploader(
                "Excel-Datei auswählen",
                type=['xlsx', 'xls'],
                help="Laden Sie eine Excel-Datei hoch, um sie zu bearbeiten"
            )
            
            # Gespeicherte Daten laden
            if st.button("🔄 Letzte sortierte Daten laden"):
                loaded_data = self.load_sorted_data()
                if loaded_data is not None:
                    st.session_state.data = {"Loaded_Data": loaded_data}
                    st.session_state.current_sheet = "Loaded_Data"
                    st.session_state.sheet_names = ["Loaded_Data"]
                    st.success("Sortierte Daten erfolgreich geladen!")
                    st.rerun()
        
        # Hauptbereich
        if uploaded_file is not None:
            # Datei verarbeiten
            if st.session_state.file_name != uploaded_file.name:
                st.session_state.file_name = uploaded_file.name
                with st.spinner("Lade Excel-Datei..."):
                    sheets_data, sheet_names = self.load_excel_data(uploaded_file)
                    if sheets_data:
                        st.session_state.data = sheets_data
                        st.session_state.sheet_names = sheet_names
                        st.session_state.current_sheet = sheet_names[0] if sheet_names else None
                        st.success(f"Datei '{uploaded_file.name}' erfolgreich geladen!")
        
        # Daten anzeigen und bearbeiten
        if st.session_state.data:
            # Sheet-Auswahl
            if len(st.session_state.sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "Sheet auswählen:",
                    st.session_state.sheet_names,
                    index=st.session_state.sheet_names.index(st.session_state.current_sheet) 
                          if st.session_state.current_sheet in st.session_state.sheet_names else 0
                )
                st.session_state.current_sheet = selected_sheet
            
            # Aktuelle Daten holen
            current_df = st.session_state.data[st.session_state.current_sheet]
            
            # Aktionen
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                if st.button("🔄 Nach Praxis sortieren", use_container_width=True):
                    sorted_df = self.sortiere_nach_praxis(current_df.copy())
                    st.session_state.data[st.session_state.current_sheet] = sorted_df
                    st.session_state.sorted_data = sorted_df
                    st.rerun()
            
            with col2:
                if st.button("💾 Backup erstellen", use_container_width=True):
                    self.create_backup(current_df, st.session_state.file_name or "data")
            
            with col3:
                if st.session_state.sorted_data is not None:
                    excel_data = self.save_sorted_data(
                        st.session_state.sorted_data, 
                        st.session_state.file_name or "sorted_data"
                    )
                    if excel_data:
                        st.download_button(
                            label="⬇️ Download Excel",
                            data=excel_data,
                            file_name=f"sorted_{st.session_state.file_name or 'data'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            
            with col4:
                if st.button("🗑️ Daten zurücksetzen", use_container_width=True):
                    for key in ['data', 'sorted_data', 'current_sheet', 'sheet_names', 'file_name']:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.rerun()
            
            st.markdown("---")
            
            # Daten-Info
            st.subheader(f"📊 Daten: {st.session_state.current_sheet}")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Zeilen", len(current_df))
            with col2:
                st.metric("Spalten", len(current_df.columns))
            with col3:
                st.metric("Speicher", f"{current_df.memory_usage(deep=True).sum() / 1024:.1f} KB")
            
            # Dateneditor
            st.subheader("✏️ Daten bearbeiten")
            edited_df = self.display_data_editor(current_df)
            
            # Änderungen speichern
            if not edited_df.equals(current_df):
                st.session_state.data[st.session_state.current_sheet] = edited_df
                st.info("Änderungen wurden automatisch gespeichert!")
        
        else:
            # Willkommensnachricht
            st.info("👆 Bitte laden Sie eine Excel-Datei über die Seitenleiste hoch, um zu beginnen.")
            
            # Verfügbare Backups anzeigen
            backup_files = list(self.backup_dir.glob("*.xlsx"))
            if backup_files:
                st.subheader("📦 Verfügbare Backups")
                for backup_file in sorted(backup_files, key=os.path.getctime, reverse=True)[:5]:
                    st.text(f"📄 {backup_file.name}")

# App starten - nur wenn über Streamlit ausgeführt
if __name__ == "__main__" and is_streamlit():
    app = ExcelViewerWeb()
    app.main() 