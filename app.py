import tkinter as tk
from tkinter import ttk, filedialog, colorchooser
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import tempfile
import shutil
import json
import pickle
import glob

class ExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Viewer Pro")
        self.root.geometry("1400x900")
        
        # Farben und Styling
        self.bg_color = "#f0f0f0"
        self.accent_color = "#007acc"
        self.text_color = "#333333"
        self.header_bg = "#1a5276"  # Dunkleres Blau
        self.header_fg = "#ffffff"  # Weiß
        self.row_color1 = "#ffffff"  # Weiß für gerade Zeilen
        self.row_color2 = "#f8f9fa"  # Sehr helles Grau für ungerade Zeilen
        self.selected_color = "#e3f2fd"  # Helles Blau für Auswahl
        self.border_color = "#dee2e6"  # Deutliche Trennlinien
        self.hover_color = "#f1f3f4"  # Hover-Farbe
        self.column_separator = "#bdc3c7"  # Farbe für Spalten-Trennlinien
        
        self.root.configure(bg=self.bg_color)
        
        # Styling
        self.style = ttk.Style()
        
        # Treeview Styling
        self.style.configure("Treeview",
                           background=self.bg_color,
                           foreground=self.text_color,
                           rowheight=30,  # Höhere Zeilen für bessere Lesbarkeit
                           fieldbackground=self.bg_color,
                           font=('Segoe UI', 10),
                           borderwidth=1,  # Rahmenbreite
                           relief="solid",
                           selectbackground=self.accent_color,
                           selectforeground="#ffffff",
                           highlightthickness=0,  # Entfernt äußere Fokus-Linie
                           lightcolor=self.border_color,  # Helle Trennlinie
                           darkcolor=self.border_color,   # Dunkle Trennlinie
                           bordercolor=self.border_color) # Rahmenfarbe
        
        # Treeview Header Styling
        self.style.configure("Treeview.Heading",
                           background="#ffffff",  # Weißer Hintergrund
                           foreground="#2196F3",  # Blauer Text für die Kategorien
                           font=('Segoe UI', 11, 'bold'),
                           relief="solid",
                           borderwidth=1,  # Rahmenbreite für Header
                           lightcolor=self.border_color,  # Helle Trennlinie
                           darkcolor=self.border_color)   # Dunkle Trennlinie
        
        # Treeview Header Hover
        self.style.map('Treeview.Heading',
                      background=[('active', '#f5f5f5')],  # Helles Grau beim Hover
                      foreground=[('active', '#1976D2')])  # Dunkleres Blau beim Hover
        
        # Treeview Zeilen-Styling
        self.style.map('Treeview',
                      background=[('selected', self.accent_color)],
                      foreground=[('selected', '#ffffff')])
        
        # Zusätzliches Styling für bessere Spalten-Trennung
        self.style.configure("Treeview.Column",
                           borderwidth=1,
                           relief="solid",
                           lightcolor=self.border_color,
                           darkcolor=self.border_color)
        
        # Erweiterte Treeview-Konfiguration für bessere Trennlinien
        self.style.layout("Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])
        
        # Spezielle Konfiguration für deutlichere Zellentrennlinien
        self.style.configure("Treeview.Item",
                           relief="solid",
                           borderwidth=1,
                           highlightthickness=0)
        
        # Andere Styling-Konfigurationen
        self.style.configure("Custom.TButton", 
                           padding=10, 
                           font=('Segoe UI', 10),
                           background=self.accent_color,
                           focuscolor='none')  # Entfernt die blaue Umrandung
        
        # Focus-Ring für alle Button-Elemente entfernen
        self.style.map("Custom.TButton",
                      focuscolor=[('!focus', 'none')])
        
        self.style.configure("Custom.TLabel", 
                           font=('Segoe UI', 10),
                           background=self.bg_color,
                           foreground=self.text_color)
        self.style.configure("Header.TLabel", 
                           font=('Segoe UI', 16, 'bold'),
                           background=self.bg_color,
                           foreground=self.accent_color)
        self.style.configure("Status.TLabel", 
                           font=('Segoe UI', 9),
                           background=self.bg_color,
                           foreground=self.text_color)
        self.style.configure("TLabelframe", 
                           background=self.bg_color)
        self.style.configure("TLabelframe.Label", 
                           font=('Segoe UI', 10, 'bold'),
                           background=self.bg_color,
                           foreground=self.accent_color)
        
        # Hauptcontainer mit Padding
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            self.header_frame,
            text="Excel Viewer Pro",
            style="Header.TLabel"
        ).pack(side=tk.LEFT)
        
        # Upload-Bereich
        self.upload_frame = ttk.LabelFrame(
            self.main_frame,
            text="Datei auswählen",
            padding="10"
        )
        self.upload_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.upload_button = ttk.Button(
            self.upload_frame,
            text="Excel-Datei öffnen",
            command=self.load_excel,
            style="Custom.TButton"
        )
        self.upload_button.pack(side=tk.LEFT, padx=5)
        
        self.sort_button = ttk.Button(
            self.upload_frame,
            text="Sortieren",
            command=self.sortiere_nach_praxis,
            style="Custom.TButton"
        )
        self.sort_button.pack(side=tk.LEFT, padx=5)
        
        self.save_button = ttk.Button(
            self.upload_frame,
            text="Speichern",
            command=self.save_sorted_data,
            style="Custom.TButton"
        )
        self.save_button.pack(side=tk.LEFT, padx=5)
        
        self.file_label = ttk.Label(
            self.upload_frame,
            text="Keine Datei ausgewählt",
            style="Custom.TLabel"
        )
        self.file_label.pack(side=tk.LEFT, padx=20)
        
        # Sheet-Auswahl
        self.sheet_frame = ttk.LabelFrame(
            self.main_frame,
            text="Sheet-Auswahl",
            padding="10"
        )
        self.sheet_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            self.sheet_frame,
            text="Aktuelles Sheet:",
            style="Custom.TLabel"
        ).pack(side=tk.LEFT, padx=5)
        
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(
            self.sheet_frame,
            textvariable=self.sheet_var,
            state="readonly",
            width=40,
            font=('Segoe UI', 10)
        )
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', self.update_display)
        
        # Tabelle mit Rahmen
        self.table_frame = ttk.LabelFrame(
            self.main_frame,
            text="Datenansicht",
            padding="10"
        )
        self.table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview mit Scrollbars
        self.tree = ttk.Treeview(self.table_frame)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Zusätzliche Treeview-Konfiguration für bessere Trennlinien
        self.tree.configure(style="Treeview")
        
        # Standard-Tags für alternierende Zeilen konfigurieren
        self.setup_default_tags()
        
        # Vertikale Scrollbar
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)
        
        # Horizontale Scrollbar
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=hsb.set)
        
        # Bearbeitungs-Entry
        self.edit_entry = ttk.Entry(self.table_frame, font=('Segoe UI', 10))
        self.edit_entry.bind('<Return>', self.save_edit)
        self.edit_entry.bind('<Escape>', self.cancel_edit)
        
        # Status-Bar
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="Bereit",
            style="Status.TLabel"
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Doppelklick-Bindung für Bearbeitung
        self.tree.bind('<Double-1>', self.start_edit)
        
        # Rechtsklick-Menü für Zeilen hinzufügen/löschen
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        # Hover-Effekte hinzufügen
        self.tree.bind('<Motion>', self.on_tree_motion)
        self.tree.bind('<Leave>', self.on_tree_leave)
        
        # Kontext-Menü erstellen
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Neue Zeile hinzufügen", command=self.add_new_row)
        self.context_menu.add_command(label="Zeile löschen", command=self.delete_row)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Farbe auswählen", command=self.choose_row_color)
        
        # Untermenü für voreingestellte Farben
        self.preset_colors_menu = tk.Menu(self.context_menu, tearoff=0)
        self.preset_colors_menu.add_command(label="🟢 Grün (Abgeschlossen)", command=lambda: self.apply_preset_color("#90EE90"))
        self.preset_colors_menu.add_command(label="🟡 Gelb (In Bearbeitung)", command=lambda: self.apply_preset_color("#FFFF99"))
        self.preset_colors_menu.add_command(label="🔴 Rot (Problematisch)", command=lambda: self.apply_preset_color("#FFB3B3"))
        self.preset_colors_menu.add_command(label="🔵 Blau (Wichtig)", command=lambda: self.apply_preset_color("#ADD8E6"))
        self.preset_colors_menu.add_command(label="🟣 Lila (Notiz)", command=lambda: self.apply_preset_color("#DDA0DD"))
        self.preset_colors_menu.add_command(label="🟠 Orange (Warnung)", command=lambda: self.apply_preset_color("#FFD700"))
        
        self.context_menu.add_cascade(label="Schnellfarben", menu=self.preset_colors_menu)
        self.context_menu.add_command(label="Farbe entfernen", command=self.remove_row_color)
        
        self.data = None
        self.current_edit = None
        self.file_path = None
        self.sorted_data = {}  # Speichert die sortierten Daten für jedes Sheet
        self.row_colors = {}  # Speichert die Farben für Zeilen pro Sheet
        self.save_file = None  # Pfad zur Speicherdatei
        self.config_file = "excel_viewer_config.json"  # Konfigurationsdatei
        self.backup_folder = "backups"  # Ordner für Backups
        
        # Backup-Ordner erstellen falls nicht vorhanden
        self.create_backup_folder()
        
        # Versuche die letzte Datei automatisch zu laden
        self.load_last_file()
        
    def start_edit(self, event):
        """Startet die Bearbeitung einer Zelle."""
        # Aktuelle Auswahl
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        column_id = int(column[1]) - 1
        
        # Aktueller Wert
        current_value = self.tree.item(item)['values'][column_id]
        
        # Position des Eintrags
        x, y, width, height = self.tree.bbox(item, column)
        
        # Entry-Widget positionieren
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.delete(0, tk.END)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus()
        
        # Aktuelle Bearbeitung speichern
        self.current_edit = (item, column_id)
        
    def save_edit(self, event):
        """Speichert die Bearbeitung einer Zelle und aktualisiert die sortierten Daten."""
        if self.current_edit:
            item, column_id = self.current_edit
            new_value = self.edit_entry.get()
            
            # Aktuelle Werte holen
            current_values = list(self.tree.item(item)['values'])
            current_values[column_id] = new_value
            
            # Werte aktualisieren
            self.tree.item(item, values=current_values)
            
            # Änderung in den sortierten Daten speichern
            self.update_sorted_data_from_tree()
            
            # Bearbeitung beenden
            self.edit_entry.place_forget()
            self.current_edit = None
            
            self.status_label.config(text="Änderung gespeichert und in sortierten Daten aktualisiert")
        
    def cancel_edit(self, event):
        """Bricht die Bearbeitung ab."""
        self.edit_entry.place_forget()
        self.current_edit = None
        self.status_label.config(text="Bearbeitung abgebrochen")
        
    def load_last_file(self):
        """Lädt automatisch die zuletzt verwendete Excel-Datei."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    last_file = config.get('last_file')
                    if last_file and os.path.exists(last_file):
                        self.load_excel_file(last_file)
                        return
            except Exception as e:
                print(f"Fehler beim Laden der Konfiguration: {str(e)}")
        
        # Wenn keine gespeicherte Datei vorhanden ist
        self.status_label.config(text="Keine vorherige Datei gefunden. Bitte Excel-Datei öffnen.")
    
    def save_config(self):
        """Speichert die Konfiguration mit dem aktuellen Dateipfad."""
        if self.file_path:
            try:
                config = {'last_file': self.file_path}
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"Fehler beim Speichern der Konfiguration: {str(e)}")
    
    def load_excel_file(self, file_path):
        """Lädt eine Excel-Datei (sowohl für manuelles als auch automatisches Laden)."""
        try:
            self.file_path = file_path
            self.data = pd.ExcelFile(file_path)
            
            # Speicherdatei-Pfad erstellen
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.save_file = f"{base_name}_sorted_data.pkl"
            
            # Versuche gespeicherte Daten zu laden
            self.load_sorted_data()
            
            self.sheet_combo['values'] = self.data.sheet_names
            if self.data.sheet_names:
                self.sheet_combo.set(self.data.sheet_names[0])
                self.update_display(None)
            self.file_label.config(text=file_path)
            
            # Konfiguration speichern
            self.save_config()
            
            self.status_label.config(text=f"Datei erfolgreich geladen: {file_path}")
        except Exception as e:
            self.status_label.config(text=f"Fehler beim Laden: {str(e)}")
    
    def load_sorted_data(self):
        """Lädt gespeicherte sortierte Daten und Zeilenfarben."""
        self.load_row_colors()
        if self.sorted_data:
            self.status_label.config(text="Gespeicherte sortierte Daten und Farben geladen.")
        else:
            self.status_label.config(text="Keine gespeicherten Daten gefunden.")
    
    def save_sorted_data(self):
        """Speichert die sortierten Daten und Zeilenfarben und erstellt automatisch ein tägliches Backup."""
        if self.save_file and self.sorted_data:
            try:
                self.save_row_colors()
                self.status_label.config(text="Sortierte Daten und Farben gespeichert.")
                
                # Automatisches tägliches Backup erstellen
                self.create_daily_backup()
                
            except Exception as e:
                self.status_label.config(text=f"Fehler beim Speichern: {str(e)}")
    
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.load_excel_file(file_path)
    
    def clean_value(self, value):
        """Konvertiert NaN-Werte in leere Strings und verbessert die Darstellung."""
        if pd.isna(value) or value == 'nan' or value == 'NaN' or value == 'None':
            return ''
        # Längere Texte kürzen für bessere Darstellung
        str_value = str(value).strip()
        if len(str_value) > 50:
            return str_value[:47] + "..."
        return str_value
    
    def update_display(self, event):
        if not self.data or not self.sheet_var.get():
            return
            
        try:
            aktuelles_sheet = self.sheet_var.get()
            
            # Prüfen, ob sortierte Daten für dieses Sheet vorhanden sind
            if aktuelles_sheet in self.sorted_data:
                df = self.sorted_data[aktuelles_sheet]
                self.status_label.config(text=f"Sortierte Daten für '{aktuelles_sheet}' angezeigt - {len(df)} Zeilen")
            else:
                # Lade das ausgewählte Sheet
                df = pd.read_excel(self.data, sheet_name=aktuelles_sheet)
                self.status_label.config(text=f"Sheet '{aktuelles_sheet}' angezeigt - {len(df)} Zeilen")
            
            # Tabelle leeren
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Spalten konfigurieren
            self.tree['columns'] = list(df.columns)
            self.tree['show'] = 'headings'
            
            # Spaltenüberschriften setzen
            for column in df.columns:
                # "Unnamed" Spalten als leere Überschrift anzeigen
                if str(column).startswith('Unnamed'):
                    header_text = ""
                else:
                    header_text = str(column)
                self.tree.heading(column, text=header_text)
                # Verbesserte Spaltenbreite basierend auf Inhalt
                max_width = max(
                    len(str(column)) * 8,  # Header-Breite
                    df[column].astype(str).str.len().max() * 8 if not df.empty else 100
                )
                # Mindest- und Maximalbreite festlegen
                column_width = max(80, min(max_width, 300))
                self.tree.column(column, width=column_width, minwidth=80, anchor='w')  # Links ausgerichtet für bessere Lesbarkeit
                
                # Styling für bessere Spaltentrennlinien
                self.tree.heading(column, 
                                text=header_text,
                                anchor='w')  # Links ausgerichtet
            
            # Daten einfügen mit alternierenden Zeilenfarben
            for i, row in df.iterrows():
                # Konvertiere alle Werte und ersetze NaN durch leere Strings
                cleaned_values = [self.clean_value(value) for value in row]
                item = self.tree.insert('', 'end', values=cleaned_values)
                
                # Prüfen, ob für diese Zeile eine gespeicherte Farbe existiert
                if (aktuelles_sheet in self.row_colors and 
                    i in self.row_colors[aktuelles_sheet]):
                    # Gespeicherte Farbe anwenden
                    color = self.row_colors[aktuelles_sheet][i]
                    tag_name = f"colored_row_{i}"
                    print(f"Loading saved color {color} for row {i} with tag {tag_name}")  # Debug
                    
                    # Tag konfigurieren
                    self.tree.tag_configure(tag_name, 
                                          background=color,
                                          foreground=self.get_contrasting_text_color(color))
                    self.tree.item(item, tags=(tag_name,))
                    
                    # Debug: Prüfen ob Tag angewendet wurde
                    current_tags = self.tree.item(item, 'tags')
                    print(f"Loaded tags for item {item}: {current_tags}")  # Debug
                else:
                    # Alternierende Zeilenfarben
                    if i % 2 == 0:
                        self.tree.item(item, tags=('evenrow',))
                    else:
                        self.tree.item(item, tags=('oddrow',))
            
            # Tag-Konfiguration ist bereits in setup_default_tags() definiert
            
            # Spalten-Styling für bessere Abgrenzung
            for i, column in enumerate(df.columns):
                self.tree.set(column, column, "")  # Leere Werte für bessere Darstellung
                
        except Exception as e:
            self.status_label.config(text=f"Fehler beim Anzeigen: {str(e)}")

    def sortiere_nach_praxis(self):
        """Sortiert alle Praxis-Sheets automatisch mit den passenden Dashboard-Zeilen. Bereits vorhandene manuelle Änderungen bleiben erhalten."""
        if not self.data:
            self.status_label.config(text="Bitte zuerst eine Excel-Datei laden.")
            return
        try:
            # Dashboard-Daten aus den sortierten Daten holen (falls vorhanden und mit manuellen Änderungen)
            if "Dashboard" in self.sorted_data:
                dashboard_df = self.sorted_data["Dashboard"].copy()
            else:
                dashboard_df = pd.read_excel(self.data, sheet_name="Dashboard")
            
            # Kopie für Vergleich erstellen (Kleinbuchstaben)
            dashboard_df_compare = dashboard_df.copy()
            dashboard_df_compare["Praxis"] = dashboard_df_compare["Praxis"].astype(str).str.strip().str.lower()
            praxis_namen = [name for name in self.data.sheet_names if name.lower().strip() != "dashboard"]
            
            sortierte_sheets = 0
            # Alle Praxis-Sheets durchgehen und sortieren
            for praxis_sheet in praxis_namen:
                praxis_clean = praxis_sheet.strip().lower()
                # Dashboard-Zeilen für diese Praxis (mit Vergleichs-DataFrame)
                dashboard_zeilen = dashboard_df_compare[dashboard_df_compare["Praxis"] == praxis_clean]
                # Originale Dashboard-Zeilen holen (mit ursprünglichen Praxisnamen)
                if not dashboard_zeilen.empty:
                    original_indices = dashboard_zeilen.index
                    dashboard_zeilen = dashboard_df.loc[original_indices]
                else:
                    dashboard_zeilen = pd.DataFrame(columns=dashboard_df.columns)
                
                # Prüfen, ob für dieses Sheet bereits sortierte Daten existieren
                if praxis_sheet in self.sorted_data:
                    # Bereits sortierte Daten vorhanden - diese mit neuen Dashboard-Zeilen aktualisieren
                    existing_df = self.sorted_data[praxis_sheet].copy()
                    
                    # Alte Dashboard-Zeilen entfernen (falls vorhanden) und neue hinzufügen
                    # Annahme: Dashboard-Zeilen stehen am Anfang, Payroll-Zeilen danach
                    if not dashboard_zeilen.empty:
                        # Prüfen, ob es bereits Dashboard-Zeilen gibt (haben meist "Dashboard" im Aufgabenbereich)
                        if "Aufgabenbereich" in existing_df.columns:
                            # Alte Dashboard-Zeilen entfernen
                            non_dashboard_rows = existing_df[existing_df["Aufgabenbereich"] != "Dashboard"]
                            # Neue Dashboard-Zeilen an den Anfang setzen
                            dashboard_zeilen_aligned = dashboard_zeilen.reindex(columns=existing_df.columns)
                            updated_df = pd.concat([dashboard_zeilen_aligned, non_dashboard_rows], ignore_index=True)
                        else:
                            # Keine Aufgabenbereich-Spalte vorhanden, einfach Dashboard-Zeilen oben hinzufügen
                            dashboard_zeilen_aligned = dashboard_zeilen.reindex(columns=existing_df.columns)
                            updated_df = pd.concat([dashboard_zeilen_aligned, existing_df], ignore_index=True)
                        
                        self.sorted_data[praxis_sheet] = updated_df
                        sortierte_sheets += 1
                else:
                    # Neue Sortierung für dieses Sheet
                    # Praxis-Sheet laden
                    try:
                        praxis_df = pd.read_excel(self.data, sheet_name=praxis_sheet)
                    except Exception:
                        praxis_df = pd.DataFrame()
                    
                    # Spalten anpassen und kombinieren
                    if not praxis_df.empty:
                        if not dashboard_zeilen.empty:
                            # Praxis-Sheet an Dashboard-Spalten anpassen (damit alle Informationen erhalten bleiben)
                            praxis_df = praxis_df.reindex(columns=dashboard_zeilen.columns)
                            gefiltert = pd.concat([dashboard_zeilen, praxis_df], ignore_index=True)  # Dashboard-Zeilen zuerst, dann Payroll-Daten
                        else:
                            # Keine Dashboard-Zeilen gefunden, nur Praxis-Daten anzeigen
                            gefiltert = praxis_df
                    else:
                        # Praxis-Sheet ist leer, nur Dashboard-Zeilen anzeigen (falls vorhanden)
                        gefiltert = dashboard_zeilen
                    
                    # Sortierte Daten für dieses Sheet speichern (mit originalem Sheet-Namen)
                    if not gefiltert.empty:
                        self.sorted_data[praxis_sheet] = gefiltert  # Ursprünglicher Name mit Großbuchstaben
                        sortierte_sheets += 1
            
            # Dashboard-Sheet in sortierten Daten speichern
            self.sorted_data["Dashboard"] = dashboard_df
            
            # Aktuelles Sheet neu anzeigen
            aktuelles_sheet = self.sheet_var.get()
            if aktuelles_sheet in self.sorted_data:
                self.update_display(None)
            
            # Sortierte Daten speichern
            self.save_sorted_data()
                
            if sortierte_sheets > 0:
                self.status_label.config(text=f"{sortierte_sheets} Praxis-Sheets aktualisiert. Manuelle Dashboard-Änderungen übernommen.")
            else:
                self.status_label.config(text="Alle Sheets aktualisiert. Manuelle Änderungen bleiben erhalten.")
        except Exception as e:
            self.status_label.config(text=f"Fehler beim Sortieren: {str(e)}")

    def on_tree_motion(self, event):
        """Hover-Effekt für Treeview-Zeilen."""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
    
    def on_tree_leave(self, event):
        """Entfernt Hover-Effekt beim Verlassen der Treeview."""
        self.tree.selection_remove(self.tree.selection())

    def update_sorted_data_from_tree(self):
        """Aktualisiert die sortierten Daten basierend auf dem aktuellen Treeview-Inhalt."""
        aktuelles_sheet = self.sheet_var.get()
        if not aktuelles_sheet:
            return
        
        # Alle Daten aus dem Treeview sammeln
        columns = self.tree['columns']
        data_rows = []
        
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            data_rows.append(values)
        
        if data_rows:
            # DataFrame aus Treeview-Daten erstellen
            df = pd.DataFrame(data_rows, columns=columns)
            # In sortierten Daten speichern
            self.sorted_data[aktuelles_sheet] = df
            # Automatisch speichern
            self.save_sorted_data()

    def show_context_menu(self, event):
        """Zeigt das Kontext-Menü für Rechtsklick."""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def add_new_row(self):
        """Fügt eine neue leere Zeile hinzu."""
        if not self.tree['columns']:
            self.status_label.config(text="Keine Tabelle geladen.")
            return
        
        # Leere Werte für alle Spalten erstellen
        empty_values = [''] * len(self.tree['columns'])
        
        # Neue Zeile hinzufügen
        new_item = self.tree.insert('', 'end', values=empty_values)
        
        # Alternierende Farben anpassen und gespeicherte Farben berücksichtigen
        aktuelles_sheet = self.sheet_var.get()
        items = self.tree.get_children()
        for i, item in enumerate(items):
            # Prüfen, ob für diese Zeile eine gespeicherte Farbe existiert
            if (aktuelles_sheet in self.row_colors and 
                i in self.row_colors[aktuelles_sheet]):
                # Gespeicherte Farbe anwenden
                color = self.row_colors[aktuelles_sheet][i]
                tag_name = f"colored_row_{i}"
                self.tree.tag_configure(tag_name, 
                                      background=color,
                                      foreground=self.get_contrasting_text_color(color))
                self.tree.item(item, tags=(tag_name,))
            else:
                # Alternierende Farben
                if i % 2 == 0:
                    self.tree.item(item, tags=('evenrow',))
                else:
                    self.tree.item(item, tags=('oddrow',))
        
        # Sortierte Daten aktualisieren
        self.update_sorted_data_from_tree()
        
        # Neue Zeile auswählen und zur Bearbeitung freigeben
        self.tree.selection_set(new_item)
        self.tree.focus(new_item)
        
        self.status_label.config(text="Neue Zeile hinzugefügt.")
    
    def delete_row(self):
        """Löscht die ausgewählte Zeile."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_label.config(text="Keine Zeile ausgewählt.")
            return
        
        # Zeilen-Indizes der zu löschenden Zeilen ermitteln
        aktuelles_sheet = self.sheet_var.get()
        deleted_indices = []
        for item in selected_items:
            row_index = self.tree.index(item)
            deleted_indices.append(row_index)
        
        # Zeilen löschen
        for item in selected_items:
            self.tree.delete(item)
        
        # Zeilenfarben-Indizes anpassen (alle Indizes nach gelöschten Zeilen verschieben)
        if aktuelles_sheet in self.row_colors:
            new_colors = {}
            for old_index, color in self.row_colors[aktuelles_sheet].items():
                # Berechnen, um wie viele Positionen der Index verschoben wird
                shift = sum(1 for deleted_index in deleted_indices if deleted_index <= old_index)
                new_index = old_index - shift
                
                # Nur beibehalten, wenn die Zeile nicht gelöscht wurde
                if old_index not in deleted_indices and new_index >= 0:
                    new_colors[new_index] = color
            
            self.row_colors[aktuelles_sheet] = new_colors
        
        # Alternierende Farben neu anwenden mit gespeicherten Farben
        items = self.tree.get_children()
        for i, item in enumerate(items):
            # Prüfen, ob für diese Zeile eine gespeicherte Farbe existiert
            if (aktuelles_sheet in self.row_colors and 
                i in self.row_colors[aktuelles_sheet]):
                # Gespeicherte Farbe anwenden
                color = self.row_colors[aktuelles_sheet][i]
                tag_name = f"colored_row_{i}"
                self.tree.tag_configure(tag_name, 
                                      background=color,
                                      foreground=self.get_contrasting_text_color(color))
                self.tree.item(item, tags=(tag_name,))
            else:
                # Alternierende Farben
                if i % 2 == 0:
                    self.tree.item(item, tags=('evenrow',))
                else:
                    self.tree.item(item, tags=('oddrow',))
        
        # Sortierte Daten aktualisieren
        self.update_sorted_data_from_tree()
        
        self.status_label.config(text="Zeile gelöscht.")

    def choose_row_color(self):
        """Öffnet einen Farbauswahl-Dialog für die ausgewählte Zeile."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_label.config(text="Keine Zeile ausgewählt.")
            return
        
        # Farbauswahl-Dialog öffnen
        color = colorchooser.askcolor(title="Farbe für Zeile auswählen")
        if color[1]:  # Wenn eine Farbe ausgewählt wurde
            aktuelles_sheet = self.sheet_var.get()
            if aktuelles_sheet not in self.row_colors:
                self.row_colors[aktuelles_sheet] = {}
            
            # Farbe für alle ausgewählten Zeilen setzen
            for item in selected_items:
                # Zeilen-Index ermitteln
                row_index = self.tree.index(item)
                self.row_colors[aktuelles_sheet][row_index] = color[1]
                
                # Neue Tag-Konfiguration für diese Farbe erstellen
                tag_name = f"colored_row_{row_index}"
                print(f"Applying color {color[1]} to row {row_index} with tag {tag_name}")  # Debug
                
                # Tag konfigurieren
                self.tree.tag_configure(tag_name, 
                                      background=color[1],
                                      foreground=self.get_contrasting_text_color(color[1]))
                
                # Tag auf die Zeile anwenden
                self.tree.item(item, tags=(tag_name,))
                
                # Debug: Prüfen ob Tag angewendet wurde
                current_tags = self.tree.item(item, 'tags')
                print(f"Current tags for item {item}: {current_tags}")  # Debug
            
            # Farb-Informationen speichern
            self.save_row_colors()
            self.status_label.config(text=f"Farbe angewendet auf {len(selected_items)} Zeile(n).")

    def remove_row_color(self):
        """Entfernt die Farbe von der ausgewählten Zeile."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_label.config(text="Keine Zeile ausgewählt.")
            return
        
        aktuelles_sheet = self.sheet_var.get()
        if aktuelles_sheet not in self.row_colors:
            self.row_colors[aktuelles_sheet] = {}
        
        # Farbe für alle ausgewählten Zeilen entfernen
        for item in selected_items:
            row_index = self.tree.index(item)
            if row_index in self.row_colors[aktuelles_sheet]:
                del self.row_colors[aktuelles_sheet][row_index]
            
            # Ursprüngliche alternierende Farben wiederherstellen
            if row_index % 2 == 0:
                self.tree.item(item, tags=('evenrow',))
            else:
                self.tree.item(item, tags=('oddrow',))
        
        # Farb-Informationen speichern
        self.save_row_colors()
        self.status_label.config(text=f"Farbe entfernt von {len(selected_items)} Zeile(n).")

    def apply_preset_color(self, color):
        """Wendet eine voreingestellte Farbe auf die ausgewählten Zeilen an."""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_label.config(text="Keine Zeile ausgewählt.")
            return
        
        aktuelles_sheet = self.sheet_var.get()
        if aktuelles_sheet not in self.row_colors:
            self.row_colors[aktuelles_sheet] = {}
        
        # Farbe für alle ausgewählten Zeilen setzen
        for item in selected_items:
            # Zeilen-Index ermitteln
            row_index = self.tree.index(item)
            self.row_colors[aktuelles_sheet][row_index] = color
            
            # Neue Tag-Konfiguration für diese Farbe erstellen
            tag_name = f"colored_row_{row_index}"
            print(f"Applying preset color {color} to row {row_index} with tag {tag_name}")  # Debug
            
            # Tag konfigurieren
            self.tree.tag_configure(tag_name, 
                                  background=color,
                                  foreground=self.get_contrasting_text_color(color))
            
            # Tag auf die Zeile anwenden
            self.tree.item(item, tags=(tag_name,))
            
            # Debug: Prüfen ob Tag angewendet wurde
            current_tags = self.tree.item(item, 'tags')
            print(f"Current tags for item {item}: {current_tags}")  # Debug
        
        # Farb-Informationen speichern
        self.save_row_colors()
        self.status_label.config(text=f"Schnellfarbe angewendet auf {len(selected_items)} Zeile(n).")

    def setup_default_tags(self):
        """Konfiguriert die Standard-Tags für die Treeview."""
        # Tag-Konfiguration für alternierende Zeilenfarben
        # Nur unterstützte Optionen für ttk.Treeview Tags verwenden
        self.tree.tag_configure('evenrow', 
                              background=self.row_color1,
                              foreground=self.text_color)
        self.tree.tag_configure('oddrow', 
                              background=self.row_color2,
                              foreground=self.text_color)

    def get_contrasting_text_color(self, bg_color):
        """Berechnet eine kontrastierende Textfarbe basierend auf der Hintergrundfarbe."""
        # Hex-Farbe in RGB umwandeln
        bg_color = bg_color.lstrip('#')
        r, g, b = tuple(int(bg_color[i:i+2], 16) for i in (0, 2, 4))
        
        # Helligkeit berechnen (0-255)
        brightness = (r * 0.299 + g * 0.587 + b * 0.114)
        
        # Weiß für dunkle Hintergründe, Schwarz für helle
        return "#ffffff" if brightness < 128 else "#000000"

    def save_row_colors(self):
        """Speichert die Zeilenfarben zusammen mit den sortierten Daten."""
        if self.save_file:
            try:
                # Sowohl sortierte Daten als auch Farb-Informationen speichern
                combined_data = {
                    'sorted_data': self.sorted_data,
                    'row_colors': self.row_colors
                }
                with open(self.save_file, 'wb') as f:
                    pickle.dump(combined_data, f)
                print("Zeilenfarben gespeichert.")
            except Exception as e:
                print(f"Fehler beim Speichern der Zeilenfarben: {str(e)}")

    def load_row_colors(self):
        """Lädt die Zeilenfarben aus der Speicherdatei."""
        if self.save_file and os.path.exists(self.save_file):
            try:
                with open(self.save_file, 'rb') as f:
                    data = pickle.load(f)
                    
                # Überprüfen, ob es sich um das neue Format handelt
                if isinstance(data, dict) and 'sorted_data' in data:
                    self.sorted_data = data.get('sorted_data', {})
                    self.row_colors = data.get('row_colors', {})
                else:
                    # Altes Format (nur sortierte Daten)
                    self.sorted_data = data
                    self.row_colors = {}
                    
                print("Zeilenfarben geladen.")
            except Exception as e:
                print(f"Fehler beim Laden der Zeilenfarben: {str(e)}")
                self.sorted_data = {}
                self.row_colors = {}

    def create_backup_folder(self):
        """Erstellt den Backup-Ordner falls er nicht existiert."""
        if not os.path.exists(self.backup_folder):
            try:
                os.makedirs(self.backup_folder)
                print(f"Backup-Ordner erstellt: {self.backup_folder}")
            except Exception as e:
                print(f"Fehler beim Erstellen des Backup-Ordners: {str(e)}")

    def get_backup_filename(self, base_name, date=None):
        """Erstellt einen Backup-Dateinamen mit Datum."""
        if date is None:
            date = datetime.now()
        date_str = date.strftime("%Y-%m-%d")
        return os.path.join(self.backup_folder, f"{base_name}_backup_{date_str}.pkl")

    def get_excel_backup_filename(self, base_name, date=None):
        """Erstellt einen Excel-Backup-Dateinamen mit Datum."""
        if date is None:
            date = datetime.now()
        date_str = date.strftime("%Y-%m-%d")
        return os.path.join(self.backup_folder, f"{base_name}_backup_{date_str}.xlsx")

    def should_create_backup(self, base_name):
        """Prüft, ob heute bereits ein Backup erstellt wurde."""
        today_backup = self.get_backup_filename(base_name)
        return not os.path.exists(today_backup)

    def create_daily_backup(self):
        """Erstellt ein tägliches Backup der sortierten Daten und Zeilenfarben."""
        if not self.save_file or not self.sorted_data:
            return
        
        try:
            base_name = os.path.splitext(os.path.basename(self.save_file))[0]
            base_name = base_name.replace("_sorted_data", "")  # Original-Namen wiederherstellen
            
            # Prüfen, ob heute bereits ein Backup existiert
            if not self.should_create_backup(base_name):
                return  # Backup bereits vorhanden
            
            # Pickle-Backup erstellen mit Daten und Farben
            backup_file = self.get_backup_filename(base_name)
            combined_data = {
                'sorted_data': self.sorted_data,
                'row_colors': self.row_colors
            }
            with open(backup_file, 'wb') as f:
                pickle.dump(combined_data, f)
            
            # Excel-Backup erstellen (menschenlesbar)
            excel_backup = self.get_excel_backup_filename(base_name)
            self.create_excel_backup(excel_backup)
            
            # Alte Backups bereinigen (älter als 30 Tage)
            self.cleanup_old_backups(base_name)
            
            print(f"Tägliches Backup erstellt: {backup_file}")
            self.status_label.config(text=f"Tägliches Backup erstellt: {datetime.now().strftime('%d.%m.%Y')}")
            
        except Exception as e:
            print(f"Fehler beim Erstellen des täglichen Backups: {str(e)}")

    def create_excel_backup(self, excel_backup_path):
        """Erstellt ein Excel-Backup aller sortierten Sheets."""
        try:
            with pd.ExcelWriter(excel_backup_path, engine='openpyxl') as writer:
                for sheet_name, df in self.sorted_data.items():
                    # Sheet-Namen für Excel bereinigen (ungültige Zeichen entfernen)
                    safe_sheet_name = str(sheet_name)[:31]  # Excel-Limit für Sheet-Namen
                    safe_sheet_name = safe_sheet_name.replace('/', '_').replace('\\', '_').replace(':', '_')
                    df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            print(f"Excel-Backup erstellt: {excel_backup_path}")
        except Exception as e:
            print(f"Fehler beim Erstellen des Excel-Backups: {str(e)}")

    def cleanup_old_backups(self, base_name, days_to_keep=10):
        """Löscht Backups, die älter als die angegebene Anzahl von Tagen sind."""
        try:
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            
            # Pickle-Backups bereinigen
            pattern = os.path.join(self.backup_folder, f"{base_name}_backup_*.pkl")
            for backup_file in glob.glob(pattern):
                try:
                    # Datum aus Dateinamen extrahieren
                    filename = os.path.basename(backup_file)
                    date_str = filename.split('_backup_')[1].replace('.pkl', '')
                    file_date = datetime.strptime(date_str, '%Y-%m-%d')
                    
                    if file_date < cutoff_date:
                        os.remove(backup_file)
                        print(f"Altes Backup gelöscht: {backup_file}")
                except Exception as e:
                    print(f"Fehler beim Löschen von {backup_file}: {str(e)}")
            
            # Excel-Backups bereinigen
            pattern = os.path.join(self.backup_folder, f"{base_name}_backup_*.xlsx")
            for backup_file in glob.glob(pattern):
                try:
                    filename = os.path.basename(backup_file)
                    date_str = filename.split('_backup_')[1].replace('.xlsx', '')
                    file_date = datetime.strptime(date_str, '%Y-%m-%d')
                    
                    if file_date < cutoff_date:
                        os.remove(backup_file)
                        print(f"Altes Excel-Backup gelöscht: {backup_file}")
                except Exception as e:
                    print(f"Fehler beim Löschen von {backup_file}: {str(e)}")
                    
        except Exception as e:
            print(f"Fehler beim Bereinigen alter Backups: {str(e)}")

    def list_available_backups(self, base_name):
        """Listet alle verfügbaren Backups für eine Datei auf."""
        try:
            pattern = os.path.join(self.backup_folder, f"{base_name}_backup_*.pkl")
            backups = []
            for backup_file in glob.glob(pattern):
                try:
                    filename = os.path.basename(backup_file)
                    date_str = filename.split('_backup_')[1].replace('.pkl', '')
                    file_date = datetime.strptime(date_str, '%Y-%m-%d')
                    backups.append((file_date, backup_file))
                except Exception:
                    continue
            
            # Nach Datum sortieren (neueste zuerst)
            backups.sort(reverse=True)
            return backups
        except Exception as e:
            print(f"Fehler beim Auflisten der Backups: {str(e)}")
            return []



def main():
    root = tk.Tk()
    app = ExcelViewer(root)
    root.mainloop()

if __name__ == "__main__":
    main() 