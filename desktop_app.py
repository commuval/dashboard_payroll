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
        
        # Data storage
        self.data = None
        self.sheet_names = []
        self.current_sheet = None
        self.sorted_data = None
        self.file_path = None
        
        # Styling
        self.style = ttk.Style()
        self.setup_styles()
        
        # UI erstellen
        self.create_widgets()
        self.setup_bindings()
        
        # Konfiguration laden
        self.load_config()
        self.load_last_file()
    
    def setup_styles(self):
        """Richtet die Styling-Konfiguration ein"""
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
        
        # Button-Styling
        self.style.configure("Custom.TButton", 
                           padding=10, 
                           font=('Segoe UI', 10),
                           background=self.accent_color,
                           focuscolor='none')  # Entfernt die blaue Umrandung
        
        self.style.map("Custom.TButton",
                      focuscolor=[('!focus', 'none')])
        
        # Label-Styling
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
        
        # Frame-Styling
        self.style.configure("TLabelframe", 
                           background=self.bg_color)
        self.style.configure("TLabelframe.Label", 
                           font=('Segoe UI', 10, 'bold'),
                           background=self.bg_color,
                           foreground=self.accent_color)
    
    def create_widgets(self):
        """Erstellt alle UI-Widgets"""
        # Hauptcontainer mit Padding
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self.create_header()
        
        # Upload-Bereich
        self.create_upload_area()
        
        # Sheet-Auswahl
        self.create_sheet_selection()
        
        # Tabelle
        self.create_table_area()
        
        # Status-Bar
        self.create_status_bar()
    
    def create_header(self):
        """Erstellt den Header-Bereich"""
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            self.header_frame,
            text="Excel Viewer Pro",
            style="Header.TLabel"
        ).pack(side=tk.LEFT)
    
    def create_upload_area(self):
        """Erstellt den Upload-Bereich"""
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
    
    def create_sheet_selection(self):
        """Erstellt die Sheet-Auswahl"""
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
    
    def create_table_area(self):
        """Erstellt den Tabellen-Bereich"""
        self.table_frame = ttk.LabelFrame(
            self.main_frame,
            text="Daten",
            padding="10"
        )
        self.table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        self.v_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        self.h_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        
        # Treeview
        self.tree = ttk.Treeview(
            self.table_frame,
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
            show='headings'
        )
        
        # Scrollbar-Bindungen
        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)
        
        # Grid-Layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Grid-Gewichte
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
    
    def create_status_bar(self):
        """Erstellt die Status-Bar"""
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="Bereit",
            style="Status.TLabel"
        )
        self.status_label.pack(side=tk.LEFT)
    
    def setup_bindings(self):
        """Richtet Event-Bindungen ein"""
        # Kontext-Menü
        self.create_context_menu()
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # Bearbeitung
        self.tree.bind("<Double-1>", self.start_edit)
    
    def create_context_menu(self):
        """Erstellt das Kontext-Menü"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Neue Zeile hinzufügen", command=self.add_new_row)
        self.context_menu.add_command(label="Zeile löschen", command=self.delete_row)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Zeilenfarbe wählen", command=self.choose_row_color)
        self.context_menu.add_command(label="Zeilenfarbe entfernen", command=self.remove_row_color)
    
    def load_excel(self):
        """Lädt eine Excel-Datei"""
        file_path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.load_excel_file(file_path)
    
    def load_excel_file(self, file_path):
        """Lädt eine spezifische Excel-Datei"""
        try:
            # Excel-Datei laden
            excel_file = pd.ExcelFile(file_path)
            self.sheet_names = excel_file.sheet_names
            
            # Alle Sheets laden
            self.data = {}
            for sheet in self.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                self.data[sheet] = df
            
            self.file_path = file_path
            self.current_sheet = self.sheet_names[0] if self.sheet_names else None
            
            # UI aktualisieren
            self.file_label.config(text=f"Geladen: {os.path.basename(file_path)}")
            self.sheet_combo['values'] = self.sheet_names
            if self.current_sheet:
                self.sheet_var.set(self.current_sheet)
            
            self.update_display()
            self.save_config()
            
            self.status_label.config(text=f"Excel-Datei geladen: {len(self.sheet_names)} Sheet(s)")
            
        except Exception as e:
            tk.messagebox.showerror("Fehler", f"Fehler beim Laden der Excel-Datei:\n{str(e)}")
    
    def update_display(self, event=None):
        """Aktualisiert die Anzeige"""
        if not self.data or not self.current_sheet:
            return
        
        # Aktuelles Sheet holen
        if event:
            self.current_sheet = self.sheet_var.get()
        
        df = self.data[self.current_sheet]
        
        # Treeview leeren
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Spalten konfigurieren
        columns = list(df.columns)
        self.tree['columns'] = columns
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, minwidth=50)
        
        # Daten hinzufügen
        for index, row in df.iterrows():
            values = [self.clean_value(val) for val in row]
            self.tree.insert('', 'end', values=values)
        
        self.status_label.config(text=f"Angezeigt: {len(df)} Zeilen, {len(df.columns)} Spalten")
    
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
    
    def sortiere_nach_praxis(self):
        """Sortiert die Daten nach Praxis"""
        if not self.data or not self.current_sheet:
            tk.messagebox.showwarning("Warnung", "Keine Daten zum Sortieren vorhanden.")
            return
        
        df = self.data[self.current_sheet].copy()
        
        try:
            # Suche nach Praxis-Spalte
            praxis_cols = [col for col in df.columns if 'praxis' in str(col).lower()]
            
            if not praxis_cols:
                tk.messagebox.showwarning("Warnung", 
                    f"Keine 'Praxis'-Spalte gefunden.\nVerfügbare Spalten: {', '.join(df.columns)}")
                return
            
            praxis_col = praxis_cols[0]
            
            # Sortieren
            sorted_df = df.sort_values(by=praxis_col, na_position='last')
            
            # Daten aktualisieren
            self.data[self.current_sheet] = sorted_df
            self.sorted_data = sorted_df
            
            # Anzeige aktualisieren
            self.update_display()
            
            # Backup erstellen
            self.create_daily_backup()
            
            tk.messagebox.showinfo("Erfolg", f"Daten erfolgreich nach '{praxis_col}' sortiert!")
            self.status_label.config(text=f"Sortiert nach: {praxis_col}")
            
        except Exception as e:
            tk.messagebox.showerror("Fehler", f"Fehler beim Sortieren:\n{str(e)}")
    
    def save_sorted_data(self):
        """Speichert die sortierten Daten"""
        if self.sorted_data is None:
            tk.messagebox.showwarning("Warnung", "Keine sortierten Daten zum Speichern vorhanden.")
            return
        
        try:
            # Dateiname generieren
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.file_path or "sortiert"))[0]
            
            # Speicherpfad auswählen
            save_path = filedialog.asksaveasfilename(
                title="Sortierte Daten speichern",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialname=f"{base_name}_sortiert_{timestamp}.xlsx"
            )
            
            if save_path:
                # Als Excel speichern
                self.sorted_data.to_excel(save_path, index=False)
                
                # Als Pickle speichern (für schnelles Laden)
                pickle_path = save_path.replace('.xlsx', '.pkl')
                with open(pickle_path, 'wb') as f:
                    pickle.dump(self.sorted_data, f)
                
                tk.messagebox.showinfo("Erfolg", f"Daten gespeichert:\n{save_path}")
                self.status_label.config(text=f"Gespeichert: {os.path.basename(save_path)}")
                
        except Exception as e:
            tk.messagebox.showerror("Fehler", f"Fehler beim Speichern:\n{str(e)}")
    
    def start_edit(self, event):
        """Startet die Bearbeitung einer Zelle"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        column = self.tree.identify_column(event.x)
        if not column:
            return
        
        # Spaltenindex ermitteln
        col_index = int(column[1:]) - 1  # column ist "#1", "#2", etc.
        
        # Aktueller Wert
        current_value = self.tree.item(item, 'values')[col_index]
        
        # Edit-Entry erstellen
        x, y, width, height = self.tree.bbox(item, column)
        
        self.edit_entry = tk.Entry(self.tree)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus()
        
        # Event-Bindungen für das Beenden der Bearbeitung
        self.edit_entry.bind('<Return>', lambda e: self.save_edit(item, column, col_index))
        self.edit_entry.bind('<Escape>', lambda e: self.cancel_edit())
        self.edit_entry.bind('<FocusOut>', lambda e: self.save_edit(item, column, col_index))
    
    def save_edit(self, item, column, col_index):
        """Speichert die Bearbeitung"""
        if not hasattr(self, 'edit_entry'):
            return
        
        new_value = self.edit_entry.get()
        
        # Wert in Treeview aktualisieren
        values = list(self.tree.item(item, 'values'))
        values[col_index] = new_value
        self.tree.item(item, values=values)
        
        # Wert in DataFrame aktualisieren
        row_index = self.tree.index(item)
        column_name = list(self.data[self.current_sheet].columns)[col_index]
        self.data[self.current_sheet].iloc[row_index, col_index] = new_value
        
        self.cancel_edit()
        self.status_label.config(text="Änderung gespeichert")
    
    def cancel_edit(self):
        """Bricht die Bearbeitung ab"""
        if hasattr(self, 'edit_entry'):
            self.edit_entry.destroy()
            del self.edit_entry
    
    def show_context_menu(self, event):
        """Zeigt das Kontext-Menü"""
        try:
            self.context_menu.post(event.x_root, event.y_root)
        except:
            pass
    
    def add_new_row(self):
        """Fügt eine neue Zeile hinzu"""
        if not self.data or not self.current_sheet:
            return
        
        df = self.data[self.current_sheet]
        new_row = pd.Series([''] * len(df.columns), index=df.columns)
        
        # Neue Zeile zum DataFrame hinzufügen
        self.data[self.current_sheet] = pd.concat([df, new_row.to_frame().T], ignore_index=True)
        
        # Anzeige aktualisieren
        self.update_display()
        self.status_label.config(text="Neue Zeile hinzugefügt")
    
    def delete_row(self):
        """Löscht die ausgewählte Zeile"""
        selected_items = self.tree.selection()
        if not selected_items:
            tk.messagebox.showwarning("Warnung", "Keine Zeile ausgewählt.")
            return
        
        if tk.messagebox.askyesno("Bestätigung", "Möchten Sie die ausgewählte Zeile wirklich löschen?"):
            for item in selected_items:
                row_index = self.tree.index(item)
                # Zeile aus DataFrame entfernen
                self.data[self.current_sheet] = self.data[self.current_sheet].drop(
                    self.data[self.current_sheet].index[row_index]
                ).reset_index(drop=True)
            
            # Anzeige aktualisieren
            self.update_display()
            self.status_label.config(text="Zeile(n) gelöscht")
    
    def choose_row_color(self):
        """Wählt eine Farbe für die ausgewählte Zeile"""
        selected_items = self.tree.selection()
        if not selected_items:
            tk.messagebox.showwarning("Warnung", "Keine Zeile ausgewählt.")
            return
        
        color = colorchooser.askcolor(title="Zeilenfarbe wählen")
        if color[1]:  # Wenn eine Farbe gewählt wurde
            for item in selected_items:
                self.tree.set(item, 'tags', (color[1],))
                self.tree.tag_configure(color[1], background=color[1])
    
    def remove_row_color(self):
        """Entfernt die Farbe von der ausgewählten Zeile"""
        selected_items = self.tree.selection()
        if not selected_items:
            tk.messagebox.showwarning("Warnung", "Keine Zeile ausgewählt.")
            return
        
        for item in selected_items:
            self.tree.set(item, 'tags', ())
    
    def create_daily_backup(self):
        """Erstellt ein tägliches Backup"""
        if not self.sorted_data is not None:
            return
        
        try:
            # Backup-Ordner erstellen
            backup_dir = "backups"
            os.makedirs(backup_dir, exist_ok=True)
            
            # Dateiname mit Zeitstempel
            timestamp = datetime.now().strftime("%Y-%m-%d")
            base_name = os.path.splitext(os.path.basename(self.file_path or "data"))[0]
            
            # Pickle Backup
            pickle_filename = f"{base_name}_backup_{timestamp}.pkl"
            pickle_path = os.path.join(backup_dir, pickle_filename)
            
            with open(pickle_path, 'wb') as f:
                pickle.dump(self.sorted_data, f)
            
            # Excel Backup
            excel_filename = f"{base_name}_backup_{timestamp}.xlsx"
            excel_path = os.path.join(backup_dir, excel_filename)
            
            self.sorted_data.to_excel(excel_path, index=False)
            
        except Exception as e:
            print(f"Backup-Fehler: {e}")
    
    def save_config(self):
        """Speichert die aktuelle Konfiguration"""
        config = {
            'last_file': self.file_path,
            'window_geometry': self.root.geometry()
        }
        
        try:
            with open('excel_viewer_config.json', 'w') as f:
                json.dump(config, f)
        except:
            pass
    
    def load_config(self):
        """Lädt die gespeicherte Konfiguration"""
        try:
            with open('excel_viewer_config.json', 'r') as f:
                config = json.load(f)
                
            if 'window_geometry' in config:
                self.root.geometry(config['window_geometry'])
                
        except:
            pass
    
    def load_last_file(self):
        """Lädt die zuletzt verwendete Datei"""
        try:
            with open('excel_viewer_config.json', 'r') as f:
                config = json.load(f)
                
            if 'last_file' in config and os.path.exists(config['last_file']):
                self.load_excel_file(config['last_file'])
                
        except:
            pass

def main():
    root = tk.Tk()
    app = ExcelViewer(root)
    root.mainloop()

if __name__ == "__main__":
    main() 