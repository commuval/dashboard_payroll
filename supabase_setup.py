#!/usr/bin/env python3
"""
Supabase Setup Script für Dashboard & Payroll System
"""

import os
import sys
from dotenv import load_dotenv
from database import DatabaseManager
from sqlalchemy import text

def main():
    """Hauptfunktion für Supabase-Setup"""
    print("=== Supabase Setup für Dashboard & Payroll System ===\n")
    
    # Lade Umgebungsvariablen
    load_dotenv()
    
    # Prüfe DATABASE_URL
    database_url = os.getenv('DATABASE_URL')
    if not database_url:
        print("❌ DATABASE_URL nicht gefunden!")
        print("Bitte erstellen Sie eine .env Datei mit Ihrer Supabase-DATABASE_URL")
        print("Beispiel:")
        print("DATABASE_URL=postgresql://postgres:[PASSWORD]@db.[PROJECT-REF].supabase.co:5432/postgres?sslmode=require")
        return False
    
    print("✅ DATABASE_URL gefunden")
    
    # Initialisiere Datenbank-Manager
    print("\n🔌 Verbinde zu Supabase...")
    db_manager = DatabaseManager()
    
    if not db_manager.engine:
        print("❌ Datenbankverbindung fehlgeschlagen!")
        print("Bitte überprüfen Sie:")
        print("1. Ist die DATABASE_URL korrekt?")
        print("2. Ist das Passwort richtig?")
        print("3. Ist die Projekt-Referenz korrekt?")
        return False
    
    print("✅ Verbindung zu Supabase erfolgreich!")
    
    # Teste Verbindung
    try:
        with db_manager.engine.connect() as conn:
            result = conn.execute(text("SELECT version()"))
            version = result.fetchone()[0]
            print(f"✅ PostgreSQL Version: {version}")
    except Exception as e:
        print(f"❌ Verbindungstest fehlgeschlagen: {str(e)}")
        return False
    
    # Erstelle Tabellen
    print("\n📋 Erstelle Datenbank-Tabellen...")
    try:
        success = db_manager.create_tables()
        if success:
            print("✅ Tabellen erfolgreich erstellt!")
        else:
            print("❌ Fehler beim Erstellen der Tabellen!")
            return False
    except Exception as e:
        print(f"❌ Fehler beim Erstellen der Tabellen: {str(e)}")
        return False
    
    # Teste Tabellen
    print("\n🔍 Teste Tabellen...")
    try:
        with db_manager.engine.connect() as conn:
            # Prüfe excel_files Tabelle
            result = conn.execute(text("SELECT COUNT(*) FROM excel_files"))
            count = result.fetchone()[0]
            print(f"✅ excel_files Tabelle: {count} Einträge")
            
            # Prüfe sheets Tabelle
            result = conn.execute(text("SELECT COUNT(*) FROM sheets"))
            count = result.fetchone()[0]
            print(f"✅ sheets Tabelle: {count} Einträge")
            
            # Prüfe backups Tabelle
            result = conn.execute(text("SELECT COUNT(*) FROM backups"))
            count = result.fetchone()[0]
            print(f"✅ backups Tabelle: {count} Einträge")
            
    except Exception as e:
        print(f"❌ Fehler beim Testen der Tabellen: {str(e)}")
        return False
    
    print("\n🎉 Supabase Setup erfolgreich abgeschlossen!")
    print("\nNächste Schritte:")
    print("1. Starten Sie die Anwendung: python app.py")
    print("2. Besuchen Sie: http://localhost:5000")
    print("3. Laden Sie eine Excel-Datei hoch")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 