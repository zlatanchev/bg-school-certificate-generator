# -*- coding: utf-8 -*-
"""
Quittungs-Generator für Schulgebühren - Finale, automatisierte Version

**Neue Features:**
- **Automatisches Finden von Input-Dateien:** Das Skript sucht beim Start im
  eigenen Verzeichnis nach `schuelerliste.xlsx`, `preise.xlsx` und
  `Quittung-Template.docx` und füllt die Pfade automatisch aus.
- **Konfigurierbares Schuljahr:** Das Schuljahr wird nicht mehr im Code,
  sondern in der `preise.xlsx` im Tabellenblatt 'Konfiguration' festgelegt.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
import os
import sys # Wird für die Pfad-Ermittlung benötigt
from datetime import datetime
from num2words import num2words

def initialize_paths():
    """
    Sucht nach Standard-Dateinamen und füllt die GUI-Pfade vor.
    Diese Logik funktioniert sowohl für das .py-Skript als auch für die .exe-Datei.
    """
    try:
        if getattr(sys, 'frozen', False):
            # Fall 1: Anwendung läuft als kompilierte .exe (erstellt mit PyInstaller).
            # PyInstaller speichert gebündelte Dateien in einem temporären Ordner,
            # dessen Pfad in `sys._MEIPASS` liegt.
            base_path = os.path.dirname(sys.executable)
        else:
            # Fall 2: Anwendung läuft als normales .py-Skript.
            # Der Basispfad ist das Verzeichnis, in dem das Skript liegt.
            base_path = os.path.dirname(os.path.abspath(__file__))

        # Definiere die Standard-Dateinamen
        default_files = {
            "schuelerliste": "mitglieder_mit_zahlungen_und_zeilen.xlsx",
            "preise": "preise.xlsx",
            "template": "Quittung-Template.docx"
        }

        # Baue die vollen Pfade und überprüfe, ob die Dateien existieren
        path_schueler = os.path.join(base_path, default_files["schuelerliste"])
        if os.path.exists(path_schueler):
            excel_path_var.set(path_schueler)

        path_preise = os.path.join(base_path, default_files["preise"])
        if os.path.exists(path_preise):
            prices_path_var.set(path_preise)
        
        path_template = os.path.join(base_path, default_files["template"])
        if os.path.exists(path_template):
            template_path_var.set(path_template)

        out_dir = os.path.join(base_path, "out")
        output_dir_var.set(out_dir)
            
    except Exception as e:
        print(f"Fehler bei der Initialisierung der Pfade: {e}")


def docx_replace_text(doc_obj, old_text, new_text):
    """
    Ersetzt rekursiv Text in einem Word-Dokumentobjekt (Absatz oder Zelle)
    und behält dabei die ursprüngliche Formatierung bei.

    Diese Funktion durchläuft die "Runs" (formatierte Textabschnitte) und
    stellt sicher, dass Stile wie Fettdruck erhalten bleiben.
    """
    # Ersetzen in Absätzen
    for p in doc_obj.paragraphs:
        if old_text in p.text:
            inline = p.runs
            # Ersetze den Text und behalte die Formatierung des ersten Teils bei
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text
                    # Entferne den Platzhalter aus den nachfolgenden Teilen, falls er aufgeteilt war
                    for j in range(i + 1, len(inline)):
                        if old_text in inline[j].text:
                            inline[j].text = inline[j].text.replace(old_text, "")
                    break # Wichtig, um nicht mehrfach im selben Absatz zu ersetzen

    # Rekursiver Aufruf für alle Tabellen im Dokumentenobjekt
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_text(cell, old_text, new_text)



def load_prices(filepath):
    """
    Lädt die Preisinformationen UND das Schuljahr aus der Excel-Datei.
    :return: Ein Tupel (Dictionary Kindergebühren, Float Mitgliedsbeitrag, String Schuljahr).
    """
    price_sheets = pd.read_excel(filepath, sheet_name=None)
    
    fee_df = price_sheets['Gebuehren']
    fee_df.columns = fee_df.columns.str.strip() 
    child_fees = pd.Series(fee_df.Betrag.values, index=fee_df.Kind_Nr).to_dict()

    contribution_df = price_sheets['Beitraege']
    contribution_df.columns = contribution_df.columns.str.strip()
    membership_fee = contribution_df[contribution_df.Posten == 'Mitgliedsbeitrag']['Betrag'].iloc[0]
    
    # Neu: Lese das Konfigurationsblatt für das Schuljahr
    config_df = price_sheets['Konfiguration']
    config_df.columns = config_df.columns.str.strip()
    school_year = config_df[config_df.Eigenschaft == 'Schuljahr']['Wert'].iloc[0]
    
    return child_fees, float(membership_fee), str(school_year)


def generate_receipts():
    """
    Die Hauptfunktion, die den gesamten Generierungsprozess steuert.
    **Neue Logik:** Bricht bei Datenfehlern nicht ab, sondern sammelt
    Fehlermeldungen und überspringt die fehlerhaften Einträge.
    """
    excel_path = excel_path_var.get()
    template_path = template_path_var.get()
    prices_path = prices_path_var.get()
    output_dir = output_dir_var.get()
    
    if not all([excel_path, template_path, prices_path, output_dir]):
        messagebox.showerror("Fehler", "Bitte alle Pfade auswählen!")
        return

    # Liste zum Sammeln von Fehlermeldungen
    errors_found = []
    quittungs_nr = 1
    
    try:
        child_fees, membership_fee, school_year = load_prices(prices_path)
        
        df = pd.read_excel(excel_path)
        
        df.dropna(subset=['Mitglied', 'Kind'], inplace=True)
        df['Mitglied'] = df['Mitglied'].astype(str).str.strip()
        
        grouped = df.groupby('Mitglied')

        for parent_full_name, group in grouped:
            try:
                # --- Validierung für diesen spezifischen Eintrag ---
                is_group_valid = True
                # 1. Daten-Typ-Prüfung
                for index, row in group.iterrows():
                    kind_value = row['Kind']
                    if not isinstance(kind_value, str):
                        excel_row_number = index + 2
                        error_message = (
                            f"Mitglied: '{parent_full_name}'\n"
                            f"Grund: Ungültiger Datentyp in Spalte 'Kind' (Excel-Zeile {excel_row_number}).\n"
                            f"Gefunden: '{kind_value}' (Typ: {type(kind_value).__name__})"
                        )
                        errors_found.append(error_message)
                        is_group_valid = False
                        break # Nächste Prüfung für diese Gruppe nicht nötig
                
                if not is_group_valid:
                    continue # Überspringe diesen Eintrag und gehe zum nächsten in der Schleife

                # 2. Status-Prüfung
                if group['Klasse'].isin(['Abgemeldet', 'Warteliste']).any():
                    continue

                # --- Generierung (nur für valide Einträge) ---
                doc = Document(template_path)
                num_children = len(group)
                children_names = " und ".join([str(name) for name in group['Kind']])

                total_school_fee = sum(child_fees.get(i, 0) for i in range(1, num_children + 1))
                total_amount = total_school_fee + membership_fee
                
                gebuehr_wort = num2words(int(total_school_fee), lang='de')
                mitglied_wort = num2words(int(membership_fee), lang='de')
                gesamt_wort = num2words(int(total_amount), lang='de')

                replacements = {
                    "{{ELTERN_NAME}}": parent_full_name,
                    "{{KINDER_NAMEN}}": children_names, "{{NR}}": f"{quittungs_nr:03d}",
                    "{{DATUM}}": datetime.now().strftime("%d.%m.%Y"), "{{SCHULJAHR}}": school_year,
                    "{{BETRAG_GEBUEHR}}": f"{total_school_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{GESAMTBETRAG}}": f"{total_amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{BETRAG_GEBUEHR_WORT}}": f"{gebuehr_wort} Euro", "{{GESAMTBETRAG_WORT}}": f"{gesamt_wort} Euro",
                    "{{BETRAG_MITGLIED}}": f"{membership_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{BETRAG_MITGLIED_WORT}}": f"{mitglied_wort} Euro",
                }

                for old, new in replacements.items():
                    docx_replace_text(doc, old, str(new))

                parent_name = parent_full_name.replace(" ", "_")
                # Nehmen Sie die erste Klasse aus der Gruppe für den Ordnernamen
                klasse = str(group['Klasse'].iloc[0])
                outdir_class = os.path.join(output_dir, klasse)
                if not os.path.exists(outdir_class):
                    os.makedirs(outdir_class)

                output_filename = os.path.join(outdir_class, f"Quittung_{parent_name}_{quittungs_nr:03d}.docx")
                if os.path.exists(output_filename):
                    os.remove(output_filename)
                doc.save(output_filename)

                quittungs_nr += 1

            except Exception as e:
                # Fängt unerwartete Fehler für eine einzelne Gruppe ab
                error_message = f"Mitglied: '{parent_full_name}'\nGrund: Unerwarteter Fehler -> {e}"
                errors_found.append(error_message)
                continue # Überspringe diesen Eintrag

    except Exception as e:
        # Fängt kritische Fehler ab (z.B. Datei kann nicht gelesen werden)
        messagebox.showerror("Kritischer Fehler", f"Ein grundlegender Fehler hat die Verarbeitung gestoppt:\n{e}")
        return

    # --- Finale Auswertung und Meldung an den Benutzer ---
    successful_count = quittungs_nr - 1
    if not errors_found:
        messagebox.showinfo("Erfolg", f"{successful_count} Quittung(en) erfolgreich erstellt!")
    else:
        # Erstelle eine zusammenfassende Nachricht mit allen gefundenen Fehlern
        error_summary = "\n\n------------------------------------\n\n".join(errors_found)
        final_message = (
            f"{successful_count} Quittung(en) wurden erstellt.\n\n"
            f"Es gab {len(errors_found)} Fehler in der Excel-Datei. Die folgenden Einträge wurden übersprungen:\n\n"
            f"{error_summary}"
        )
        # Zeige eine Warnung statt eines Fehlers, da der Prozess teilweise erfolgreich war
        messagebox.showwarning("Vorgang abgeschlossen (mit Fehlern)", final_message)

# --- GUI Code ---
def select_excel_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls")])
    if filepath: excel_path_var.set(filepath)

def select_template_file():
    filepath = filedialog.askopenfilename(filetypes=[("Word-Dokumente", "*.docx")])
    if filepath: template_path_var.set(filepath)

def select_prices_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls")])
    if filepath: prices_path_var.set(filepath)

def select_output_dir():
    dirpath = filedialog.askdirectory()
    if dirpath: output_dir_var.set(dirpath)

# Erstelle das Hauptfenster
root = tk.Tk()
root.title("Quittungs-Generator (Auto-Detect Version)")
root.geometry("600x320")

# Erstelle die String-Variablen für die Pfade
excel_path_var, template_path_var, prices_path_var, output_dir_var = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()

# Erstelle den Haupt-Frame
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(expand=True, fill=tk.BOTH)

# Erstelle die GUI-Elemente (Widgets)
tk.Label(frame, text="1. Excel-Datei (Schülerliste) auswählen:").grid(row=0, column=0, sticky="w", pady=2)
tk.Entry(frame, textvariable=excel_path_var, width=60).grid(row=1, column=0, padx=(0, 5))
tk.Button(frame, text="Durchsuchen...", command=select_excel_file).grid(row=1, column=1)
tk.Label(frame, text="2. Excel-Datei (Preise) auswählen:").grid(row=2, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=prices_path_var, width=60).grid(row=3, column=0, padx=(0, 5))
tk.Button(frame, text="Durchsuchen...", command=select_prices_file).grid(row=3, column=1)
tk.Label(frame, text="3. Word-Vorlagendatei auswählen:").grid(row=4, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=template_path_var, width=60).grid(row=5, column=0, padx=(0, 5))
tk.Button(frame, text="Durchsuchen...", command=select_template_file).grid(row=5, column=1)
tk.Label(frame, text="4. Ausgabeordner auswählen:").grid(row=6, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=output_dir_var, width=60).grid(row=7, column=0, padx=(0, 5))
tk.Button(frame, text="Durchsuchen...", command=select_output_dir).grid(row=7, column=1)
tk.Button(frame, text="🚀 Quittungen generieren", font=("Helvetica", 12, "bold"), command=generate_receipts, bg="#4CAF50", fg="white").grid(row=8, column=0, columnspan=2, pady=20, ipadx=10, ipady=5)

# **NEU**: Rufe die Initialisierungsfunktion nach dem Erstellen der GUI auf
initialize_paths()

# Starte die Anwendung
root.mainloop()