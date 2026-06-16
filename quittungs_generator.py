# -*- coding: utf-8 -*-
"""
Quittungs-Generator für Schulgebühren - Finale Version mit flexibler, linksbündiger GUI und Info-Feld
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
import os
import sys
from datetime import datetime
from num2words import num2words

# Versuche Pillow zu importieren (für die Bildskalierung)
try:
    from PIL import Image, ImageTk
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

def initialize_paths():
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        default_files = {
            "schuelerliste": "schuelerliste.xlsx",
            "preise": "preise.xlsx",
            "template": "Quittung-Template.docx",
            "logo": "logo.png"
        }

        path_schueler = os.path.join(base_path, default_files["schuelerliste"])
        if os.path.exists(path_schueler):
            excel_path_var.set(path_schueler)

        path_preise = os.path.join(base_path, default_files["preise"])
        if os.path.exists(path_preise):
            prices_path_var.set(path_preise)
        
        path_template = os.path.join(base_path, default_files["template"])
        if os.path.exists(path_template):
            template_path_var.set(path_template)
            
        path_logo = os.path.join(base_path, default_files["logo"])
        if os.path.exists(path_logo):
            logo_path_var.set(path_logo)

        out_dir = os.path.join(base_path, "out")
        output_dir_var.set(out_dir)
            
    except Exception as e:
        print(f"Fehler bei der Initialisierung der Pfade: {e}")

def docx_replace_text(doc_obj, old_text, new_text):
    for p in doc_obj.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text
                    for j in range(i + 1, len(inline)):
                        if old_text in inline[j].text:
                            inline[j].text = inline[j].text.replace(old_text, "")
                    break 

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_text(cell, old_text, new_text)

def load_prices(filepath):
    price_sheets = pd.read_excel(filepath, sheet_name=None)
    
    fee_df = price_sheets['Gebuehren']
    fee_df.columns = fee_df.columns.str.strip() 
    child_fees = pd.Series(fee_df.Betrag.values, index=fee_df.Kind_Nr).to_dict()

    contribution_df = price_sheets['Beitraege']
    contribution_df.columns = contribution_df.columns.str.strip()
    membership_fee = contribution_df[contribution_df.Posten == 'Mitgliedsbeitrag']['Betrag'].iloc[0]
    
    config_df = price_sheets['Konfiguration']
    config_df.columns = config_df.columns.str.strip()
    school_year = config_df[config_df.Eigenschaft == 'Schuljahr']['Wert'].iloc[0]
    
    return child_fees, float(membership_fee), str(school_year)

def generate_receipts():
    excel_path = excel_path_var.get()
    template_path = template_path_var.get()
    prices_path = prices_path_var.get()
    output_dir = output_dir_var.get()
    
    if not all([excel_path, template_path, prices_path, output_dir]):
        messagebox.showerror("Fehler", "Bitte alle Pfade auswählen!")
        return

    errors_found = []
    quittungs_nr = 1
    
    try:
        child_fees, membership_fee, school_year = load_prices(prices_path)
        df = pd.read_excel(excel_path)
        df.dropna(subset=['Eltern 1 - Emailadresse', 'Name Kind'], inplace=True)
        df['Eltern 1 - Emailadresse'] = df['Eltern 1 - Emailadresse'].astype(str).str.strip()
        grouped = df.groupby('Eltern 1 - Emailadresse')

        for parent_full_name, group in grouped:
            try:
                is_group_valid = True
                for index, row in group.iterrows():
                    kind_value = row['Name Kind']
                    parent_full_name = row['Eltern 1 - Name']
                    if not isinstance(kind_value, str):
                        excel_row_number = index + 2
                        error_message = (
                            f"Mitglied: '{parent_full_name}'\n"
                            f"Grund: Ungültiger Datentyp in Spalte 'Name Kind' (Excel-Zeile {excel_row_number}).\n"
                            f"Gefunden: '{kind_value}' (Typ: {type(kind_value).__name__})"
                        )
                        errors_found.append(error_message)
                        is_group_valid = False
                        break 
                
                if not is_group_valid:
                    continue 

                if group['In Klasse'].isin(['Warteliste','', ' ']).any():
                    continue

                doc = Document(template_path)
                num_children = len(group)
                
                kinder_liste = [str(name) for name in group['Name Kind']]
                if len(kinder_liste) > 2:
                    children_names = ", ".join(kinder_liste[:-1]) + " und " + kinder_liste[-1]
                else:
                    children_names = " und ".join(kinder_liste)

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
                klasse = str(group['In Klasse'].iloc[0])
                outdir_class = os.path.join(output_dir, klasse)
                if not os.path.exists(outdir_class):
                    os.makedirs(outdir_class)

                output_filename = os.path.join(outdir_class, f"Quittung_{parent_name}_{quittungs_nr:03d}.docx")
                if os.path.exists(output_filename):
                    os.remove(output_filename)
                doc.save(output_filename)

                quittungs_nr += 1

            except Exception as e:
                error_message = f"Mitglied: '{parent_full_name}'\nGrund: Unerwarteter Fehler -> {e}"
                errors_found.append(error_message)
                continue 

    except Exception as e:
        messagebox.showerror("Kritischer Fehler", f"Ein grundlegender Fehler hat die Verarbeitung gestoppt:\n{e}")
        return

    successful_count = quittungs_nr - 1
    if not errors_found:
        messagebox.showinfo("Erfolg", f"{successful_count} Quittung(en) erfolgreich erstellt!")
    else:
        error_summary = "\n\n------------------------------------\n\n".join(errors_found)
        final_message = (
            f"{successful_count} Quittung(en) wurden erstellt.\n\n"
            f"Es gab {len(errors_found)} Fehler in der Excel-Datei. Die folgenden Einträge wurden übersprungen:\n\n"
            f"{error_summary}"
        )
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

root = tk.Tk()
root.title("Quittungs-Generator (Auto-Detect Version)")
# Höhe leicht erhöht, um Platz für das Info-Feld zu schaffen
root.geometry("600x500") 

excel_path_var = tk.StringVar()
template_path_var = tk.StringVar()
prices_path_var = tk.StringVar()
output_dir_var = tk.StringVar()
logo_path_var = tk.StringVar()

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(expand=True, fill=tk.X) 

frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=0)

initialize_paths()

# --- Logo laden, skalieren und einbinden ---
logo_path = logo_path_var.get()
if os.path.exists(logo_path):
    if HAS_PILLOW:
        try:
            img = Image.open(logo_path)
            max_width = 500
            max_height = 120
            try:
                resample_filter = Image.Resampling.LANCZOS 
            except AttributeError:
                resample_filter = Image.ANTIALIAS 

            img.thumbnail((max_width, max_height), resample_filter)
            logo_img = ImageTk.PhotoImage(img)
            root.logo_img = logo_img 
            
            tk.Label(frame, image=logo_img).grid(row=0, column=0, columnspan=2, pady=(0, 15))
        except Exception as e:
            print(f"Konnte Logo nicht verarbeiten: {e}")
            tk.Label(frame, text="[Fehler bei der Logo-Verarbeitung]").grid(row=0, column=0, columnspan=2, pady=(0, 15))
    else:
        tk.Label(frame, text="[Bitte 'Pillow' installieren (pip install Pillow) für Logo-Skalierung]", fg="red").grid(row=0, column=0, columnspan=2, pady=(0, 15))

# Eingabefelder und Buttons
tk.Label(frame, text="1. Excel-Datei (Schülerliste) auswählen:").grid(row=1, column=0, sticky="w", pady=2)
tk.Entry(frame, textvariable=excel_path_var, width=60).grid(row=2, column=0, padx=(0, 5), sticky="ew")
tk.Button(frame, text="Durchsuchen...", command=select_excel_file).grid(row=2, column=1)

tk.Label(frame, text="2. Excel-Datei (Preise) auswählen:").grid(row=3, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=prices_path_var, width=60).grid(row=4, column=0, padx=(0, 5), sticky="ew")
tk.Button(frame, text="Durchsuchen...", command=select_prices_file).grid(row=4, column=1)

tk.Label(frame, text="3. Word-Vorlagendatei auswählen:").grid(row=5, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=template_path_var, width=60).grid(row=6, column=0, padx=(0, 5), sticky="ew")
tk.Button(frame, text="Durchsuchen...", command=select_template_file).grid(row=6, column=1)

tk.Label(frame, text="4. Ausgabeordner auswählen:").grid(row=7, column=0, sticky="w", pady=(10, 2))
tk.Entry(frame, textvariable=output_dir_var, width=60).grid(row=8, column=0, padx=(0, 5), sticky="ew")
tk.Button(frame, text="Durchsuchen...", command=select_output_dir).grid(row=8, column=1)

# Haupt-Button
tk.Button(frame, text="🚀 Quittungen generieren", font=("Helvetica", 12, "bold"), command=generate_receipts, bg="#4CAF50", fg="white").grid(row=9, column=0, columnspan=2, pady=(20, 10), ipadx=10, ipady=5)

# --- NEU: Info-Feld mit Versionshinweis und Autor ---
tk.Label(frame, text="Version 17.06.2026; I. Zlat.", font=("Helvetica", 8), fg="gray").grid(row=10, column=0, columnspan=2, pady=(0, 5))

root.mainloop()