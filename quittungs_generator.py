# -*- coding: utf-8 -*-
"""
Quittungs-Generator für Schulgebühren - Finale, stabile Version

Korrektur: Platzhalter für {{BETRAG_MITGLIED_WORT}} wurde wieder hinzugefügt.
Diese Version ist für die Verwendung mit einer VEREINFACHTEN Word-Vorlage optimiert,
bei der die Platzhalter KEINE Sonderzeichen wie Schrägstriche enthalten.
z.B. ({{BETRAG_GEBUEHR_WORT}}) statt /{{BETRAG_GEBUEHR_WORT}}/
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
import os
from datetime import datetime
from num2words import num2words

# Globale Konfiguration
SCHULJAHR = "2023/2024"

def docx_replace_text(doc_obj, old_text, new_text):
    """
    Ersetzt rekursiv Text in einem Word-Dokumentobjekt (Absatz oder Zelle)
    und behält dabei die ursprüngliche Formatierung bei.
    """
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
    """Lädt die Preisinformationen aus der Excel-Datei."""
    price_sheets = pd.read_excel(filepath, sheet_name=None)
    fee_df = price_sheets['Gebuehren']
    fee_df.columns = fee_df.columns.str.strip() 
    child_fees = pd.Series(fee_df.Betrag.values, index=fee_df.Kind_Nr).to_dict()
    contribution_df = price_sheets['Beitraege']
    contribution_df.columns = contribution_df.columns.str.strip()
    membership_fee = contribution_df[contribution_df.Posten == 'Mitgliedsbeitrag']['Betrag'].iloc[0]
    return child_fees, float(membership_fee)

def generate_receipts():
    """Die Hauptfunktion, die den gesamten Generierungsprozess steuert."""
    excel_path = excel_path_var.get()
    template_path = template_path_var.get()
    prices_path = prices_path_var.get()
    output_dir = output_dir_var.get()
    
    if not all([excel_path, template_path, prices_path, output_dir]):
        messagebox.showerror("Fehler", "Bitte alle Pfade auswählen!")
        return

    try:
        child_fees, membership_fee = load_prices(prices_path)
        df = pd.read_excel(excel_path)
        grouped = df.groupby(['Familienname', 'Vorname_Elternteil'])
        
        quittungs_nr = 1

        for (family_name, parent_first_name), group in grouped:
            doc = Document(template_path)
            num_children = len(group)
            children_names = " und ".join(group['Vorname_Kind'])
            parent_full_name = f"{parent_first_name} {family_name}"
            
            total_school_fee = sum(child_fees.get(i, 0) for i in range(1, num_children + 1))
            total_amount = total_school_fee + membership_fee

            # Erzeuge die Wörter-Versionen der Beträge
            gebuehr_wort = num2words(int(total_school_fee), lang='de')
            mitglied_wort = num2words(int(membership_fee), lang='de')
            gesamt_wort = num2words(int(total_amount), lang='de')

            # Angepasstes Dictionary für die VEREINFACHTEN Platzhalter
            replacements = {
                "{{ELTERN_NAME}}": parent_full_name,
                "{{KINDER_NAMEN}}": children_names,
                "{{NR}}": f"{quittungs_nr:03d}",
                "{{DATUM}}": datetime.now().strftime("%d.%m.%Y"),
                "{{SCHULJAHR}}": SCHULJAHR,
                "{{BETRAG_GEBUEHR}}": f"{total_school_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                "{{GESAMTBETRAG}}": f"{total_amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                "{{BETRAG_GEBUEHR_WORT}}": f"{gebuehr_wort} Euro",
                "{{GESAMTBETRAG_WORT}}": f"{gesamt_wort} Euro",
                "{{BETRAG_MITGLIED}}": f"{membership_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                # KORREKTUR: Die folgende Zeile wurde wieder hinzugefügt
                "{{BETRAG_MITGLIED_WORT}}": f"{mitglied_wort} Euro",
            }

            for old, new in replacements.items():
                docx_replace_text(doc, old, str(new))

            output_filename = os.path.join(output_dir, f"Quittung_{family_name}_{quittungs_nr:03d}.docx")
            doc.save(output_filename)
            quittungs_nr += 1

        messagebox.showinfo("Erfolg", f"{quittungs_nr - 1} Quittung(en) erfolgreich erstellt!")
    except Exception as e:
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}")

# --- GUI Code (unverändert) ---
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
root.title("Quittungs-Generator für Schulgebühren")
root.geometry("600x320")
excel_path_var, template_path_var, prices_path_var, output_dir_var = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(expand=True, fill=tk.BOTH)
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
root.mainloop()