# -*- coding: utf-8 -*-
"""
Quittungs-Generator für Schulgebühren - Threading, Cancelling & PDF/Word Edition
Inklusive Excel-Übersichtstabelle, Filial-Sortierung, robuster COM-Parallelisierung 
mit sofortigem Sammelbeleg-Export in die übergeordnete Ordnerstruktur.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import sys
import shutil
import threading
import concurrent.futures
from datetime import datetime
from num2words import num2words

# --- NEU: FIX FÜR DIE .EXE DATEI (NOCONSOLE) ---
class DummyOutput:
    def write(self, x): pass
    def flush(self): pass

if sys.stdout is None:
    sys.stdout = DummyOutput()
if sys.stderr is None:
    sys.stderr = DummyOutput()
# -----------------------------------------------

# --- IMPORTE FÜR PDF ---
try:
    import win32com.client 
    from pypdf import PdfWriter
    import pythoncom  
    HAS_PDF_TOOLS = True
except ImportError:
    HAS_PDF_TOOLS = False

try:
    from PIL import Image, ImageTk
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

cancel_flag = False

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

def toggle_buttons(running=False):
    state = tk.DISABLED if running else tk.NORMAL
    cancel_state = tk.NORMAL if running else tk.DISABLED
    
    btn_generate_word.config(state=state)
    btn_generate_pdf.config(state=state)
    btn_cancel.config(state=cancel_state)

def cancel_process():
    global cancel_flag
    cancel_flag = True
    status_var.set("Abbruch wird eingeleitet... Bitte warten.")
    btn_cancel.config(state=tk.DISABLED)

def show_message_threadsafe(title, msg, is_error=False, is_warning=False):
    if is_error:
        messagebox.showerror(title, msg)
    elif is_warning:
        messagebox.showwarning(title, msg)
    else:
        messagebox.showinfo(title, msg)


# ==========================================
# PHASE 1: WORD-DOKUMENTE GENERIEREN
# ==========================================
def start_word_generation():
    global cancel_flag
    cancel_flag = False
    
    excel_path = excel_path_var.get()
    template_path = template_path_var.get()
    prices_path = prices_path_var.get()
    output_dir = output_dir_var.get()
    
    if not all([excel_path, template_path, prices_path, output_dir]):
        messagebox.showerror("Fehler", "Bitte alle Pfade auswählen!")
        return

    toggle_buttons(running=True)
    progress_var.set(0)
    
    threading.Thread(target=generate_word_receipts_task, args=(excel_path, template_path, prices_path, output_dir), daemon=True).start()

def generate_word_receipts_task(excel_path, template_path, prices_path, output_dir):
    global cancel_flag
    errors_found = []
    class_folders = set()
    quittungs_nr = 1 
    
    summary_data = []
    
    try:
        child_fees, membership_fee, school_year = load_prices(prices_path)
        
        df = pd.read_excel(excel_path)
        
        if 'Filiale' not in df.columns:
            root.after(0, show_message_threadsafe, "Fehlendes Feld", "Die Spalte 'Filiale' fehlt in der Excel-Liste! Bitte ergänze diese Spalte.", True)
            root.after(0, toggle_buttons, False)
            return

        df.dropna(subset=['Eltern 1 - Emailadresse', 'Name Kind'], inplace=True)
        df['Eltern 1 - Emailadresse'] = df['Eltern 1 - Emailadresse'].astype(str).str.strip()
        df['Filiale'] = df['Filiale'].astype(str).str.strip()
        
        df['In Klasse Sortierung'] = df['In Klasse'].astype(str)
        df.sort_values(by=['Filiale', 'In Klasse Sortierung'], inplace=True)
        
        grouped = df.groupby('Eltern 1 - Emailadresse', sort=False)
        total_parents = len(grouped)
        
        root.after(0, lambda: progress_bar.config(maximum=total_parents))
        root.after(0, lambda: status_var.set(f"Starte Generierung von {total_parents} Word-Quittungen..."))
        
        current_progress = 0

        for parent_email, group in grouped:
            if cancel_flag: break

            try:
                parent_full_name = str(group['Eltern 1 - Name'].iloc[0]).strip()
                
                is_group_valid = True
                for index, row in group.iterrows():
                    kind_value = row['Name Kind']
                    if not isinstance(kind_value, str):
                        excel_row_number = index + 2
                        errors_found.append(f"Mitglied: '{parent_full_name}' ({parent_email})\nGrund: Ungültiger Datentyp in Spalte 'Name Kind' (Zeile {excel_row_number}).")
                        is_group_valid = False
                        break 
                
                if not is_group_valid: continue 
                if group['In Klasse'].isin(['Warteliste', '', ' ']).any(): continue

                doc = Document(template_path)
                num_children = len(group)
                
                kinder_liste = [str(name) for name in group['Name Kind']]
                if len(kinder_liste) > 2:
                    children_names = ", ".join(kinder_liste[:-1]) + " und " + kinder_liste[-1]
                else:
                    children_names = " und ".join(kinder_liste)

                total_school_fee = sum(child_fees.get(i, 0) for i in range(1, num_children + 1))
                total_amount = total_school_fee + membership_fee
                
                filiale = str(group['Filiale'].iloc[0])
                klasse = str(group['In Klasse'].iloc[0])
                
                safe_filiale = filiale.replace("/", "_").replace("\\", "_")
                safe_klasse = klasse.replace("/", "_").replace("\\", "_")
                eindeutige_nummer = f"{quittungs_nr:03d}"
                
                replacements = {
                    "{{ELTERN_NAME}}": parent_full_name,
                    "{{KINDER_NAMEN}}": children_names, 
                    "{{NR}}": eindeutige_nummer, 
                    "{{DATUM}}": datetime.now().strftime("%d.%m.%Y"), 
                    "{{SCHULJAHR}}": str(school_year),
                    "{{BETRAG_GEBUEHR}}": f"{total_school_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{GESAMTBETRAG}}": f"{total_amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{BETRAG_GEBUEHR_WORT}}": f"{num2words(int(total_school_fee), lang='de')} Euro", 
                    "{{GESAMTBETRAG_WORT}}": f"{num2words(int(total_amount), lang='de')} Euro",
                    "{{BETRAG_MITGLIED}}": f"{membership_fee:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{BETRAG_MITGLIED_WORT}}": f"{num2words(int(membership_fee), lang='de')} Euro",
                }

                for old, new in replacements.items():
                    docx_replace_text(doc, old, str(new))

                outdir_class = os.path.join(output_dir, safe_filiale, safe_klasse)
                if not os.path.exists(outdir_class):
                    os.makedirs(outdir_class)
                class_folders.add(outdir_class)

                safe_parent_name = parent_full_name.replace(" ", "_").replace("/", "_")
                jahr_erstellung = datetime.now().strftime("%Y")
                dateiname = f"{jahr_erstellung}_Quittung_{eindeutige_nummer}_{safe_parent_name}.docx"
                output_filename = os.path.join(outdir_class, dateiname)
                
                doc.save(output_filename)

                summary_data.append({
                    'Quittung Nr.': eindeutige_nummer,
                    'Filiale': filiale,
                    'Eltern 1 - Name': parent_full_name,
                    'Eltern 1 - Emailadresse': parent_email,
                    'In Klasse': klasse,
                    'Rechnungsbeitrag (€)': total_amount,
                    'Namen aller Kinder': children_names
                })

                quittungs_nr += 1
                current_progress += 1
                root.after(0, progress_var.set, current_progress)
                root.after(0, status_var.set, f"Erstelle Word-Dokumente... ({current_progress}/{total_parents})")

            except Exception as e:
                errors_found.append(f"Mitglied: '{parent_email}'\nGrund: Unerwarteter Fehler -> {e}")
                continue 

        if cancel_flag:
            root.after(0, status_var.set, "Prozess durch Benutzer abgebrochen.")
            return

        if summary_data:
            try:
                summary_df = pd.DataFrame(summary_data)
                jahr_erstellung = datetime.now().strftime("%Y")
                summary_file = os.path.join(output_dir, f"{jahr_erstellung}_Quittungen_Uebersicht.xlsx")
                summary_df.to_excel(summary_file, index=False)
            except Exception as e:
                errors_found.append(f"Fehler beim Erstellen der Übersichtstabelle: {e}")

        root.after(0, progress_var.set, total_parents)
        generierte_quittungen = quittungs_nr - 1
        anzahl_klassen = len(class_folders)

        zusammenfassung = (
            f"Statistik:\n"
            f"➜ {generierte_quittungen} Quittungen (Familien) erstellt.\n"
            f"➜ Aufgeteilt in {anzahl_klassen} verschiedene Klassen (nach Filialen sortiert).\n"
            f"➜ Übersichtstabelle (Excel) wurde generiert."
        )
        
        if not errors_found:
            root.after(0, status_var.set, "Word-Generierung erfolgreich abgeschlossen!")
            root.after(0, show_message_threadsafe, "Schritt 1 abgeschlossen", f"Word-Dateien erfolgreich generiert!\n\n{zusammenfassung}")
        else:
            root.after(0, status_var.set, "Mit Warnungen abgeschlossen.")
            error_summary = "\n\n------------------------------------\n\n".join(errors_found)
            final_message = f"Word-Dateien wurden generiert.\n\n{zusammenfassung}\n\nEs gab jedoch Probleme/Fehler:\n\n{error_summary}"
            root.after(0, show_message_threadsafe, "Word-Generierung (mit Warnungen)", final_message, False, True)

    except Exception as e:
        root.after(0, status_var.set, "Kritischer Fehler aufgetreten!")
        root.after(0, show_message_threadsafe, "Kritischer Fehler", f"Ein Fehler hat die Verarbeitung gestoppt:\n{e}", True)
    finally:
        root.after(0, toggle_buttons, False)


# ==========================================
# PHASE 2: PARALLELE PDF-GENERIERUNG (COM API)
# ==========================================

# Thread-sichere Funktion zur Verarbeitung eines kompletten Klassen-Ordners
def convert_class_folder_to_pdf(class_path, filiale_path, filiale_folder, klasse_name, klasse_nummer):
    global cancel_flag
    errors = []
    word = None
    erstelltes_sammel_pdf = None
    pdfs_merged_count = 0
    
    # WICHTIG: COM für diesen spezifischen Thread initialisieren
    pythoncom.CoInitialize() 
    
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  
        
        # 1. Trennseite generieren (falls nicht existent)
        trenn_docx = os.path.join(class_path, f"0000_Trennseite_{klasse_name}.docx")
        if not os.path.exists(trenn_docx):
            t_doc = Document()
            p = t_doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"\n\n\n\n\n\nFiliale: {filiale_folder}\nKlasse: {klasse_name}")
            run.bold = True
            run.font.size = Pt(40)
            t_doc.save(trenn_docx)
        
        # 2. Ordner komplett in PDFs umwandeln
        docx_files = [f for f in os.listdir(class_path) if f.endswith('.docx') and not f.startswith('~')]
        for d_file in docx_files:
            if cancel_flag: break
            
            in_path = os.path.abspath(os.path.join(class_path, d_file))
            out_path = os.path.abspath(os.path.join(class_path, d_file.replace('.docx', '.pdf')))
            
            if not os.path.exists(out_path):
                try:
                    doc = word.Documents.Open(in_path, ReadOnly=True)
                    doc.SaveAs(out_path, FileFormat=17) # 17 = PDF
                    doc.Close(0) # 0 = wdDoNotSaveChanges
                except Exception as e:
                    errors.append(f"Fehler bei Konvertierung von {d_file}: {e}")

        # 3. DIREKTES ZUSAMMENFÜHREN (im Thread)
        if not cancel_flag:
            class_merger = PdfWriter()
            # Lade alle frisch generierten PDFs aus dem Klassenordner
            pdf_files = sorted([f for f in os.listdir(class_path) if f.endswith('.pdf')])
            
            if pdf_files:
                for pdf in pdf_files:
                    pdf_path = os.path.join(class_path, pdf)
                    class_merger.append(pdf_path)
                    pdfs_merged_count += 1
                
                # Sammel-PDF eine Ebene höher (im filiale_path) abspeichern!
                dateiname_klasse_sammel = f"{klasse_nummer:03d}_Sammel_PDF_{klasse_name}.pdf"
                class_pdf_path = os.path.join(filiale_path, dateiname_klasse_sammel)
                
                class_merger.write(class_pdf_path)
                class_merger.close()
                erstelltes_sammel_pdf = class_pdf_path
                    
    except Exception as e:
        errors.append(f"Ordner-Fehler (Klasse: {klasse_name}): {e}")
    finally:
        if word is not None:
            try:
                word.Quit()
            except:
                pass
        # COM für Thread sauber beenden
        pythoncom.CoUninitialize()
        
    return filiale_folder, klasse_name, erstelltes_sammel_pdf, pdfs_merged_count, errors


def start_pdf_generation():
    global cancel_flag
    
    if not HAS_PDF_TOOLS:
        messagebox.showerror("Fehlende Pakete", "Bitte installiere die PDF-Erweiterungen im Terminal:\n\npip install pypdf pywin32")
        return

    output_dir = output_dir_var.get()
    if not output_dir:
        messagebox.showerror("Fehler", "Bitte den Ausgabeordner auswählen!")
        return

    cancel_flag = False
    toggle_buttons(running=True)
    progress_var.set(0)
    
    threading.Thread(target=generate_pdf_receipts_task, args=(output_dir,), daemon=True).start()

def generate_pdf_receipts_task(output_dir):
    global cancel_flag
    errors_found = []
    
    # Sammelt alle Klassen-Sammelbelege (geordnet nach Filiale) für den finalen Schritt
    filial_sammel_pdfs_map = {} 
    total_pdfs_merged = 0

    try:
        filiale_folders = [f for f in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, f))]

        if not filiale_folders:
            root.after(0, status_var.set, "Warte auf Start...")
            root.after(0, show_message_threadsafe, "Info", "Keine Filial-Ordner gefunden. Bitte führe Schritt 1 aus.")
            return

        # Alle Aufgaben (Klassen) mit einer fortlaufenden Nummer einsammeln
        tasks = []
        global_class_counter = 1
        
        for f_folder in filiale_folders:
            f_path = os.path.join(output_dir, f_folder)
            class_folders = sorted([c for c in os.listdir(f_path) if os.path.isdir(os.path.join(f_path, c))])
            for c_folder in class_folders:
                c_path = os.path.join(f_path, c_folder)
                # Task = (Klassenpfad, Filialpfad, Filialname, Klassenname, Laufende_Nummer)
                tasks.append((c_path, f_path, f_folder, c_folder, global_class_counter))
                global_class_counter += 1

        anzahl_klassen = len(tasks)
        anzahl_filialen = len(filiale_folders)
        
        # Fortschritt = Alle Klassen (PDF+Merge) + Finales Filial-Merge
        root.after(0, lambda: progress_bar.config(maximum=anzahl_klassen + anzahl_filialen))
        current_progress = 0

        # =========================================================
        # SCHRITT A: Parallele PDF-Konvertierung & Klassen-Merge
        # =========================================================
        root.after(0, status_var.set, f"Erstelle PDFs und Sammelbelege (max. 4 Threads)...")
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(convert_class_folder_to_pdf, t[0], t[1], t[2], t[3], t[4]): t for t in tasks}
            
            for future in concurrent.futures.as_completed(futures):
                if cancel_flag:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                
                f_folder, c_folder, class_pdf_path, merged_count, errs = future.result()
                
                if errs:
                    errors_found.extend(errs)
                    
                if class_pdf_path:
                    total_pdfs_merged += merged_count
                    if f_folder not in filial_sammel_pdfs_map:
                        filial_sammel_pdfs_map[f_folder] = []
                    filial_sammel_pdfs_map[f_folder].append(class_pdf_path)
                
                current_progress += 1
                root.after(0, progress_var.set, current_progress)
                root.after(0, status_var.set, f"Abgeschlossen: Filiale {f_folder}, Klasse {c_folder}...")

        if cancel_flag:
            root.after(0, status_var.set, "Prozess durch Benutzer abgebrochen.")
            return

        # =========================================================
        # SCHRITT B: Finale Sammel-PDFs pro Filiale generieren
        # =========================================================
        root.after(0, status_var.set, "Erstelle abschließende Filial-Sammelbelege...")
        erstellte_sammelbelege = []
        global_filial_counter = 1

        for filiale_folder in filiale_folders:
            if cancel_flag: break
            
            if filiale_folder in filial_sammel_pdfs_map and filial_sammel_pdfs_map[filiale_folder]:
                filiale_merger = PdfWriter()
                
                # PDFs sind durch den Präfix (001_, 002_) beim Einsammeln leicht sortierbar
                sorted_class_pdfs = sorted(filial_sammel_pdfs_map[filiale_folder])
                
                for class_pdf in sorted_class_pdfs:
                    filiale_merger.append(class_pdf)
                
                jahr_erstellung = datetime.now().strftime("%Y")
                dateiname_sammel = f"{global_filial_counter:02d}_{filiale_folder}_Sammelbelege_{jahr_erstellung}.pdf"
                final_pdf_path = os.path.join(output_dir, dateiname_sammel)
                
                filiale_merger.write(final_pdf_path)
                filiale_merger.close()
                erstellte_sammelbelege.append(dateiname_sammel)
                global_filial_counter += 1
                
            current_progress += 1
            root.after(0, progress_var.set, current_progress)

        if cancel_flag:
            root.after(0, status_var.set, "Prozess durch Benutzer abgebrochen.")
            return

        # Abschlussmeldung
        root.after(0, progress_var.set, anzahl_klassen + anzahl_filialen)
        
        sammelbelege_text = "\n".join([f"   ➜ {name}" for name in erstellte_sammelbelege])
        
        zusammenfassung = (
            f"Statistik:\n"
            f"➜ {total_pdfs_merged} Einzel-Seiten (Quittungen inkl. Trennseiten) verarbeitet.\n"
            f"➜ Klassen-Sammel-PDFs wurden direkt in den übergeordneten Filial-Ordnern gespeichert.\n"
            f"➜ {len(erstellte_sammelbelege)} Filial-Sammel-PDF(s) erfolgreich im Hauptordner erstellt:\n"
            f"{sammelbelege_text}"
        )

        if not errors_found:
            root.after(0, status_var.set, "PDF-Sammelquittungen erfolgreich generiert!")
            root.after(0, show_message_threadsafe, "Schritt 2 abgeschlossen", f"PDF-Prozess erfolgreich beendet!\n\n{zusammenfassung}")
        else:
            root.after(0, status_var.set, "Mit Warnungen abgeschlossen.")
            error_summary = "\n\n------------------------------------\n\n".join(errors_found)
            final_message = f"PDF-Sammelquittungen wurden generiert.\n\n{zusammenfassung}\n\nEs gab jedoch Probleme/Fehler:\n\n{error_summary}"
            root.after(0, show_message_threadsafe, "PDF-Generierung (mit Warnungen)", final_message, False, True)

    except Exception as e:
        root.after(0, status_var.set, "Kritischer Fehler aufgetreten!")
        root.after(0, show_message_threadsafe, "Kritischer Fehler", f"Ein Fehler hat die PDF-Verarbeitung gestoppt:\n{e}", True)
    finally:
        root.after(0, toggle_buttons, False)


# --- GUI Code ---
def select_excel_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls")])
    if filepath:
        excel_path_var.set(filepath)

def select_template_file():
    filepath = filedialog.askopenfilename(filetypes=[("Word-Dokumente", "*.docx")])
    if filepath:
        template_path_var.set(filepath)

def select_prices_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls")])
    if filepath:
        prices_path_var.set(filepath)

def select_output_dir():
    dirpath = filedialog.askdirectory()
    if dirpath:
        output_dir_var.set(dirpath)

root = tk.Tk()
root.title("Quittungs-Generator (PDF & Word Edition - Parallel)")
root.geometry("650x550") 

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
            tk.Label(frame, text="[Fehler bei der Logo-Verarbeitung]").grid(row=0, column=0, columnspan=2, pady=(0, 15))
    else:
        tk.Label(frame, text="[Bitte 'Pillow' installieren für Logo-Skalierung]", fg="red").grid(row=0, column=0, columnspan=2, pady=(0, 15))

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

# Frame für die Steuerungsknöpfe
button_frame = tk.Frame(frame)
button_frame.grid(row=9, column=0, columnspan=2, pady=(20, 5))

btn_generate_word = tk.Button(button_frame, text="📝 1. Word generieren", font=("Helvetica", 11, "bold"), command=start_word_generation, bg="#2196F3", fg="white")
btn_generate_word.pack(side=tk.LEFT, padx=5, ipadx=5, ipady=5)

btn_generate_pdf = tk.Button(button_frame, text="📄 2. PDFs generieren", font=("Helvetica", 11, "bold"), command=start_pdf_generation, bg="#4CAF50", fg="white")
btn_generate_pdf.pack(side=tk.LEFT, padx=5, ipadx=5, ipady=5)

btn_cancel = tk.Button(button_frame, text="🛑 Abbrechen", font=("Helvetica", 11, "bold"), command=cancel_process, bg="#cc3025", fg="white", state=tk.DISABLED)
btn_cancel.pack(side=tk.LEFT, padx=5, ipadx=5, ipady=5)

# Status-Text
status_var = tk.StringVar()
status_var.set("Warte auf Start...")
tk.Label(frame, textvariable=status_var, fg="blue", font=("Helvetica", 10)).grid(row=10, column=0, columnspan=2, pady=(0, 5))

# Fortschrittsbalken
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(frame, variable=progress_var, mode='determinate')
progress_bar.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(0, 15))

# Info-Feld
tk.Label(frame, text="Version 04.08.2026; I. Zlat. (Parallel, Inline-Merge & Hierarchie)", font=("Helvetica", 8), fg="gray").grid(row=12, column=0, columnspan=2, pady=(0, 5))

root.mainloop()