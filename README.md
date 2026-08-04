# Quittungs-Generator für Schulgebühren

Ein in Python geschriebenes, GUI-basiertes Desktop-Tool zur automatisierten Massengenerierung von Schulgebühren-Quittungen. Die Anwendung liest Schüler- und Preisdaten aus Excel-Listen aus, befüllt dynamisch Word-Vorlagen und konvertiert diese anschließend über eine robuste COM-Schnittstelle in strukturierte und nummerierte PDF-Sammelbelege. 

Das Tool wurde speziell für die administrativen Anforderungen von Schulvereinen entwickelt, um den Prozess der Rechnungserstellung effizient, fehlerfrei und optisch ansprechend zu gestalten.

## 🌟 Hauptfunktionen

* **Automatisierte Word-Generierung:** Dynamisches Ersetzen von Platzhaltern in einer `.docx`-Vorlage basierend auf Excel-Daten (Schülerliste & Preisliste).
* **Parallele PDF-Konvertierung:** Schnelle Verarbeitung durch Multithreading (`concurrent.futures`), wobei bis zu 4 Word-Instanzen unsichtbar im Hintergrund arbeiten.
* **Hierarchische Strukturierung:** Automatisches Anlegen einer sauberen Ordnerstruktur nach dem Muster `Output-Ordner > Filiale > Klasse`.
* **Sammelbelege (Inline-Merge):** Sofortiges Zusammenfügen der generierten Einzel-Quittungen (inklusive Trennseiten) zu Klassen-Sammelbelegen und finalen Filial-Sammelbelegen.
* **Eindeutige Nummerierung:** Fortlaufende, übersichtliche Nummerierung der Dateinamen (z. B. `001_Sammel_PDF_Klasse_1a.pdf`, `01_West_Sammelbelege.pdf`).
* **Excel-Übersichtsbericht:** Generierung einer tabellarischen Zusammenfassung aller erstellten Belege zur einfachen Kontrolle.
* **Benutzerfreundliche GUI:** Tkinter-basierte Oberfläche mit Fortschrittsbalken, Statusanzeige und einer Abbruchfunktion ("Cancelling").

---

## 🛠️ Systemvoraussetzungen

1. **Betriebssystem:** Windows (zwingend erforderlich für die `win32com`-Schnittstelle).
2. **Microsoft Word:** Muss lokal auf dem System installiert sein (für die unsichtbare PDF-Konvertierung via COM API).
3. **Python:** Version 3.8 oder höher.

---


## 📦 Installation & Setup (Für Entwickler)

Klone das Repository und installiere die benötigten Abhängigkeiten.

1. **Repository klonen:**
   ```bash
   git clone <deine-repo-url>
   cd <repo-ordner>



2. **Abhängigkeiten installieren:**
Es wird empfohlen, eine virtuelle Umgebung (venv) zu nutzen. Die benötigten Pakete lassen sich wie folgt installieren:

    ```bash
    pip install pandas openpyxl python-docx pywin32 pypdf Pillow num2words

##  Bedienung des Tools

Das Tool ist in zwei aufeinanderfolgende Phasen unterteilt:

### Phase 1: Word generieren (📝)
Wähle die Schülerliste (Excel) aus (muss die Spalten Filiale, In Klasse, Eltern 1 - Name, Eltern 1 - Emailadresse und Name Kind enthalten).

Wähle die Preisliste (Excel) aus (Reiter: Gebuehren, Beitraege, Konfiguration).

Wähle das Quittungs-Template (.docx) aus (mit den entsprechenden {{PLATZHALTER}} Variablen).

Definiere den Ausgabeordner.

Klicke auf 1. Word generieren. Das Tool erstellt nun alle Einzel-Quittungen in Word.

### Phase 2: PDFs generieren (📄)
Sobald Phase 1 abgeschlossen ist (oder wenn die Word-Dokumente im Ausgabeordner bereits existieren), klicke auf 2. PDFs generieren.

Das Tool konvertiert die Word-Dateien parallel in PDFs.

Gleichzeitig werden Trennseiten generiert und die Dateien zu nummerierten Klassen- und Filial-Sammelbelegen zusammengefasst.

🏗️ Erstellen der ausführbaren Datei (.exe)
Um das Skript an Nutzer ohne Python-Installation weiterzugeben, kann es mit PyInstaller in eine Standalone-Applikation umgewandelt werden. Der Code enthält bereits die notwendigen Patches (z. B. DummyOutput), um ohne sichtbares Konsolenfenster fehlerfrei zu laufen.

## PyInstaller installieren:

Bash
pip install pyinstaller
Kompilieren:

Bash
pyinstaller --noconsole --onefile quittungs_generator.py
Die fertige Datei befindet sich anschließend im Ordner dist. Hinweis: Lege die .exe-Datei im selben Verzeichnis wie die Vorlagen (logo.png, preise.xlsx etc.) ab, damit diese beim Start automatisch in der GUI vorausgewählt werden.

## 📂 Daten- und Ordnerstruktur
Benötigte Eingabedateien (Standardnamen)
schuelerliste.xlsx (Die Mitglieder- und Klassendatenbank)

preise.xlsx (Zusammensetzung der Gebühren, Beiträge und Schuljahr)

Quittung-Template.docx (Die optische Vorlage)

logo.png (Für die Anzeige in der GUI, max. 500x120 Pixel)

Generierte Ausgabestruktur
Plaintext

    ```bash
    Ausgabeordner/
            ├── 2026_Quittungen_Uebersicht.xlsx
            ├── 01_Nord_Sammelbelege_2026.pdf
            ├── 02_Sued_Sammelbelege_2026.pdf
            ├── Nord/
            │   ├── 001_Sammel_PDF_Klasse_1a.pdf
            │   └── 1a/
            │       ├── 0000_Trennseite_1a.pdf
            │       ├── 2026_Quittung_001_Familie_Müller.pdf
            │       └── ...
            └── Sued/
                ├── 002_Sammel_PDF_Klasse_2b.pdf
                └── 2b/
                    └── ...

## ⚠️ Bekannte Fallstricke & Fehlerbehebung
Programm friert bei PDF-Erstellung ein: Wird durch die Integration von word.DisplayAlerts = 0 nun automatisch unterdrückt. Sollte es dennoch passieren, prüfe im Task-Manager, ob versteckte Microsoft Word-Prozesse hängen geblieben sind, und beende diese (WINWORD.EXE).

Fehlermeldung CoInitialize has not been called: Wurde im Code durch thread-lokale Initialisierung (pythoncom.CoInitialize()) beim Multithreading behoben.

Logo wird nicht angezeigt: Stelle sicher, dass das Paket Pillow in der Umgebung installiert ist.

Autor: Ivan Zlatanchev        