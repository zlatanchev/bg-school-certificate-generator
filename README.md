# Quittungs-Generator für Schulgebühren

Ein Desktop-Werkzeug mit grafischer Benutzeroberfläche (GUI) zur automatisierten Erstellung von personalisierten Quittungen. Das Tool liest Schüler- und Preisdaten aus Excel-Dateien, fügt diese in eine Word-Vorlage ein und generiert für jede Familie eine separate Quittungsdatei.

*(Optional: Sie können hier einen Screenshot Ihrer Anwendung einfügen, um die README-Datei visueller zu gestalten)*

## ✨ Features

  - **Grafische Benutzeroberfläche (GUI):** Einfache Bedienung über ein Fenster, keine Notwendigkeit, den Code zu bearbeiten.
  - **Dynamische Preisberechnung:** Die Schulgebühren werden automatisch anhand der Anzahl der Kinder pro Familie berechnet.
  - **Externe Konfiguration:** Alle wichtigen Parameter (Preise, Mitgliedsbeiträge, Schuljahr) werden flexibel in einer separaten Excel-Datei verwaltet.
  - **Automatisierte Dokumentenerstellung:** Generiert `.docx`-Dateien basierend auf einer anpassbaren Word-Vorlage.
  - **Formatierungserhalt:** Behält die Formatierungen (Fettdruck, Schriftarten etc.) aus der Word-Vorlage bei.
  - **Zahl-zu-Wort-Umwandlung:** Rechnet Zahlbeträge automatisch in ausgeschriebenen Text um (z.B. 495 -\> "vierhundertfünfundneunzig").
  - **Automatische Dateierkennung:** Sucht beim Start nach den Standard-Dateinamen im Programmverzeichnis und füllt die Pfade automatisch aus, um die Bedienung zu beschleunigen.

## 📋 Anforderungen

  - Python 3.8 oder neuer
  - Die in der `requirements.txt` aufgeführten Python-Bibliotheken

## ⚙️ Installation (für Entwickler)

1.  **Klonen Sie das Repository (falls zutreffend):**

    ```bash
    git clone https://github.com/ihr-benutzername/ihr-projekt.git
    cd ihr-projekt
    ```

2.  **Erstellen Sie eine virtuelle Umgebung (empfohlen):**

    ```bash
    python -m venv venv
    # Windows
    .\venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **Installieren Sie die benötigten Bibliotheken:**
    Der folgende Befehl installiert alle notwendigen externen Module.

    ```bash
    pip install pandas python-docx num2words openpyxl
    ```

## 📂 Benötigte Dateien & Struktur

Damit das Werkzeug funktioniert, müssen die folgenden Dateien vorbereitet und idealerweise im selben Ordner wie das Skript abgelegt werden:

1.  **`quittungsgenerator.py`**

      - Die Haupt-Skriptdatei, die das Programm enthält.

2.  **`schuelerliste.xlsx`**

      - Die Liste aller Schüler und ihrer zugehörigen Elternteile.
      - **Benötigte Spalten:**
          - `Familienname`: Nachname der Familie.
          - `Vorname_Elternteil`: Vorname des Elternteils.
          - `Vorname_Kind`: Vorname des Kindes.
      - *Hinweis: Kinder werden anhand der Kombination aus `Familienname` und `Vorname_Elternteil` gruppiert.*

3.  **`preise.xlsx`**

      - Die Konfigurationsdatei für alle Beträge und das Schuljahr.
      - **Benötigte Tabellenblätter (`Sheets`):**
          - **`Gebuehren`**: Enthält die gestaffelten Preise pro Kind.
              - Spalte A: `Kind_Nr` (z.B. 1, 2, 3)
              - Spalte B: `Betrag` (z.B. 360, 220, 170)
          - **`Beitraege`**: Enthält fixe Beiträge pro Familie.
              - Spalte A: `Posten` (z.B. "Mitgliedsbeitrag")
              - Spalte B: `Betrag` (z.B. 35)
          - **`Konfiguration`**: Enthält allgemeine Einstellungen.
              - Spalte A: `Eigenschaft` (z.B. "Schuljahr")
              - Spalte B: `Wert` (z.B. "2024/2025")

4.  **`Quittung-Template.docx`**

      - Die Word-Vorlage für die Quittung.
      - **Benötigte Platzhalter:**
          - `{{ELTERN_NAME}}`, `{{KINDER_NAMEN}}`, `{{NR}}`, `{{DATUM}}`, `{{SCHULJAHR}}`
          - `{{BETRAG_GEBUEHR}}`, `{{BETRAG_MITGLIED}}`, `{{GESAMTBETRAG}}`
          - `{{BETRAG_GEBUEHR_WORT}}`, `{{BETRAG_MITGLIED_WORT}}`, `{{GESAMTBETRAG_WORT}}`

## 🚀 Anwendung

### Für Endanwender (Nutzung der `.exe`-Datei)

1.  Stellen Sie sicher, dass die `.exe`-Datei im selben Ordner wie die drei benötigten Dateien (`schuelerliste.xlsx`, `preise.xlsx`, `Quittung-Template.docx`) liegt.
2.  Doppelklicken Sie auf die `.exe`-Datei, um das Programm zu starten.
3.  Die Pfade zu den Dateien sollten automatisch ausgefüllt werden. Falls nicht, wählen Sie die Dateien manuell über die "Durchsuchen..."-Knöpfe aus.
4.  Wählen Sie einen **Ausgabeordner**, in den die fertigen Quittungen gespeichert werden sollen.
5.  Klicken Sie auf den Knopf **"🚀 Quittungen generieren"**.
6.  Nach Abschluss des Vorgangs erscheint eine Erfolgsmeldung. Sie finden die generierten `.docx`-Dateien im gewählten Ausgabeordner.

### Für Entwickler (Ausführen des Skripts)

1.  Stellen Sie sicher, dass Sie alle Anforderungen aus dem Abschnitt "Installation" erfüllt haben.
2.  Öffnen Sie eine Kommandozeile oder ein Terminal im Projektverzeichnis.
3.  Führen Sie das Skript aus:
    ```bash
    python quittungsgenerator.py
    ```
4.  Folgen Sie den Schritten 3-6 aus der Anleitung für Endanwender.

## 📦 Erstellen einer Executable (`.exe`)

Um eine eigenständige `.exe`-Datei für Windows zu erstellen, die ohne Python-Installation läuft, können Sie **PyInstaller** verwenden.

1.  **PyInstaller installieren:**

    ```bash
    pip install pyinstaller
    ```

2.  **Executable erstellen:**
    Öffnen Sie eine Kommandozeile im Projektordner (wo alle Ihre Dateien liegen) und führen Sie den folgenden Befehl aus:

    ```bash
    pyinstaller --onefile --windowed --add-data "Quittung-Template.docx;." --add-data "preise.xlsx;." --add-data "schuelerliste.xlsx;." quittungsgenerator.py
    ```

      - `--onefile`: Bündelt alles in eine einzige `.exe`-Datei.
      - `--windowed`: Versteckt das schwarze Konsolenfenster beim Start.
      - `--add-data "Dateiname;."`: Packt die notwendigen Datenfiles mit in die `.exe`-Datei.

3.  **Ergebnis finden:**
    Ihre fertige `quittungsgenerator.exe`-Datei finden Sie im neu erstellten `dist`-Ordner. Diese Datei können Sie weitergeben.


---

&copy; 2025 GitHub &bull; [Code of Conduct](https://www.contributor-covenant.org/version/2/1/code_of_conduct/code_of_conduct.md) &bull; [MIT License](https://gh.io/mit)