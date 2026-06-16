# Quittungs-Generator für Schulgebühren

Ein Desktop-Werkzeug mit grafischer Benutzeroberfläche (GUI) zur automatisierten Erstellung von personalisierten Quittungen. Das Tool liest Schüler- und Preisdaten aus Excel-Dateien, fügt diese in eine Word-Vorlage ein und generiert für jede Familie eine separate Quittungsdatei, die übersichtlich in Ordnern nach Klassen sortiert wird.

## ✨ Features

- **Grafische Benutzeroberfläche (GUI):** Einfache Bedienung über ein Fenster.
- **Dynamische Preisberechnung:** Berechnet die Schulgebühren automatisch anhand der Anzahl der Kinder pro Familie.
- **Externe Konfiguration:** Alle wichtigen Parameter (Preise, Mitgliedsbeiträge, Schuljahr) werden flexibel in einer separaten Excel-Datei verwaltet.
- **Automatisierte Dokumentenerstellung:** Generiert `.docx`-Dateien basierend auf einer anpassbaren Word-Vorlage.
- **Formatierungserhalt:** Behält die Formatierungen (Fettdruck, Schriftarten etc.) aus der Word-Vorlage bei.
- **Zahl-zu-Wort-Umwandlung:** Rechnet Zahlbeträge automatisch in ausgeschriebenen Text um.
- **Automatische Dateierkennung:** Sucht beim Start nach den Standard-Dateinamen im Programmverzeichnis und füllt die Pfade automatisch aus.
- **Organisierte Ausgabe:** Erstellt automatisch einen `out`-Ordner und darin Unterordner für jede Klasse.
- **Robuste Fehlerbehandlung:** Bricht bei fehlerhaften Daten in der Excel-Datei nicht ab, sondern überspringt diese und meldet alle Probleme am Ende gesammelt.
- **Automatisches Filtern:** Einträge mit dem Status `Abgemeldet` oder `Warteliste` werden ignoriert.

---

## ⚙️ Installation (für Entwickler)

1.  **Repository klonen:**
    ```bash
    git clone [https://github.com/zlatanchev/bg-school-certificate-generator.git](https://github.com/zlatanchev/bg-school-certificate-generator.git)
    cd bg-school-certificate-generator
    ```

2.  **Virtuelle Umgebung erstellen und aktivieren (empfohlen):**
    ```bash
    python -m venv venv
    # Windows
    .\venv\Scripts\activate
    ```

3.  **Abhängigkeiten installieren:**
    ```bash
    pip install pandas python-docx num2words openpyxl
    ```

---

## 📂 Benötigte Dateien & Struktur

Damit das Werkzeug funktioniert, müssen die folgenden Dateien vorbereitet und im selben Ordner wie das Skript abgelegt werden:

1.  **`quittungs_generator.py`**
    - Die Haupt-Skriptdatei.

2.  **`schuelerliste.xlsx`**
    - Die Liste aller Schüler und ihrer zugehörigen Mitglieder.
    - **Benötigte Spalten:**
        - `Eltern 1 - Name`: Der vollständige Name des zahlenden Mitglieds/Elternteils. **Die Gruppierung erfolgt nach dieser Spalte        
        - `Eltern 1 - Emailadresse`: Die E-Mail des Mitglieds/Elternteils. **Die Gruppierung erfolgt nach dieser Spalte
        - `Name Kind`:  Der vollständige Name des Kindes.
        - `In Klasse`: Die Bezeichnung der Klasse des Kindes