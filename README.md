# Running Dinner Calculator V2

Dies ist die aktualisierte Python-Version des Running Dinner Calculators. Das Skript automatisiert die Zuteilung von Teams zu Gängen und erstellt Textvorlagen für die Benachrichtigungs-E-Mails.

## Voraussetzungen

- **Python** muss installiert sein.
- Ein **OpenRouteService API Key** (kostenlos erstellbar unter [openrouteservice.org](https://openrouteservice.org/)).
- Die Dateien `data.xlsx` und `Afterparty.xlsx` müssen im selben Ordner liegen und korrekt formatiert sein.

## Installation

1. Öffnen Sie die Kommandozeile (Terminal) in diesem Ordner.
2. Installieren Sie die benötigten Python-Bibliotheken mit folgendem Befehl:
   ```bash
   pip install -r requirements.txt
   ```

## Konfiguration

1. Öffnen Sie die Datei `running_dinner_calculator.py` in einem Texteditor oder einer Python-IDE.
2. Suchen Sie die Zeile `ORS_API_KEY = """secretaccesskey="""` und fügen Sie Ihren persönlichen API-Key zwischen die Anführungszeichen ein.
3. Stellen Sie sicher, dass Ihre Excel-Dateien korrekt sind:
   - **data.xlsx**: Benötigte Spalten:
     - `Team Nr.` (ID)
     - `Name 1`, `Name 2`
     - `Adress` (Standort)
     - `will to double` (Zahl, Bereitschaft doppelt zu kochen)
     - `readiness starter` (Wahr/Falsch, Bereitschaft für Vorspeise)
     - `Ring at` (Klingelschild)
     - `Allergies or else` (Allergien/Infos)
   - **Afterparty.xlsx**: Benötigte Spalte:
     - `Adress` (Adresse der Afterparty)

## Ausführung

Starten Sie das Skript über die Kommandozeile:

```bash
python running_dinner_calculator.py
```

Folgen Sie den Anweisungen im Terminal, um Zeiten und Organisator-Infos einzugeben.

## Output

Das Skript generiert:
- **`overview.txt`**: Eine Übersicht, wer welchen Gang ausrichtet und wen diese Teams bewirten.
- **`Mail Team Nr. [X].txt`**: Für jedes Team eine individuelle Textdatei mit dem Entwurf für die Informations-E-Mail.
