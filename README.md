# eClassWebshopMapper

Dieses Skript dient dazu, eClass-Kategorien mit Webshop-Kategorien zu verknüpfen. Es nutzt OpenAI-Embeddings, um die Ähnlichkeit zwischen den Beschreibungen der Kategorien zu berechnen und die beste Übereinstimmung zu finden.

## Voraussetzungen

- Python 3.x
- Die folgenden Python-Bibliotheken müssen installiert sein:
  - `pandas`
  - `numpy`
  - `openai`
  - `concurrent.futures`

Du kannst diese Bibliotheken mit dem folgenden Befehl installieren:

```bash
pip install pandas numpy openai
Dateistruktur
Stelle sicher, dass der Data-Ordner im selben Verzeichnis wie das Skript liegt. Dieser Ordner muss die folgenden Dateien enthalten:

Webshop-Kategorien.json
EClass7.1.json
Falls du die Embeddings neu erstellen möchtest, lösche einfach die folgenden Dateien aus dem Data-Ordner (falls vorhanden):

eclass_embeddings.json
webshop_embeddings.json
Diese Dateien werden beim nächsten Skriptlauf automatisch neu erstellt.

Verwendung
Um das Skript auszuführen, musst du den Pfad zu einer Excel-Datei übergeben, die die eClass-Daten enthält. Die Excel-Datei muss eine Spalte mit dem Namen eclass-nummer enthalten, welche die eClass-Nummern der zuzuordnenden Kategorien beinhaltet.

Startbefehl
Verwende den folgenden Befehl, um das Skript auszuführen:

bash
Code kopieren
python eClassWebshopMapper.py <Pfad zur Excel-Datei>
Beispiel
bash
Code kopieren
python eClassWebshopMapper.py Data/eClass-Mapping-Input.xlsx
Funktionsweise
Einlesen der Dateien: Das Skript liest die JSON-Dateien und die Excel-Datei ein.
Erstellen der Embeddings: Falls die Embeddings-Dateien nicht vorhanden sind, werden die Embeddings für die eClass- und Webshop-Kategorien erstellt und gespeichert.
Berechnung der Ähnlichkeit: Das Skript berechnet die Ähnlichkeit zwischen den eClass- und Webshop-Kategorien und findet die beste Übereinstimmung.
Speichern der Ergebnisse: Die Ergebnisse werden in einer Excel-Datei im Data-Ordner gespeichert.
Fehlerbehandlung
Falls die Excel-Datei nicht übergeben wird oder es Probleme beim Einlesen der Dateien gibt, wird eine Fehlermeldung ausgegeben und das Skript wird abgebrochen.