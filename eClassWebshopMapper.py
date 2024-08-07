import os
import json
import pandas as pd
import numpy as np
import openai
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

# OpenAI API-Schlüssel setzen
openai.api_key = 'YOUR_API_KEY'

# Farbcodes für die Konsole
class ConsoleColor:
    RED = '\033[91m'
    GREEN = '\033[92m'
    END = '\033[0m'

# Funktion zum Einlesen der Dateien
def read_files(excel_file_path):
    try:
        print('Dateien werden eingelesen...')
        base_path = os.path.join(os.path.dirname(__file__), 'Data')
        
        with open(os.path.join(base_path, 'Webshop-Kategorien.json'), 'r') as f:
            webshop_data = json.load(f)
            webshop_categories = webshop_data["Webshop Kategorien"]

        with open(os.path.join(base_path, 'EClass7.1.json'), 'r') as f:
            eclass_data = json.load(f)
            eclasses = eclass_data["eClass Kategorien"]

        eclass_df = pd.read_excel(excel_file_path)
        print(f"{ConsoleColor.GREEN}Dateien erfolgreich eingelesen.{ConsoleColor.END}")
        return webshop_categories, eclasses, eclass_df
    except PermissionError as e:
        print(f"{ConsoleColor.RED}PermissionError: {e}{ConsoleColor.END}")
    except Exception as e:
        print(f"{ConsoleColor.RED}Fehler beim Einlesen der Dateien: {e}{ConsoleColor.END}")

# Funktion zur Berechnung der Embeddings
def get_embedding(text):
    response = openai.embeddings.create(input=text, model="text-embedding-3-large")
    return response.data[0].embedding

# Funktion zum Erstellen der Embeddings
def create_embeddings(eclasses, webshop_categories):
    try:
        print('Erstelle eClass und Webshop Embeddings...')
        
        base_path = os.path.join(os.path.dirname(__file__), 'Data')
        eclass_embeddings_file = os.path.join(base_path, 'eclass_embeddings.json')
        webshop_embeddings_file = os.path.join(base_path, 'webshop_embeddings.json')

        # eClass Embeddings erstellen
        if os.path.exists(eclass_embeddings_file):
            print('eClass Embeddings aus Datei werden geladen...')
            with open(eclass_embeddings_file, 'r') as f:
                eclass_embeddings = json.load(f)
            print(f"{ConsoleColor.GREEN}eClass Embeddings aus Datei erfolgreich geladen.{ConsoleColor.END}")
        else:
            print('eClass Embeddings werden erstellt...')
            eclass_embeddings = {}
            with ThreadPoolExecutor() as executor:
                futures = {executor.submit(get_embedding, get_full_hierarchy_description(eclass, eclasses)): eclass for eclass in eclasses}
                for future in as_completed(futures):
                    eclass = futures[future]
                    try:
                        eclass_embeddings[eclass['eclass_nummer']] = future.result()
                    except Exception as exc:
                        print(f"{ConsoleColor.RED}eClass Embedding für {eclass['eclass_nummer']} erzeugte einen Fehler: {exc}{ConsoleColor.END}")

            with open(eclass_embeddings_file, 'w') as f:
                json.dump(eclass_embeddings, f)
            print(f"{ConsoleColor.GREEN}eClass Embeddings erfolgreich erstellt.{ConsoleColor.END}")

        # Webshop Embeddings erstellen
        if os.path.exists(webshop_embeddings_file):
            print('Webshop Embeddings aus Datei werden geladen...')
            with open(webshop_embeddings_file, 'r') as f:
                webshop_embeddings = json.load(f)
            print(f"{ConsoleColor.GREEN}Webshop Embeddings aus Datei erfolgreich geladen.{ConsoleColor.END}")
        else:
            print('Webshop Embeddings werden erstellt...')
            webshop_embeddings = {}
            hierarchical_descriptions = create_hierarchical_descriptions(webshop_categories)
            with ThreadPoolExecutor() as executor:
                futures = {executor.submit(get_embedding, description): (category_id, description) for category_id, description in hierarchical_descriptions}
                for future in as_completed(futures):
                    category_id, description = futures[future]
                    try:
                        webshop_embeddings[category_id] = future.result()
                    except Exception as exc:
                        print(f"{ConsoleColor.RED}Webshop Embedding für {category_id} erzeugte einen Fehler: {exc}{ConsoleColor.END}")

            with open(webshop_embeddings_file, 'w') as f:
                json.dump(webshop_embeddings, f)
            print(f"{ConsoleColor.GREEN}Webshop Embeddings erfolgreich erstellt.{ConsoleColor.END}")

        return eclass_embeddings, webshop_embeddings
    except PermissionError as e:
        print(f"{ConsoleColor.RED}PermissionError: {e}{ConsoleColor.END}")
    except Exception as e:
        print(f"{ConsoleColor.RED}Fehler beim Erstellen der Embeddings: {e}{ConsoleColor.END}")

# Funktion zum Erstellen der eclass Hierarchie
def get_full_hierarchy_description(category, mapping):
    levels = [category['eclass_nummer'][i:i+2] for i in range(0, len(category['eclass_nummer']), 2)]
    name_hierarchy = []
    for i in range(1, len(levels) + 1):
        key = ''.join(levels[:i]) + '0' * (8 - 2 * i)
        matching_category = next((item for item in mapping if item["eclass_nummer"] == key), None)
        if matching_category:
            name_hierarchy.append(matching_category['name'])
    name_hierarchy.append(category['name'])
    return " > ".join(name_hierarchy)

# Funktion zum Erstellen der Webshop Kategorie Hierarchie
def create_hierarchical_descriptions(categories):
    hierarchical_descriptions = []
    for category in categories:
        parts = category['fullpath'].split('/')
        hierarchy = " > ".join(parts[2:])  # Erzeugt die volle Hierarchie ohne "/shop/"
        hierarchical_descriptions.append((category["erpCategoryId"], hierarchy, "/".join(parts[2:])))
    return hierarchical_descriptions

# Funktion zur Berechnung der Ähnlichkeit zwischen zwei Vektoren
def cosine_similarity(vec1, vec2):
    return np.dot(vec1, vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))

# Hauptfunktion zur Durchführung der Zuordnung
def main():
    try:
        if len(sys.argv) < 2:
            print(f"{ConsoleColor.RED}Fehler: Die Excel-Datei muss als Argument übergeben werden: >> python 'Skriptname' <tabelle.xlsx>{ConsoleColor.END}")
            sys.exit(1)

        excel_file_path = sys.argv[1]

        print('Skript startet...')
        webshop_categories, eclasses, eclass_df = read_files(excel_file_path)
        if webshop_categories is None or eclasses is None or eclass_df is None:
            print(f"{ConsoleColor.RED}Skript abgebrochen: Fehler beim Einlesen der Dateien.{ConsoleColor.END}")
            return

        eclass_embeddings, webshop_embeddings = create_embeddings(eclasses, webshop_categories)
        if eclass_embeddings is None or webshop_embeddings is None:
            print(f"{ConsoleColor.RED}Skript abgebrochen: Fehler beim Erstellen der Embeddings.{ConsoleColor.END}")
            return

        # Entfernen Sie führende und nachfolgende Leerzeichen und wandeln Sie die eclass_nummern in Zeichenketten um
        eclass_nummern = [eclass['eclass_nummer'].strip() for eclass in eclasses]
        eclass_df['eclass-nummer'] = eclass_df['eclass-nummer'].astype(str).str.strip()

        print('Zuordnung der passenden Webshop-Kategorien beginnt...')
        # Erstellen Sie ein Dictionary für den schnellen Zugriff auf den fullpath
        webshop_fullpath_dict = {str(category['erpCategoryId']): "/".join(category['fullpath'].split('/')[2:]) for category in webshop_categories}

        # Passende Webshop Kategorien finden
        results = []
        for index, row in eclass_df.iterrows():
            eclass_nummer = row['eclass-nummer']
            if eclass_nummer in eclass_nummern:  # Prüfen, ob die eclass_nummer in eclass_nummern enthalten ist
                if eclass_nummer in eclass_embeddings:
                    eclass_embedding = eclass_embeddings[eclass_nummer]
                    similarities = {
                        category_id: cosine_similarity(eclass_embedding, webshop_embedding)
                        for category_id, webshop_embedding in webshop_embeddings.items()
                    }
                    best_match = max(similarities, key=similarities.get)
                    
                    fullpath = webshop_fullpath_dict.get(str(best_match), "N/A")
                
                    results.append({'eclass_nummer': eclass_nummer, 'best_match': best_match, 'fullpath': fullpath})
                else:
                    print(f"{ConsoleColor.RED}eClass-Nummer {eclass_nummer} nicht in Embedding enthalten.{ConsoleColor.END}")
            else:
                print(f"{ConsoleColor.RED}eClass-Nummer {eclass_nummer} nicht in den eclass_nummern enthalten.{ConsoleColor.END}")

        if results:
            print(f"{ConsoleColor.GREEN}Zuordnung abgeschlossen.{ConsoleColor.END}")
            # Ergebnisse in Excel-Datei speichern
            try:
                print('Speichere Ergebnisse in Excel-Datei...')
                results_df = pd.DataFrame(results)
                results_df.to_excel(os.path.join(os.path.dirname(__file__), 'Data', 'eclass_to_webshop_matches.xlsx'), index=False)
                print(f"{ConsoleColor.GREEN}Ergebnisse erfolgreich in 'eclass_to_webshop_matches.xlsx' gespeichert.{ConsoleColor.END}")
            except PermissionError as e:
                print(f"{ConsoleColor.RED}PermissionError beim Speichern der Datei: {e}{ConsoleColor.END}")
            except Exception as e:
                print(f"{ConsoleColor.RED}Fehler beim Speichern der Ergebnisse: {e}{ConsoleColor.END}")
        else:
            print(f"{ConsoleColor.RED}Kein Ergebnis gefunden.{ConsoleColor.END}")

        print(f"{ConsoleColor.GREEN}Skript abgeschlossen.{ConsoleColor.END}")
    except Exception as e:
        print(f"{ConsoleColor.RED}Fehler im Skript: {e}{ConsoleColor.END}")

# Ausführung des Skripts
if __name__ == "__main__":
    main()
