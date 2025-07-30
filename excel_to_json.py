import pandas as pd
import json
import re

def excel_to_calendar_json(excel_file_path):
    """
    Konvertiert eine Excel-Tabelle in eine JSON-Datei, indem sie die
    Datumszeile findet und die Events sowie Minuten in den Zeilen darunter extrahiert.
    """
    
    events = []
    
    try:
        excel_data = pd.read_excel(excel_file_path, sheet_name=None, header=None, dtype=str, keep_default_na=False)
        print("Excel-Datei erfolgreich geladen. Suche auf allen Seiten...")
    except FileNotFoundError:
        print(f"FEHLER: Datei '{excel_file_path}' nicht gefunden.")
        return
    
    for sheet_name, df in excel_data.items():
        print(f"Prüfe Seite: '{sheet_name}'...")
        
        for row_index, row in df.iterrows():
            
            date_matches_in_row = 0
            for cell_value in row:
                if re.search(r'\d{4}-\d{2}-\d{2}', str(cell_value)):
                    date_matches_in_row += 1
            
            if date_matches_in_row >= 3:
                date_row_index = row_index
                print(f"Datumszeile gefunden in Zeile: {date_row_index}")
                
                for col_index in range(len(df.columns)):
                    cell_value = df.iloc[date_row_index, col_index]
                    date_iso_string = None

                    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', str(cell_value))
                    if date_match:
                        date_iso_string = date_match.group(1)
                    
                    if date_iso_string:
                        day_minutes = 0
                        
                        for event_offset in range(1, 5): 
                            event_row_index = date_row_index + event_offset
                            
                            if event_row_index >= len(df):
                                break
                                
                            event_title_cell = df.iloc[event_row_index, col_index]
                            
                            if str(event_title_cell).strip() != '':
                                event_dict = {
                                    "title": str(event_title_cell),
                                    "start": date_iso_string,
                                    "allDay": True
                                }

                                # Suche nach Minuten in der nächsten Spalte
                                if col_index + 1 < len(df.columns):
                                    minutes_cell = df.iloc[event_row_index, col_index + 1]
                                    if str(minutes_cell).strip().isdigit():
                                        minutes = int(str(minutes_cell).strip())
                                        event_dict["minutes"] = minutes
                                        day_minutes += minutes

                                events.append(event_dict)
                                
                # Wichtig: Die break-Anweisung wurde entfernt, damit das Skript
                # alle weiteren Datumszeilen findet.
    
    output_json_file = 'kalender_events.json'
    with open(output_json_file, 'w', encoding='utf-8') as f:
        json.dump(events, f, indent=4, ensure_ascii=False)
        
    print(f"Erfolgreich {len(events)} Events in '{output_json_file}' gespeichert.")


if __name__ == "__main__":
    excel_file_path = "kalender_training.xlsx" 
    excel_to_calendar_json(excel_file_path)