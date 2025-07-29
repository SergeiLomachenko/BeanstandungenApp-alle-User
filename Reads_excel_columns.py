import pandas as pd
import argparse 
import os
import numpy as np

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('month', nargs='?', default='January', help='Monat für Filterung')  # Необязательный аргумент
    parser.add_argument('--recl', default='recl.xlsx', help='Pfad zur recl.xlsx Datei')
    parser.add_argument('--grp', default='grp.xlsx', help='Pfad zur grp.xlsx Datei')
    return parser.parse_args()

# Parcer ergänzt
args = parse_args()

print(f"Verarbeitung für Monat: {args.month}")
print(f"Recl-Datei: {args.recl}")

# Überprüfen ob die Datei exists
if not os.path.exists(args.recl):
    print(f"Fehler: Datei {args.recl} nicht gefunden!")
    exit(1)

# 1. Rohdaten komplett einlesen (ohne Header)
df_raw = pd.read_excel(args.recl, header=None, engine='openpyxl')
# df_grp = pd.read_excel(args.grp, header=None, engine='openpyxl') 
df_raw.to_excel('file2_raw.xlsx', header=False, index=False)
print("Rohdaten: file2_raw.xlsx")

# 2. Erste 3 Zeilen entfernen
df_no_rows = df_raw.iloc[3:].reset_index(drop=True)

# 3. Spalten entfernen:
#    - die ersten 4 Spalten (0–3)
#    - zusätzlich Spalten 11, 14, 15, 16, 17, 21, 22, 23, 24, 25
#    - sowie Spalten 27 bis 41 (inklusive)
cols_to_drop = [0, 1, 2, 3,
                11, 14, 15, 16, 17,
                21, 22, 23, 24, 25] + list(range(27, 41))

df_processed = df_no_rows.drop(columns=cols_to_drop, errors='ignore')
df_processed.to_excel('file2_processed.xlsx', header=False, index=False)
print("Ohne ausgewählte Spalten: file2_processed.xlsx")

# 4. Filtern:
#    – Die erste Zeile (Header) unberührt lassen
header_row = df_processed.iloc[[0]]
data_rows  = df_processed.iloc[1:]

#    – Nur Zeilen, in denen Spalte 18 == "Einsteller"
#      UND Spalte 7 mit "AG" beginnt
mask = (
    (data_rows[18] == 'Einsteller') &
    (data_rows[7].astype(str).str.startswith('AG')) &
    (data_rows[19] != 'Wurde abgelehnt'))

filtered_rows = data_rows[mask]

# Monatfilterung ergänzt
if args.month and len(filtered_rows) > 0:
    # Monatszuordnung
    month_map = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12
    }

    try:
        print(f"\n=== Monatsfilterung für {args.month} ===")
        
        # Nach dem Löschen der Spalten ist die Datumsspalte jetzt Spalte 0
        # (weil Spalten 0,1,2,3 gelöscht wurden)

        date_column_position = 0  # Erste Spalte (positionsmäßig)
        print(f"Verwende Spalte an Position {date_column_position} für Datumsfilterung")

        print(f"Erste 5 Werte in Datumsspalte:")
        print(filtered_rows.iloc[:5, date_column_position] if len(filtered_rows) > 0 else "Keine Daten")
        
        # Datumskonvertierung mit dem richtigen Format DD/MM/YYYY
        date_series = pd.to_datetime(
            filtered_rows.iloc[:, 0],  
            errors='coerce'
        )
        
        # Wenn nicht alle Daten erkannt wurden, versuche spezifische Formate
        if date_series.notna().sum() < len(filtered_rows) * 0.5:  # Wenn weniger als 50% erkannt
            for fmt in date_formats:
                try:
                    date_series = pd.to_datetime(filtered_rows[0], format=fmt, errors='coerce')
                    if date_series.notna().sum() > 0:
                        print(f"Verwendetes Datumsformat: {fmt}")
                        break
                except:
                    continue
        
        # Debug-Informationen
        valid_dates = date_series.notna().sum()
        print(f"Erfolgreich konvertierte Datumsangaben: {valid_dates} von {len(date_series)}")
        
        month_number = month_map.get(args.month, 1)
        print(f"Filtern nach Monat: {args.month} (Nummer {month_number})")

        if valid_dates > 0:
            available_months = sorted(date_series.dt.month.dropna().unique())
            print(f"Verfügbare Monate in den Daten: {available_months}")
            

            # Prüfe ob der gewünschte Monat in den Daten vorhanden ist
            if month_number not in available_months:
                print(f"ACHTUNG: Monat {args.month} ({month_number}) ist in den Daten nicht vorhanden!")
            
            # Monatsfilterung anwenden
            month_mask = date_series.dt.month == month_number
            rows_before = len(filtered_rows)
            filtered_rows = filtered_rows[month_mask]
            rows_after = len(filtered_rows)
            
            print(f"Zeilen vor Filter: {rows_before}")
            print(f"Zeilen nach Filter: {rows_after}")
            
            print(f"Gefiltert nach Monat: {args.month}")
        else:
            print("Keine Datumsangaben konnten konvertiert werden!")
            
    except Exception as e:
        print(f"Fehler bei der Monatsfilterung: {e}")
        import traceback
        traceback.print_exc()

#    Header und gefilterte Daten wieder zusammenfügen
df_final = pd.concat([header_row, filtered_rows], ignore_index=True)

# 5. Timestamp-Spalten 0 und 6 in reines Datum wandeln (ab Zeile 1)
for col_idx in (0, 6):
    df_final.iloc[1:, col_idx] = pd.to_datetime(
        df_final.iloc[1:, col_idx],
        format="%d.%m.%Y %H:%M:%S", 
        # infer_datetime_format=True,
        errors='coerce'
    ).dt.date

# 6. Ergebnis speichern
df_final.to_excel('file2_filtered.xlsx', header=False, index=False)
print("Gefilterte Datei ist in der file2_filtered.xlsx Datei")