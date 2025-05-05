import pandas as pd
import os
from openpyxl.utils import get_column_letter

def compare_excel_files(subdir, output_filename):
    base_dir = os.path.dirname(os.path.abspath(__file__))  # Basisverzeichnis des Projekts
    subdir_path = os.path.join(base_dir, subdir)
    output_file = os.path.join(base_dir, output_filename)
    
    # Alle Excel-Dateien im Verzeichnis einlesen
    excel_files = [f for f in os.listdir(subdir_path) if f.endswith(".xlsx")]
    
    if len(excel_files) != 2:
        raise ValueError("Es müssen genau zwei Excel-Dateien im Verzeichnis vorhanden sein!")
    
    file1_path = os.path.join(subdir_path, excel_files[0])
    file2_path = os.path.join(subdir_path, excel_files[1])
    
    # Nutzer gibt auszuschließende Zeilen und Spalten ein
    exclude_rows = input("Welche Zeilen sollen ausgeschlossen werden? (z.B. 1,3,5): ")
    exclude_cols = input("Welche Spalten sollen ausgeschlossen werden? (z.B. A,C,E): ")
    
    exclude_rows = set(map(int, exclude_rows.split(','))) if exclude_rows else set()
    exclude_cols = set(ord(c.upper()) - 65 for c in exclude_cols.split(',')) if exclude_cols else set()
    
    # Beide Excel-Dateien einlesen
    df1 = pd.read_excel(file1_path, sheet_name=None, engine='openpyxl', header=None)
    df2 = pd.read_excel(file2_path, sheet_name=None, engine='openpyxl', header=None)
    
    with open(output_file, 'w') as f:
        for sheet in df1.keys():  # Annahme: Beide Dateien haben dieselben Sheets
            if sheet not in df2:
                f.write(f"Sheet {sheet} existiert nicht in {excel_files[1]}\n")
                continue
            
            sheet1 = df1[sheet].fillna('')  # NaN-Werte durch leeren String ersetzen
            sheet2 = df2[sheet].fillna('')
            
            max_rows = max(sheet1.shape[0], sheet2.shape[0])
            max_cols = max(sheet1.shape[1], sheet2.shape[1])
            
            for row in range(max_rows):
                if row + 1 in exclude_rows:
                    continue
                
                for col in range(max_cols):
                    if col in exclude_cols:
                        continue
                    
                    value1 = sheet1.iloc[row, col] if row < sheet1.shape[0] and col < sheet1.shape[1] else ''
                    value2 = sheet2.iloc[row, col] if row < sheet2.shape[0] and col < sheet2.shape[1] else ''
                    
                    if value1 != value2:
                        cell = f"{get_column_letter(col + 1)}{row + 1}"
                        f.write(f"{cell}: difference = {value1} -> {value2}\n")
    
    print(f"Vergleich abgeschlossen. Unterschiede gespeichert in {output_file}")

compare_excel_files("tobediffchecked", "differenzen.txt")
