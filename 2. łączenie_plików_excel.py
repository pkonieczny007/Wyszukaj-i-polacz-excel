import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

def przeszukaj_i_utworz_df(folder):
    # Sprawdź czy podana lokalizacja istnieje
    if not os.path.exists(folder):
        print("Podana lokalizacja nie istnieje.")
        return None

    # Lista przechowująca DataFrame'y z różnych plików
    dfs = []

    # Przeszukaj folder i wszystkie podfoldery w poszukiwaniu plików Excel
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xlsm'):
                if file.startswith("Produktionsstückliste"):
                    # Wczytaj wybrane kolumny z pliku Excel do DataFrame
                    filepath = os.path.join(root, file)
                    df = pd.read_excel(filepath, usecols=["Stufe", "Bezeichnung.............................", "Zeinr", "ZEF", "ZEV","Gewicht"])
                    # Dodaj kolumnę z nazwą pliku
                    df['Plik'] = file
                    dfs.append(df)

    if not dfs:
        print("Nie znaleziono plików spełniających warunki.")
        return None

    # Połącz wszystkie DataFrame'y w jeden
    result_df = pd.concat(dfs, ignore_index=True)
    return result_df

def zapisz_do_xlsx(df, output_file):
    if df is None:
        print("Nie można zapisać danych, ponieważ DataFrame jest pusty.")
        return

    # Utwórz nowy arkusz w pliku Excel
    wb = Workbook()
    ws = wb.active

    # Zapisz nagłówki kolumn
    ws.append(df.columns.tolist() + ['Plik'])

    # Zapisz dane z DataFrame do arkusza Excel
    for r in dataframe_to_rows(df):
        ws.append(r)

    # Zapisz plik Excel
    wb.save(output_file)
    print(f"Dane zapisano do pliku: {output_file}")

def dataframe_to_rows(df):
    # Funkcja pomocnicza do konwersji DataFrame na wiersze dla arkusza Excel
    for row in df.itertuples(index=False):
        yield row

def wybierz_folder():
    folder_selected = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(0, folder_selected)

def przeszukaj_i_zapisz():
    folder_to_search = entry_folder.get()
    df = przeszukaj_i_utworz_df(folder_to_search)
    if df is not None:
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Pliki Excel", "*.xlsx")])
        if output_file:
            zapisz_do_xlsx(df, output_file)

# Utwórz okno główne
root = tk.Tk()
root.title("Przeszukiwanie i zapis danych do pliku Excel")

# Pole tekstowe do wyświetlania wybranej lokalizacji folderu
entry_folder = tk.Entry(root, width=50)
entry_folder.grid(row=0, column=0, padx=10, pady=5, columnspan=2)

# Przycisk do wyboru folderu
btn_browse = tk.Button(root, text="Wybierz folder", command=wybierz_folder)
btn_browse.grid(row=1, column=0, padx=10, pady=5)

# Przycisk rozpoczynający przeszukiwanie i zapis danych
btn_search = tk.Button(root, text="Przeszukaj i zapisz do Excel", command=przeszukaj_i_zapisz)
btn_search.grid(row=1, column=1, padx=10, pady=5)

# Uruchom pętlę główną
root.mainloop()
