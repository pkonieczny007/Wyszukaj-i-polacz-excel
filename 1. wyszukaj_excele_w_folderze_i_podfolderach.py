import os
import shutil
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def znajdz_i_skopiuj_pliki(folder, dest_folder):
    # Sprawdź czy podana lokalizacja istnieje
    if not os.path.exists(folder):
        print("Podana lokalizacja nie istnieje.")
        return

    # Utwórz folder docelowy jeśli nie istnieje
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    # Przeszukaj folder i wszystkie podfoldery w poszukiwaniu plików Excel
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xlsm'):
                if file.startswith("Produktionsstückliste"):
                    # Skopiuj plik do docelowej lokalizacji
                    src_file = os.path.join(root, file)
                    dest_file = os.path.join(dest_folder, file)
                    shutil.copy(src_file, dest_file)
                    print(f"Skopiowano plik: {file}")

def wybierz_folder():
    folder_selected = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(0, folder_selected)

def przeszukaj_i_skopiuj():
    folder_to_search = entry_folder.get()
    dest_folder = "C:/DANE"
    znajdz_i_skopiuj_pliki(folder_to_search, dest_folder)

# Utwórz okno główne
root = tk.Tk()
root.title("Przeszukiwanie i kopiowanie plików Excel")

# Pole tekstowe do wyświetlania wybranej lokalizacji folderu
entry_folder = tk.Entry(root, width=50)
entry_folder.grid(row=0, column=0, padx=10, pady=5, columnspan=2)

# Przycisk do wyboru folderu
btn_browse = tk.Button(root, text="Wybierz folder", command=wybierz_folder)
btn_browse.grid(row=1, column=0, padx=10, pady=5)

# Przycisk rozpoczynający przeszukiwanie i kopiowanie plików
btn_search = tk.Button(root, text="Przeszukaj i skopiuj pliki", command=przeszukaj_i_skopiuj)
btn_search.grid(row=1, column=1, padx=10, pady=5)

# Uruchom pętlę główną
root.mainloop()
