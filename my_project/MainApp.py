import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class XLSXMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XLSM to XLSX Merger")
        self.file_list = []

        # Stałe kolumny
        self.columns_to_copy = ["Kod", "ProduktNazwa", "VAT", "CenaBrutto"]

        self.label = tk.Label(root, text="Wybierz pliki XLSM do skopiowania")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Wybierz pliki XLSM", command=self.load_files)
        self.select_button.pack(pady=5)

        self.file_listbox = tk.Listbox(root, height=6)
        self.file_listbox.pack(fill=tk.X, padx=10, pady=5)

        self.select_target_button = tk.Button(root, text="Wybierz plik docelowy XLSX", command=self.select_target_file)
        self.select_target_button.pack(pady=5)

        self.target_label = tk.Label(root, text="Plik docelowy: brak")
        self.target_label.pack(pady=5)

        self.merge_button = tk.Button(root, text="Kopiuj dane", command=self.merge_columns)
        self.merge_button.pack(pady=10)

        self.target_file = None

    def load_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Macro files", "*.xlsm")])
        if files:
            for f in files:
                if f not in self.file_list:
                    self.file_list.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))

    def select_target_file(self):
        self.target_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Wybierz plik docelowy XLSX"
        )
        if self.target_file:
            self.target_label.config(text=f"Plik docelowy: {os.path.basename(self.target_file)}")

    def validate_columns(self, df, file_path):
        missing = [col for col in self.columns_to_copy if col not in df.columns]
        if missing:
            messagebox.showerror(
                "Błąd kolumn",
                f"Plik '{os.path.basename(file_path)}' nie zawiera wymaganych kolumn:\n{', '.join(missing)}"
            )
            return False
        return True

    def merge_columns(self):
        if not self.file_list or not self.target_file:
            messagebox.showerror("Błąd", "Wybierz pliki źródłowe i docelowy")
            return

        # wczytanie lub utworzenie pliku docelowego
        if os.path.exists(self.target_file):
            target_df = pd.read_excel(self.target_file, engine="openpyxl")
        else:
            target_df = pd.DataFrame(columns=self.columns_to_copy)

        # startujemy od końca
        start_row = len(target_df)

        for file in self.file_list:
            df = pd.read_excel(file, engine="openpyxl")
            if not self.validate_columns(df, file):
                return
            df_selected = df[self.columns_to_copy]

            target_df = pd.concat([target_df, df_selected], ignore_index=True)

        target_df.to_excel(self.target_file, index=False, engine="openpyxl")
        messagebox.showinfo("Sukces", "Dane zostały skopiowane!")

        # odświeżenie listy plików
        self.file_list.clear()
        self.file_listbox.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = XLSXMergerApp(root)
    root.geometry("500x350")
    root.mainloop()
