import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Mapowanie nazw kolumn z CSV na wynikowy XLSX/CSV
COLUMN_MAPPING = {
    "Kod": "Kod",
    "ProduktNazwa": "ProduktNazwa",
    "Cena": "Cena",
    "VAT": "VAT"
}
TARGET_COLUMNS = ["Kod", "ProduktNazwa", "Cena", "VAT"]

class CSVtoXLSXApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV → XLSX / CSV Merger + Sort")

        self.csv_files = []
        self.target_file = None

        tk.Label(root, text="Wybierz pliki CSV").pack(pady=5)
        tk.Button(root, text="Wybierz CSV", command=self.load_csv).pack()

        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        self.scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(frame, height=8, yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Button(root, text="Wybierz plik docelowy XLSX", command=self.select_target).pack(pady=5)
        self.target_label = tk.Label(root, text="Plik docelowy: brak")
        self.target_label.pack()

        tk.Button(root, text="Kopiuj dane + Sort", command=self.merge).pack(pady=10)

        root.geometry("500x400")

    def load_csv(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        for file in files:
            if file not in self.csv_files:
                self.csv_files.append(file)
                self.update_listbox()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, file in enumerate(self.csv_files, start=1):
            display_name = f"{idx}. {os.path.basename(file)}"
            self.listbox.insert(tk.END, display_name)

    def select_target(self):
        self.target_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.target_file:
            self.target_label.config(text=os.path.basename(self.target_file))

    def merge(self):
        if not self.csv_files or not self.target_file:
            messagebox.showerror("Błąd", "Wybierz pliki CSV i plik docelowy")
            return

        # wczytaj lub utwórz dane wynikowe
        if os.path.exists(self.target_file):
            target_df = pd.read_excel(self.target_file)
        else:
            target_df = pd.DataFrame(columns=TARGET_COLUMNS)

        for file in self.csv_files:
            try:
                df = pd.read_csv(
                    file,
                    sep=";",
                    header=0,
                    encoding="cp1250",
                    dtype=str,
                    keep_default_na=False  # puste pola jako ""
                )
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się wczytać pliku {os.path.basename(file)}: {e}")
                return

            # sprawdzamy kolumny i wybieramy tylko potrzebne
            missing_cols = [col for col in COLUMN_MAPPING.keys() if col not in df.columns]
            if missing_cols:
                messagebox.showerror(
                    "Błąd",
                    f"Plik {os.path.basename(file)} nie zawiera wymaganych kolumn: {missing_cols}"
                )
                return

            selected = df[list(COLUMN_MAPPING.keys())].rename(columns=COLUMN_MAPPING)
            target_df = pd.concat([target_df, selected], ignore_index=True)

        # sortowanie od ostatniego wpisu i zachowanie unikalnych Kodów
        target_df = target_df.iloc[::-1].reset_index(drop=True)
        target_df = target_df.drop_duplicates(subset="Kod", keep="first")

        # zapis XLSX
        target_df.to_excel(self.target_file, index=False)

        # zapis CSV
        csv_output = os.path.splitext(self.target_file)[0] + ".csv"
        target_df.to_csv(csv_output, sep=";", index=False, encoding="cp1250")

        messagebox.showinfo("OK", "Dane zostały zapisane do XLSX i CSV")

        self.csv_files.clear()
        self.update_listbox()


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVtoXLSXApp(root)
    root.mainloop()
