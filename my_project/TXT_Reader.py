import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# TXT mapping (0-based)
COLUMN_INDEXES = [0, 3, 2, 4]  # Kod, ProduktNazwa, Cena, VAT
TARGET_COLUMNS = ["Kod", "ProduktNazwa", "Cena", "VAT"]

class TXTtoXLSXApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TXT → XLSX / CSV Merger")

        self.txt_files = []
        self.target_file = None

        tk.Label(root, text="Wybierz pliki TXT").pack(pady=5)
        tk.Button(root, text="Wybierz TXT", command=self.load_txt).pack()

        # Listbox + scrollbar
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

        tk.Button(root, text="Kopiuj dane", command=self.merge).pack(pady=10)
        root.geometry("500x400")

    def load_txt(self):
        files = filedialog.askopenfilenames(filetypes=[("TXT files", "*.txt")])
        for file in files:
            if file not in self.txt_files:
                self.txt_files.append(file)
                self.update_listbox()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, file in enumerate(self.txt_files, start=1):
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
        if not self.txt_files or not self.target_file:
            messagebox.showerror("Błąd", "Wybierz pliki TXT i plik docelowy")
            return

        # wczytaj lub utwórz dane wynikowe
        if os.path.exists(self.target_file):
            target_df = pd.read_excel(self.target_file)
        else:
            target_df = pd.DataFrame(columns=TARGET_COLUMNS)

        for file in self.txt_files:
            df_list = []
            try:
                with open(file, "r", encoding="utf-8", errors="replace") as f:
                    for line in f:
                        line = line.rstrip("\n")
                        parts = line.split(";")
                        # uzupełniamy brakujące kolumny pustym stringiem
                        while len(parts) < max(COLUMN_INDEXES) + 1:
                            parts.append("")
                        # wybieramy tylko potrzebne kolumny
                        selected = [parts[i] for i in COLUMN_INDEXES]
                        df_list.append(selected)
                df = pd.DataFrame(df_list, columns=TARGET_COLUMNS)
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się wczytać pliku {os.path.basename(file)}: {e}")
                return

            target_df = pd.concat([target_df, df], ignore_index=True)

        # zapis XLSX
        target_df.to_excel(self.target_file, index=False)

        # zapis CSV
        csv_output = os.path.splitext(self.target_file)[0] + ".csv"
        target_df.to_csv(csv_output, sep=";", index=False, encoding="utf-8")

        messagebox.showinfo("OK", "Dane zostały zapisane do XLSX i CSV")

        self.txt_files.clear()
        self.update_listbox()

if __name__ == "__main__":
    root = tk.Tk()
    app = TXTtoXLSXApp(root)
    root.mainloop()
