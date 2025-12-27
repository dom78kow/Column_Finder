import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os

class XLSXMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XLSM to XLSX Column Selector")
        self.file_list = []
        self.columns = []

        self.label = tk.Label(root, text="Przeciągnij pliki XLSM tutaj lub użyj przycisku")
        self.label.pack(pady=10)

        self.drop_area = tk.Listbox(root, height=4)
        self.drop_area.pack(fill=tk.X, padx=10, pady=5)
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.drop_files)

        self.select_button = tk.Button(root, text="Wybierz pliki XLSM", command=self.load_files)
        self.select_button.pack(pady=5)

        self.column_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)
        self.column_listbox.pack(pady=5, fill=tk.BOTH, expand=True)

        self.order_frame = tk.Frame(root)
        self.order_frame.pack(pady=5)
        self.up_button = tk.Button(self.order_frame, text="↑", command=self.move_up)
        self.up_button.pack(side=tk.LEFT, padx=5)
        self.down_button = tk.Button(self.order_frame, text="↓", command=self.move_down)
        self.down_button.pack(side=tk.LEFT, padx=5)

        self.select_target_button = tk.Button(root, text="Wybierz plik docelowy XLSX", command=self.select_target_file)
        self.select_target_button.pack(pady=5)

        self.target_label = tk.Label(root, text="Plik docelowy: brak")
        self.target_label.pack(pady=5)

        self.row_frame = tk.Frame(root)
        self.row_frame.pack(pady=5)
        tk.Label(self.row_frame, text="Wiersz startowy:").pack(side=tk.LEFT)
        self.start_row_entry = tk.Entry(self.row_frame, width=5)
        self.start_row_entry.pack(side=tk.LEFT)
        self.start_row_entry.insert(0, "0")  # początkowo 0, zostanie zaktualizowane

        self.merge_button = tk.Button(root, text="Kopiuj wybrane kolumny", command=self.merge_columns)
        self.merge_button.pack(pady=10)

        self.target_file = None

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".xlsm") and f not in self.file_list:
                self.file_list.append(f)
                self.drop_area.insert(tk.END, os.path.basename(f))
        if self.file_list:
            self.load_columns(self.file_list[0])

    def load_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Macro files", "*.xlsm")])
        if files:
            for f in files:
                if f not in self.file_list:
                    self.file_list.append(f)
                    self.drop_area.insert(tk.END, os.path.basename(f))
            self.load_columns(self.file_list[0])

    def load_columns(self, file):
        df = pd.read_excel(file, engine="openpyxl")
        self.columns = list(df.columns)
        self.column_listbox.delete(0, tk.END)
        for col in self.columns:
            self.column_listbox.insert(tk.END, col)

    def move_up(self):
        selected = list(self.column_listbox.curselection())
        for i in selected:
            if i == 0:
                continue
            value = self.column_listbox.get(i)
            self.column_listbox.delete(i)
            self.column_listbox.insert(i-1, value)
            self.column_listbox.select_set(i-1)
            self.column_listbox.select_clear(i)

    def move_down(self):
        selected = list(self.column_listbox.curselection())
        for i in reversed(selected):
            if i == self.column_listbox.size()-1:
                continue
            value = self.column_listbox.get(i)
            self.column_listbox.delete(i)
            self.column_listbox.insert(i+1, value)
            self.column_listbox.select_set(i+1)
            self.column_listbox.select_clear(i)

    def select_target_file(self):
        self.target_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Wybierz plik docelowy XLSX"
        )
        if self.target_file:
            self.target_label.config(text=f"Plik docelowy: {os.path.basename(self.target_file)}")
            # automatycznie wykrywamy ostatni wiersz
            if os.path.exists(self.target_file):
                target_df = pd.read_excel(self.target_file, engine="openpyxl")
                last_row = len(target_df)
            else:
                last_row = 0
            self.start_row_entry.delete(0, tk.END)
            self.start_row_entry.insert(0, str(last_row))

    def merge_columns(self):
        if not self.file_list or not self.target_file:
            messagebox.showerror("Błąd", "Wybierz pliki źródłowe i docelowy")
            return

        selected_indices = self.column_listbox.curselection()
        selected_columns = [self.column_listbox.get(i) for i in selected_indices]

        if not selected_columns:
            messagebox.showerror("Błąd", "Nie wybrano żadnych kolumn")
            return

        # pobieramy aktualny wiersz startowy z pola
        try:
            start_row = int(self.start_row_entry.get())
            if start_row < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Błąd", "Wiersz startowy musi być liczbą całkowitą >= 0")
            return

        # wczytanie lub utworzenie pliku docelowego
        if os.path.exists(self.target_file):
            target_df = pd.read_excel(self.target_file, engine="openpyxl")
        else:
            target_df = pd.DataFrame()

        for file in self.file_list:
            df = pd.read_excel(file, engine="openpyxl")
            df_selected = df[selected_columns]

            # jeśli start_row większy niż aktualny rozmiar, uzupełniamy pustymi wierszami
            if start_row > len(target_df):
                filler = pd.DataFrame(index=range(start_row - len(target_df)), columns=target_df.columns)
                target_df = pd.concat([target_df, filler], ignore_index=True)

            target_df = pd.concat([target_df.iloc[:start_row], df_selected, target_df.iloc[start_row:]], ignore_index=True)
            start_row += len(df_selected)

        target_df.to_excel(self.target_file, index=False, engine="openpyxl")
        messagebox.showinfo("Sukces", "Kolumny zostały skopiowane!")
        # odświeżenie numeru wiersza startowego
        self.start_row_entry.delete(0, tk.END)
        self.start_row_entry.insert(0, str(len(target_df)))

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = XLSXMergerApp(root)
    root.geometry("500x600")
    root.mainloop()
