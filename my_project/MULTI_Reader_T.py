import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ============================
# Configuration
# ============================

# Column indexes for TXT / TXT4 files (0-based)
# Order: Kod, ProduktNazwa, Cena, VAT
TXT_COLUMN_INDEXES = [0, 3, 2, 4]

# Final output columns
TARGET_COLUMNS = ["Kod", "ProduktNazwa", "Cena", "VAT"]


class UniversalMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV / TXT / TXT4 → XLSX / CSV Merger")

        self.files = []
        self.target_file = None

        # ============================
        # File selection
        # ============================
        tk.Label(root, text="Wybierz pliki (CSV, TXT, TXT4)").pack(pady=5)
        tk.Button(root, text="Wybierz pliki", command=self.load_files).pack()

        # ============================
        # Selected files listbox
        # ============================
        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)

        self.scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(
            frame,
            height=6,
            yscrollcommand=self.scrollbar.set
        )
        self.scrollbar.config(command=self.listbox.yview)

        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ============================
        # Preview window
        # ============================
        tk.Label(root, text="Podgląd (5 pierwszych wierszy – dane wynikowe)").pack(pady=5)

        self.preview_text = tk.Text(root, height=10, width=140)
        self.preview_text.pack(fill=tk.BOTH, padx=10, pady=5)

        # ============================
        # Options
        # ============================
        self.test_mode = tk.BooleanVar(value=False)
        self.clear_after_merge = tk.BooleanVar(value=True)

        tk.Checkbutton(
            root,
            text="Tryb testowy (bez zapisu plików)",
            variable=self.test_mode
        ).pack(anchor="w", padx=10)

        tk.Checkbutton(
            root,
            text="Wyczyść listę plików po zakończeniu",
            variable=self.clear_after_merge
        ).pack(anchor="w", padx=10)

        # ============================
        # Target file selection
        # ============================
        tk.Button(
            root,
            text="Wybierz plik docelowy XLSX",
            command=self.select_target
        ).pack(pady=5)

        self.target_label = tk.Label(root, text="Plik docelowy: brak")
        self.target_label.pack()

        # ============================
        # Action button
        # ============================
        tk.Button(
            root,
            text="Kopiuj dane + Sortowanie",
            command=self.merge
        ).pack(pady=10)

        # Window size
        root.geometry("800x600")

        # Preview refresh on listbox selection
        self.listbox.bind("<<ListboxSelect>>", self.show_preview)

    # ============================
    # Load files
    # ============================
    def load_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("CSV / TXT / TXT4", "*.csv *.txt *.txt4")]
        )
        for file in files:
            if file not in self.files:
                self.files.append(file)
        self.update_listbox()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, file in enumerate(self.files, start=1):
            self.listbox.insert(tk.END, f"{idx}. {os.path.basename(file)}")
        self.show_preview()

    # ============================
    # Target file
    # ============================
    def select_target(self):
        self.target_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.target_file:
            self.target_label.config(text=os.path.basename(self.target_file))

    # ============================
    # Preview logic
    # ============================
    def show_preview(self, event=None):
        self.preview_text.delete(1.0, tk.END)

        if not self.files:
            return

        file = self.files[0]
        preview_lines = []

        try:
            if file.lower().endswith(".csv"):
                df = pd.read_csv(
                    file,
                    sep=";",
                    encoding="cp1250",
                    dtype=str,
                    keep_default_na=False
                )

                mapped_cols = {}
                for col in TARGET_COLUMNS:
                    match = [c for c in df.columns if col.lower() in c.lower()]
                    if match:
                        mapped_cols[match[0]] = col
                    else:
                        df[col] = ""
                        mapped_cols[col] = col

                preview_df = df[list(mapped_cols.keys())].rename(columns=mapped_cols)
                preview_lines = preview_df.head(5).to_string(index=False).split("\n")

            else:
                with open(file, "r", encoding="utf-8", errors="replace") as f:
                    for i, line in enumerate(f):
                        if i >= 5:
                            break
                        parts = line.rstrip("\n").split(";")
                        while len(parts) <= max(TXT_COLUMN_INDEXES):
                            parts.append("")
                        selected = [parts[i] for i in TXT_COLUMN_INDEXES]
                        preview_lines.append("; ".join(selected))

            self.preview_text.insert(tk.END, "\n".join(preview_lines))

        except Exception as e:
            self.preview_text.insert(tk.END, f"Preview error: {e}")

    # ============================
    # Merge logic
    # ============================
    def merge(self):
        if not self.files:
            messagebox.showerror("Błąd", "Nie wybrano plików")
            return

        if not self.test_mode.get() and not self.target_file:
            messagebox.showerror("Błąd", "Nie wybrano pliku docelowego")
            return

        target_df = pd.DataFrame(columns=TARGET_COLUMNS)

        for file in self.files:
            try:
                if file.lower().endswith(".csv"):
                    df = pd.read_csv(
                        file,
                        sep=";",
                        encoding="cp1250",
                        dtype=str,
                        keep_default_na=False
                    )

                    mapped_cols = {}
                    for col in TARGET_COLUMNS:
                        match = [c for c in df.columns if col.lower() in c.lower()]
                        if match:
                            mapped_cols[match[0]] = col
                        else:
                            df[col] = ""
                            mapped_cols[col] = col

                    selected_df = df[list(mapped_cols.keys())].rename(columns=mapped_cols)

                else:
                    rows = []
                    with open(file, "r", encoding="utf-8", errors="replace") as f:
                        for line in f:
                            parts = line.rstrip("\n").split(";")
                            while len(parts) <= max(TXT_COLUMN_INDEXES):
                                parts.append("")
                            rows.append([parts[i] for i in TXT_COLUMN_INDEXES])

                    selected_df = pd.DataFrame(rows, columns=TARGET_COLUMNS)

                target_df = pd.concat([target_df, selected_df], ignore_index=True)

            except Exception as e:
                messagebox.showerror("Błąd", f"{os.path.basename(file)}: {e}")
                return

        # Sort newest first and keep unique Kod
        target_df = target_df.iloc[::-1].reset_index(drop=True)
        target_df = target_df.drop_duplicates(subset="Kod", keep="first")

        # ============================
        # Save (unless test mode)
        # ============================
        if not self.test_mode.get():
            target_df.to_excel(self.target_file, index=False)

            csv_output = os.path.splitext(self.target_file)[0] + ".csv"
            target_df.to_csv(csv_output, sep=";", index=False, encoding="utf-8")

            messagebox.showinfo("OK", "Dane zapisane do XLSX i CSV")
        else:
            messagebox.showinfo(
                "Tryb testowy",
                f"Przetworzono {len(target_df)} wierszy\nZapis pominięty"
            )

        # Clear file list if enabled
        if self.clear_after_merge.get():
            self.files.clear()
            self.update_listbox()


# ============================
# App start
# ============================
if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalMergerApp(root)
    root.mainloop()
