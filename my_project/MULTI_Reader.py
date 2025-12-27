import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ============================================================
# CONFIGURATION
# ============================================================

# Column indexes for TXT / TXT4 files (0-based indexing)
# Order: Kod, ProduktNazwa, Cena, VAT
TXT_COLUMN_INDEXES = [0, 3, 2, 4]

# Final column names used in output files
TARGET_COLUMNS = ["Kod", "ProduktNazwa", "Cena", "VAT"]


# ============================================================
# MAIN APPLICATION CLASS
# ============================================================

class UniversalMergerApp:
    def __init__(self, root):
        # Root window configuration
        self.root = root
        self.root.title("CSV/TXT/TXT4 → XLSX / CSV Merger + Preview")

        # Internal state
        self.files = []          # List of selected input files
        self.target_file = None  # Target XLSX file path

        # ====================================================
        # FILE SELECTION CONTROLS
        # ====================================================

        tk.Label(root, text="Select files (CSV, TXT, TXT4)").pack(pady=5)
        tk.Button(root, text="Select files", command=self.load_files).pack()

        # ====================================================
        # LISTBOX WITH SCROLLBAR (SELECTED FILES)
        # ====================================================

        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)

        self.scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(
            frame,
            height=8,
            yscrollcommand=self.scrollbar.set
        )

        self.scrollbar.config(command=self.listbox.yview)

        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ====================================================
        # PREVIEW WINDOW (FIRST 5 ROWS OF FIRST FILE)
        # ====================================================

        tk.Label(
            root,
            text="Preview of first 5 rows from the first selected file"
        ).pack(pady=5)

        self.preview_text = tk.Text(root, height=10)
        self.preview_text.pack(fill=tk.BOTH, padx=10, pady=5, expand=False)

        # ====================================================
        # TARGET FILE SELECTION & MERGE BUTTON
        # ====================================================

        tk.Button(
            root,
            text="Select target XLSX file",
            command=self.select_target
        ).pack(pady=5)

        self.target_label = tk.Label(root, text="Target file: none")
        self.target_label.pack()

        tk.Button(
            root,
            text="Merge data + Sort",
            command=self.merge
        ).pack(pady=10)

        # Initial window size
        root.geometry("800x600")

        # Bind preview refresh to listbox selection
        self.listbox.bind("<<ListboxSelect>>", self.show_preview)

    # ========================================================
    # FILE LOADING
    # ========================================================

    def load_files(self):
        """Open file dialog and add selected files to the list."""
        files = filedialog.askopenfilenames(
            filetypes=[("CSV/TXT/TXT4 files", "*.csv *.txt *.txt4")]
        )

        for file in files:
            if file not in self.files:
                self.files.append(file)

        self.update_listbox()

    def update_listbox(self):
        """Refresh listbox content with selected files."""
        self.listbox.delete(0, tk.END)

        for idx, file in enumerate(self.files, start=1):
            display_name = f"{idx}. {os.path.basename(file)}"
            self.listbox.insert(tk.END, display_name)

        # Always refresh preview after update
        self.show_preview()

    # ========================================================
    # TARGET FILE SELECTION
    # ========================================================

    def select_target(self):
        """Select target XLSX output file."""
        self.target_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if self.target_file:
            self.target_label.config(text=os.path.basename(self.target_file))

    # ========================================================
    # PREVIEW HANDLING
    # ========================================================

    def show_preview(self, event=None):
        """Show preview of first 5 rows from the first selected file."""
        self.preview_text.delete(1.0, tk.END)

        if not self.files:
            return

        file = self.files[0]  # Preview always uses the first file
        preview_lines = []

        try:
            # ---------------- CSV PREVIEW ----------------
            if file.lower().endswith(".csv"):
                df = pd.read_csv(
                    file,
                    sep=";",
                    encoding="cp1250",
                    dtype=str,
                    keep_default_na=False
                )

                # Map existing CSV columns to target columns
                mapped_cols = {}
                for col in TARGET_COLUMNS:
                    match = [c for c in df.columns if col.lower() in c.lower()]
                    if match:
                        mapped_cols[match[0]] = col
                    else:
                        df[col] = ""
                        mapped_cols[col] = col

                preview_df = df[list(mapped_cols.keys())].head(5)
                preview_lines = preview_df.to_string(index=False).split("\n")

            # ---------------- TXT / TXT4 PREVIEW ----------------
            else:
                with open(file, "r", encoding="utf-8", errors="replace") as f:
                    for i, line in enumerate(f):
                        if i >= 5:
                            break

                        parts = line.rstrip("\n").split(";")

                        # Ensure enough columns exist
                        while len(parts) < max(TXT_COLUMN_INDEXES) + 1:
                            parts.append("")

                        selected = [parts[j] for j in TXT_COLUMN_INDEXES]
                        preview_lines.append("; ".join(selected))

            self.preview_text.insert(tk.END, "\n".join(preview_lines))

        except Exception as e:
            self.preview_text.insert(
                tk.END,
                f"Failed to read file preview: {e}"
            )

    # ========================================================
    # MERGE & SAVE LOGIC
    # ========================================================

    def merge(self):
        """Merge all selected files, sort data and save XLSX + CSV."""
        if not self.files or not self.target_file:
            messagebox.showerror(
                "Error",
                "Please select input files and a target file"
            )
            return

        # Load existing target file if it exists
        if os.path.exists(self.target_file):
            target_df = pd.read_excel(self.target_file)
        else:
            target_df = pd.DataFrame(columns=TARGET_COLUMNS)

        # Process each selected file
        for file in self.files:
            df_list = []

            try:
                # ---------------- CSV ----------------
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

                    selected_df = (
                        df[list(mapped_cols.keys())]
                        .rename(columns=mapped_cols)
                    )

                # ---------------- TXT / TXT4 ----------------
                else:
                    with open(file, "r", encoding="utf-8", errors="replace") as f:
                        for line in f:
                            parts = line.rstrip("\n").split(";")

                            while len(parts) < max(TXT_COLUMN_INDEXES) + 1:
                                parts.append("")

                            selected = [parts[i] for i in TXT_COLUMN_INDEXES]
                            df_list.append(selected)

                    selected_df = pd.DataFrame(
                        df_list,
                        columns=TARGET_COLUMNS
                    )

                target_df = pd.concat(
                    [target_df, selected_df],
                    ignore_index=True
                )

            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"Failed to read file {os.path.basename(file)}: {e}"
                )
                return

        # Sort by newest entries and keep unique product codes
        target_df = target_df.iloc[::-1].reset_index(drop=True)
        target_df = target_df.drop_duplicates(
            subset="Kod",
            keep="first"
        )

        # Save XLSX
        target_df.to_excel(self.target_file, index=False)

        # Save CSV
        csv_output = os.path.splitext(self.target_file)[0] + ".csv"
        target_df.to_csv(
            csv_output,
            sep=";",
            index=False,
            encoding="utf-8"
        )

        messagebox.showinfo(
            "Success",
            "Data has been saved to XLSX and CSV"
        )

        # Reset state
        self.files.clear()
        self.update_listbox()


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalMergerApp(root)
    root.mainloop()
