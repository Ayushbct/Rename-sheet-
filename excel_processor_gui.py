import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils import column_index_from_string
import os
from pathlib import Path
from threading import Thread

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Processor")
        self.root.geometry("640x560")

        self.selected_folder = tk.StringVar()
        self.column1_var = tk.StringVar(value="Z")
        self.column2_var = tk.StringVar(value="AA")

        # Title
        tk.Label(root, text="Excel Sheet Processor",
                 font=("Arial", 16, "bold")).pack(pady=10)

        # Folder UI
        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(frame, text="Selected Folder:").pack(anchor=tk.W)

        tk.Entry(frame, textvariable=self.selected_folder,
                 state='readonly').pack(fill=tk.X, pady=5)

        tk.Button(frame, text="Browse Folder",
                  command=self.browse_folder,
                  bg="#4CAF50", fg="white").pack()

        # Column UI
        col_frame = tk.Frame(root)
        col_frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(col_frame, text="Column #1").grid(row=0, column=0)
        tk.Label(col_frame, text="Column #2").grid(row=1, column=0)

        tk.Entry(col_frame, textvariable=self.column1_var,
                 width=6).grid(row=0, column=1)
        tk.Entry(col_frame, textvariable=self.column2_var,
                 width=6).grid(row=1, column=1)

        # Process button
        self.process_btn = tk.Button(root, text="Process Files",
                                     command=self.start_processing,
                                     bg="#2196F3", fg="white")
        self.process_btn.pack(pady=10)

        # Log box
        self.output = scrolledtext.ScrolledText(root, state=tk.DISABLED)
        self.output.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Status
        self.status = tk.StringVar(value="Ready")
        tk.Label(root, textvariable=self.status,
                 relief=tk.SUNKEN).pack(fill=tk.X)

    # ---------------- UI HELPERS ----------------
    def log(self, msg):
        self.root.after(0, self._log, msg)

    def _log(self, msg):
        self.output.config(state=tk.NORMAL)
        self.output.insert(tk.END, msg)
        self.output.see(tk.END)
        self.output.config(state=tk.DISABLED)

    def set_status(self, msg):
        self.root.after(0, lambda: self.status.set(msg))

    # ---------------- CORE ----------------
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder.set(folder)
            self.log(f"Selected: {folder}\n")

    def parse_column(self, col):
        return column_index_from_string(col.strip().upper()) - 1

    def process_file(self, file_path, col1, col2):
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        col1_idx = self.parse_column(col1)
        col2_idx = self.parse_column(col2)

        cleaned = 0

        for row in sheet.iter_rows(min_row=2):
            for col_idx in [col1_idx, col2_idx]:
                if col_idx < len(row):
                    cell = row[col_idx]
                    if cell.value:
                        original = str(cell.value)
                        new = original.replace(" ", "").replace("-", "")
                        if new != original:
                            cell.value = new
                            cleaned += 1

        # Rename sheet if needed
        if len(wb.sheetnames) == 1:
            if wb.sheetnames[0] != "Sheet1":
                wb.active.title = "Sheet1"

        wb.save(file_path)
        return cleaned

    def process_all(self, folder):
        files = list(Path(folder).glob("*.xlsx"))

        if not files:
            self.log("❌ No .xlsx files found\n")
            return

        total = len(files)
        success, fail = 0, 0

        col1 = self.column1_var.get()
        col2 = self.column2_var.get()

        for i, file in enumerate(files, 1):
            name = os.path.basename(file)
            self.set_status(f"{i}/{total} Processing {name}")
            self.log(f"\n[{i}/{total}] {name}\n")

            try:
                cleaned = self.process_file(file, col1, col2)
                self.log(f"✔ Cleaned {cleaned} cells\n")
                success += 1
            except Exception as e:
                self.log(f"✗ Error: {e}\n")
                fail += 1

        self.set_status("Done")
        self.log(f"\nDone! Success: {success}, Failed: {fail}\n")

        self.root.after(0, lambda:
            messagebox.showinfo("Done",
                                f"Success: {success}\nFailed: {fail}")
        )

        self.root.after(0, lambda:
            self.process_btn.config(state=tk.NORMAL)
        )

    def start_processing(self):
        folder = self.selected_folder.get()

        if not folder:
            messagebox.showwarning("Warning", "Select folder first")
            return

        self.process_btn.config(state=tk.DISABLED)

        self.output.config(state=tk.NORMAL)
        self.output.delete(1.0, tk.END)
        self.output.config(state=tk.DISABLED)

        Thread(target=self.process_all,
               args=(folder,), daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()