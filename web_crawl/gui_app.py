import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook

import main  # imports your main.py (must be in same folder)


def list_sheets(xlsx_path: str):
    wb = load_workbook(xlsx_path, read_only=True)
    return wb.sheetnames


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("YaPro GC335/GC337 Tool")
        self.geometry("820x520")

        self.plates_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.plates_sheet = tk.StringVar()
        self.output_sheet = tk.StringVar()
        self.plates_col = tk.StringVar(value="A")
        self.vehicle_type = tk.StringVar(value="335")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        # Plates file
        ttk.Label(frm, text="Plates Excel (read plates from):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.plates_file, width=70).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse...", command=self.pick_plates_file).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Plates sheet:").grid(row=1, column=0, sticky="w", **pad)
        self.cbo_plates_sheet = ttk.Combobox(frm, textvariable=self.plates_sheet, state="readonly", width=30)
        self.cbo_plates_sheet.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Plates column (A/B/C...):").grid(row=1, column=2, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.plates_col, width=8).grid(row=1, column=3, sticky="w", **pad)

        # Output file
        ttk.Label(frm, text="Output Excel (write results into):").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_file, width=70).grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse...", command=self.pick_output_file).grid(row=2, column=2, **pad)

        ttk.Label(frm, text="Output sheet:").grid(row=3, column=0, sticky="w", **pad)
        self.cbo_output_sheet = ttk.Combobox(frm, textvariable=self.output_sheet, state="readonly", width=30)
        self.cbo_output_sheet.grid(row=3, column=1, sticky="w", **pad)

        # Vehicle type
        box = ttk.LabelFrame(frm, text="Vehicle Type / Endpoint")
        box.grid(row=4, column=0, columnspan=4, sticky="we", **pad)

        ttk.Radiobutton(box, text="Gasoline (GC335)", variable=self.vehicle_type, value="335").pack(side="left", padx=10, pady=6)
        ttk.Radiobutton(box, text="Electric (GC337)", variable=self.vehicle_type, value="337").pack(side="left", padx=10, pady=6)

        # Run button
        self.btn_run = ttk.Button(frm, text="Run", command=self.on_run)
        self.btn_run.grid(row=5, column=0, sticky="w", **pad)

        self.lbl_status = ttk.Label(frm, text="Ready.")
        self.lbl_status.grid(row=5, column=1, columnspan=3, sticky="w", **pad)

        # Log box
        self.txt = tk.Text(frm, height=18)
        self.txt.grid(row=6, column=0, columnspan=4, sticky="nsew", **pad)

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)

    def log(self, s: str):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")

    def pick_plates_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.plates_file.set(path)
        try:
            sheets = list_sheets(path)
            self.cbo_plates_sheet["values"] = sheets
            if sheets:
                self.plates_sheet.set(sheets[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets:\n{e}")

    def pick_output_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.output_file.set(path)
        try:
            sheets = list_sheets(path)
            self.cbo_output_sheet["values"] = sheets
            if sheets:
                self.output_sheet.set(sheets[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets:\n{e}")

    def on_run(self):
        # quick validation
        if not self.plates_file.get().strip():
            messagebox.showwarning("Missing", "Please choose Plates Excel.")
            return
        if not self.output_file.get().strip():
            messagebox.showwarning("Missing", "Please choose Output Excel.")
            return
        if not self.plates_sheet.get().strip():
            messagebox.showwarning("Missing", "Please choose Plates sheet.")
            return
        if not self.output_sheet.get().strip():
            messagebox.showwarning("Missing", "Please choose Output sheet.")
            return
        if not self.plates_col.get().strip():
            messagebox.showwarning("Missing", "Please enter Plates column.")
            return

        self.btn_run.config(state="disabled")
        self.lbl_status.config(text="Running...")
        self.log("=== START ===")

        t = threading.Thread(target=self._run_worker, daemon=True)
        t.start()

    def _run_worker(self):
        try:
            # main.run_job prints results; we’ll just show some progress here
            self.log(f"Plates: {self.plates_file.get()} | sheet={self.plates_sheet.get()} | col={self.plates_col.get()}")
            self.log(f"Output: {self.output_file.get()} | sheet={self.output_sheet.get()}")
            self.log(f"Mode: {self.vehicle_type.get()}")

            main.run_job(
    excel_path=self.plates_file.get(),
    sheet_name=self.plates_sheet.get(),
    plate_col=self.plates_col.get(),
    vehicle_type=self.vehicle_type.get(),
    headless=True,
    output_path=self.output_file.get(),
)

            self.log("=== DONE (saved) ===")
            self._ui_done("Done.")
        except Exception:
            err = traceback.format_exc()
            self.log(err)
            self._ui_done("Error (see log).")

    def _ui_done(self, status_text: str):
        def _():
            self.lbl_status.config(text=status_text)
            self.btn_run.config(state="normal")
        self.after(0, _)


if __name__ == "__main__":
    App().mainloop()