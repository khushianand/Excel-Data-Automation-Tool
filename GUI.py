import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import threading

from core_process_execution import (
    compare_excel_data,
    highlight_excel_rows,
    add_site_and_product,
    add_new_data_to_tracking,
)


class ExcelToolApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel Comparison & Highlighting Tool")
        master.geometry("720x720")
        master.resizable(True, True)

        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        # Notebook and tabs
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(pady=10, expand=True, fill='both')

        self.comparison_frame = ttk.Frame(self.notebook, padding=(10, 10))
        self.highlighting_frame = ttk.Frame(self.notebook, padding=(10, 10))
        self.add_site_product_frame = ttk.Frame(self.notebook, padding=(10, 10))
        self.add_new_data_frame = ttk.Frame(self.notebook, padding=(10, 10))

        self.notebook.add(self.comparison_frame, text="Compare & Split")
        self.notebook.add(self.highlighting_frame, text="Highlight Matching Rows")
        self.notebook.add(self.add_site_product_frame, text="Add Site&Product")
        self.notebook.add(self.add_new_data_frame, text="Add New Data")

        self.init_variables()
        self.create_comparison_widgets(self.comparison_frame)
        self.create_highlighting_widgets(self.highlighting_frame)
        self.create_add_site_product_widgets(self.add_site_product_frame)
        self.create_add_new_data_widgets(self.add_new_data_frame)
        self.create_status_bar()

    def init_variables(self):
        self.comp_tracking_file_path = tk.StringVar()
        self.comp_raw_file_path = tk.StringVar()
        self.comp_output_file_path = tk.StringVar()
        self.comp_tracking_sheet_name = tk.StringVar()
        self.comp_raw_sheet_name = tk.StringVar()
        self.comp_tracking_area = tk.StringVar()
        self.comp_selected_sites = []

        self.hl_file_to_highlight_path = tk.StringVar()
        self.hl_data_file_path = tk.StringVar()
        self.hl_file_to_highlight_sheet_name = tk.StringVar()
        self.hl_data_sheet_name = tk.StringVar()
        self.hl_color = tk.StringVar(value="ffff00")

        self.asp_file1_path = tk.StringVar()
        self.asp_file2_path = tk.StringVar()
        self.asp_sheet1_name = tk.StringVar()
        self.asp_sheet2_name = tk.StringVar()

        self.and_new_data_file_path = tk.StringVar()
        self.and_new_data_sheet_name = tk.StringVar() # ADDED
        self.and_tracking_file_path = tk.StringVar()
        self.and_tracking_sheet_name = tk.StringVar()
        self.and_highlight_color = tk.StringVar(value="ffff00")

    # ---------- Helper UI builders ----------
    def create_file_selector(self, parent, label_text, path_var, callback):
        ttk.Label(parent, text=label_text).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(parent, textvariable=path_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(parent, text="Browse...", command=callback).grid(row=0, column=2, padx=5, pady=5)

    def get_excel_sheet_names(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return []
        try:
            xls = pd.ExcelFile(file_path)
            return xls.sheet_names
        except Exception as e:
            self.update_status(f"Error reading sheets from {os.path.basename(file_path)}: {e}")
            messagebox.showerror("File Read Error", f"Could not read sheet names from {os.path.basename(file_path)}: {e}")
            return []

    def populate_sheet_dropdown(self, combobox_widget, file_path, sheet_var):
        sheet_names = self.get_excel_sheet_names(file_path)
        combobox_widget['values'] = sheet_names
        if sheet_names:
            sheet_var.set(sheet_names[0])
            combobox_widget.config(state="readonly")
        else:
            sheet_var.set("")
            combobox_widget.config(state="disabled")

    def update_status(self, message):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")

    def clear_status_log(self):
        self.status_text.config(state="normal")
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state="disabled")

    # ---------- Tab: Comparison ----------
    def create_comparison_widgets(self, parent):
        tracking_frame = ttk.LabelFrame(parent, text="Tracking Excel File", padding=(10, 10))
        tracking_frame.pack(padx=10, pady=5, fill="x", expand=True)

        self.create_file_selector(tracking_frame, "File Path:", self.comp_tracking_file_path, self.on_tracking_file_selected)

        ttk.Label(tracking_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.comp_tracking_sheet_combo = ttk.Combobox(tracking_frame, textvariable=self.comp_tracking_sheet_name, state="disabled", width=47)
        self.comp_tracking_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)
        self.comp_tracking_sheet_combo.bind("<<ComboboxSelected>>", self.populate_filters_dropdowns)

        ttk.Label(tracking_frame, text="Area Filter:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.comp_tracking_area_combo = ttk.Combobox(tracking_frame, textvariable=self.comp_tracking_area, state="disabled", width=47)
        self.comp_tracking_area_combo.grid(row=2, column=1, padx=5, pady=5, sticky="ew", columnspan=2)
        self.comp_tracking_area_combo.set("Select a sheet first")

        ttk.Label(tracking_frame, text="Sites Filter:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.comp_tracking_sites_button = ttk.Button(tracking_frame, text="Select Sites", command=self.open_sites_multiselect, state="disabled")
        self.comp_tracking_sites_button.grid(row=3, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        raw_frame = ttk.LabelFrame(parent, text="Raw Excel File", padding=(10, 10))
        raw_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(raw_frame, "File Path:", self.comp_raw_file_path, self.on_raw_file_selected)

        ttk.Label(raw_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.comp_raw_sheet_combo = ttk.Combobox(raw_frame, textvariable=self.comp_raw_sheet_name, state="disabled", width=47)
        self.comp_raw_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        output_frame = ttk.LabelFrame(parent, text="Output File", padding=(10, 10))
        output_frame.pack(padx=10, pady=5, fill="x", expand=True)
        ttk.Label(output_frame, text="Save As:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(output_frame, textvariable=self.comp_output_file_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(output_frame, text="Browse...", command=self.select_comparison_output_file).grid(row=0, column=2, padx=5, pady=5)

        run_button_frame = ttk.Frame(parent, padding=(10,5))
        run_button_frame.pack(pady=10)
        ttk.Button(run_button_frame, text="Run Comparison", command=self.run_comparison, style='TButton').pack()

        tracking_frame.grid_columnconfigure(1, weight=1)
        raw_frame.grid_columnconfigure(1, weight=1)
        output_frame.grid_columnconfigure(1, weight=1)

    def populate_filters_dropdowns(self, event=None):
        file_path = self.comp_tracking_file_path.get()
        sheet_name = self.comp_tracking_sheet_name.get()

        self.comp_tracking_area_combo['values'] = []
        self.comp_tracking_area.set("Select a sheet first")
        self.comp_tracking_area_combo.config(state="disabled")

        self.comp_tracking_sites_button.config(state="disabled")
        self.comp_selected_sites = []

        if not file_path or not os.path.exists(file_path) or not sheet_name:
            return

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if 'Area' in df.columns:
                areas = sorted(df['Area'].dropna().astype(str).str.strip().unique().tolist())
                self.comp_tracking_area_combo['values'] = areas
                if areas:
                    self.comp_tracking_area_combo.set(areas[0])
                self.comp_tracking_area_combo.config(state="readonly")
            else:
                self.update_status(f"Warning: 'Area' column not found in '{sheet_name}'. Area filter disabled.")
                self.comp_tracking_area_combo.set("'Area' column missing")
                self.comp_tracking_area_combo.config(state="disabled")

            if 'Site' in df.columns:
                self.comp_tracking_sites_button.config(state="enabled")
            else:
                self.update_status(f"Warning: 'Site' column not found in '{sheet_name}'. Sites filter disabled.")
                self.comp_tracking_sites_button.config(state="disabled")

        except Exception as e:
            self.update_status(f"Error reading filters: {e}")
            self.comp_tracking_area_combo.set("Error loading areas")
            self.comp_tracking_area_combo.config(state="disabled")

    def open_sites_multiselect(self):
        file_path = self.comp_tracking_file_path.get()
        sheet_name = self.comp_tracking_sheet_name.get()

        if not file_path or not sheet_name:
            messagebox.showwarning("Missing Information", "Please select a tracking file and sheet first.")
            return

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if 'Site' not in df.columns:
                messagebox.showerror("Column Missing", "'Site' column not found in the selected sheet.")
                return

            sites = sorted(df['Site'].dropna().astype(str).str.strip().unique().tolist())

            sites_window = tk.Toplevel(self.master)
            sites_window.title("Select Sites")
            sites_window.geometry("300x400")

            frame = ttk.Frame(sites_window)
            frame.pack(fill='both', expand=True, padx=10, pady=10)

            canvas = tk.Canvas(frame)
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(
                    scrollregion=canvas.bbox("all")
                )
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            selected_vars = {}
            for site in sites:
                var = tk.BooleanVar(value=(site in self.comp_selected_sites))
                ttk.Checkbutton(scrollable_frame, text=site, variable=var).pack(anchor='w')
                selected_vars[site] = var

            def on_ok():
                self.comp_selected_sites = [site for site, var in selected_vars.items() if var.get()]
                self.comp_tracking_sites_button.config(text=f"Sites ({len(self.comp_selected_sites)} selected)")
                sites_window.destroy()

            ttk.Button(sites_window, text="OK", command=on_ok).pack(pady=5)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

        except Exception as e:
            messagebox.showerror("Error", f"Could not load sites: {e}")

    def on_tracking_file_selected(self):
        file_path = filedialog.askopenfilename(title="Select Tracking Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.comp_tracking_file_path.set(file_path)
            self.populate_sheet_dropdown(self.comp_tracking_sheet_combo, file_path, self.comp_tracking_sheet_name)
            self.populate_filters_dropdowns()

    def on_raw_file_selected(self):
        file_path = filedialog.askopenfilename(title="Select Raw Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.comp_raw_file_path.set(file_path)
            self.populate_sheet_dropdown(self.comp_raw_sheet_combo, file_path, self.comp_raw_sheet_name)

    def select_comparison_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")), title="Save Output Excel File As")
        if file_path:
            self.comp_output_file_path.set(file_path)

    def run_comparison(self):
        t_file = self.comp_tracking_file_path.get()
        t_sheet = self.comp_tracking_sheet_name.get()
        r_file = self.comp_raw_file_path.get()
        r_sheet = self.comp_raw_sheet_name.get()
        selected_area = self.comp_tracking_area.get()
        output_file = self.comp_output_file_path.get()

        if not all([t_file, t_sheet, r_file, r_sheet, output_file]):
            messagebox.showwarning("Missing Input", "Please select all files, sheets, and an output path.")
            return

        if self.comp_tracking_area_combo['state'] == 'readonly' and not selected_area:
            messagebox.showwarning("Missing Input", "Please select an 'Area' from the dropdown for the Tracking file.")
            return

        self.clear_status_log()
        thread = threading.Thread(
            target=compare_excel_data,
            args=(t_file, t_sheet, r_file, r_sheet, self.update_status, output_file, selected_area, self.comp_selected_sites),
            daemon=True,
        )
        thread.start()

    # ---------- Tab: Highlighting ----------
    def create_highlighting_widgets(self, parent):
        hl_file_frame = ttk.LabelFrame(parent, text="File to Highlight", padding=(10, 10))
        hl_file_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(hl_file_frame, "File:", self.hl_file_to_highlight_path, self.on_file_to_highlight_selected)
        ttk.Label(hl_file_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.hl_file_to_highlight_sheet_combo = ttk.Combobox(hl_file_frame, textvariable=self.hl_file_to_highlight_sheet_name, state="disabled", width=47)
        self.hl_file_to_highlight_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        hl_data_frame = ttk.LabelFrame(parent, text="Data File (for comparison)", padding=(10, 10))
        hl_data_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(hl_data_frame, "File:", self.hl_data_file_path, self.on_data_file_selected)
        ttk.Label(hl_data_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.hl_data_sheet_combo = ttk.Combobox(hl_data_frame, textvariable=self.hl_data_sheet_name, state="disabled", width=47)
        self.hl_data_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        color_frame = ttk.Frame(parent, padding=(10,5))
        color_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(color_frame, text="Highlight Color (Hex, e.g., '#ffff00' or 'ffff00'):").pack(side='left', padx=5, pady=5)
        ttk.Entry(color_frame, textvariable=self.hl_color, width=10).pack(side='left', padx=5, pady=5)

        run_button_frame = ttk.Frame(parent, padding=(10,5))
        run_button_frame.pack(pady=10)
        ttk.Button(run_button_frame, text="Run Highlighting", command=self.run_highlighting, style='TButton').pack()

        hl_file_frame.grid_columnconfigure(1, weight=1)
        hl_data_frame.grid_columnconfigure(1, weight=1)

    def on_file_to_highlight_selected(self):
        file_path = filedialog.askopenfilename(title="Select File to Highlight", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.hl_file_to_highlight_path.set(file_path)
            self.populate_sheet_dropdown(self.hl_file_to_highlight_sheet_combo, file_path, self.hl_file_to_highlight_sheet_name)

    def on_data_file_selected(self):
        file_path = filedialog.askopenfilename(title="Select Data File for Comparison", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.hl_data_file_path.set(file_path)
            self.populate_sheet_dropdown(self.hl_data_sheet_combo, file_path, self.hl_data_sheet_name)

    def run_highlighting(self):
        file_to_highlight = self.hl_file_to_highlight_path.get()
        sheet_to_highlight = self.hl_file_to_highlight_sheet_name.get()
        data_file = self.hl_data_file_path.get()
        data_sheet = self.hl_data_sheet_name.get()
        color = self.hl_color.get()

        if not all([file_to_highlight, sheet_to_highlight, data_file, data_sheet, color]):
            messagebox.showwarning("Missing Input", "Please select both files and their sheets, and provide a highlight color.")
            return

        self.clear_status_log()
        threading.Thread(
            target=highlight_excel_rows,
            args=(file_to_highlight, sheet_to_highlight, data_file, data_sheet, color, self.update_status),
            daemon=True,
        ).start()

    # ---------- Tab: Add Site & Product ----------
    def create_add_site_product_widgets(self, parent):
        file1_frame = ttk.LabelFrame(parent, text="1st Excel File (Source for IPs)", padding=(10, 10))
        file1_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(file1_frame, "File Path:", self.asp_file1_path, self.on_asp_file1_selected)
        ttk.Label(file1_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.asp_sheet1_combo = ttk.Combobox(file1_frame, textvariable=self.asp_sheet1_name, state="disabled", width=47)
        self.asp_sheet1_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        file2_frame = ttk.LabelFrame(parent, text="2nd Excel File (Source for Site & Product)", padding=(10, 10))
        file2_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(file2_frame, "File Path:", self.asp_file2_path, self.on_asp_file2_selected)
        ttk.Label(file2_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.asp_sheet2_combo = ttk.Combobox(file2_frame, textvariable=self.asp_sheet2_name, state="disabled", width=47)
        self.asp_sheet2_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        run_button_frame = ttk.Frame(parent, padding=(10,5))
        run_button_frame.pack(pady=10)
        ttk.Button(run_button_frame, text="Run Add Site & Product", command=self.run_add_site_product, style='TButton').pack()

        file1_frame.grid_columnconfigure(1, weight=1)
        file2_frame.grid_columnconfigure(1, weight=1)

    def on_asp_file1_selected(self):
        file_path = filedialog.askopenfilename(title="Select 1st Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.asp_file1_path.set(file_path)
            self.populate_sheet_dropdown(self.asp_sheet1_combo, file_path, self.asp_sheet1_name)

    def on_asp_file2_selected(self):
        file_path = filedialog.askopenfilename(title="Select 2nd Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.asp_file2_path.set(file_path)
            self.populate_sheet_dropdown(self.asp_sheet2_combo, file_path, self.asp_sheet2_name)

    def run_add_site_product(self):
        file1 = self.asp_file1_path.get()
        sheet1 = self.asp_sheet1_name.get()
        file2 = self.asp_file2_path.get()
        sheet2 = self.asp_sheet2_name.get()

        if not all([file1, sheet1, file2, sheet2]):
            messagebox.showwarning("Missing Input", "Please select both files and their sheets.")
            return

        self.clear_status_log()
        threading.Thread(
            target=add_site_and_product,
            args=(file1, sheet1, file2, sheet2, self.update_status),
            daemon=True,
        ).start()

    # ---------- Tab: Add New Data ----------
    def create_add_new_data_widgets(self, parent):
        new_data_frame = ttk.LabelFrame(parent, text="New Data File (output of comparison)", padding=(10, 10))
        new_data_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(new_data_frame, "File Path:", self.and_new_data_file_path, self.select_new_data_file)
        
        # ADDED
        ttk.Label(new_data_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.and_new_data_sheet_combo = ttk.Combobox(new_data_frame, textvariable=self.and_new_data_sheet_name, state="disabled", width=47)
        self.and_new_data_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        tracking_frame = ttk.LabelFrame(parent, text="Tracking Excel File (destination)", padding=(10, 10))
        tracking_frame.pack(padx=10, pady=5, fill="x", expand=True)
        self.create_file_selector(tracking_frame, "File Path:", self.and_tracking_file_path, self.select_tracking_file)

        ttk.Label(tracking_frame, text="Sheet Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.and_tracking_sheet_combo = ttk.Combobox(tracking_frame, textvariable=self.and_tracking_sheet_name, state="disabled", width=47)
        self.and_tracking_sheet_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew", columnspan=2)

        color_frame = ttk.Frame(parent, padding=(10,5))
        color_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(color_frame, text="Highlight Color (Hex, e.g., '#ffff00' or 'ffff00'):").pack(side='left', padx=5, pady=5)
        ttk.Entry(color_frame, textvariable=self.and_highlight_color, width=10).pack(side='left', padx=5, pady=5)

        run_button_frame = ttk.Frame(parent, padding=(10,5))
        run_button_frame.pack(pady=10)
        ttk.Button(run_button_frame, text="Add Data to Tracking", command=self.run_add_new_data, style='TButton').pack()

        new_data_frame.grid_columnconfigure(1, weight=1)
        tracking_frame.grid_columnconfigure(1, weight=1)

    def select_new_data_file(self):
        file_path = filedialog.askopenfilename(title="Select New Data Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.and_new_data_file_path.set(file_path)
            self.populate_sheet_dropdown(self.and_new_data_sheet_combo, file_path, self.and_new_data_sheet_name) # ADDED

    def select_tracking_file(self):
        file_path = filedialog.askopenfilename(title="Select Tracking Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if file_path:
            self.and_tracking_file_path.set(file_path)
            self.populate_sheet_dropdown(self.and_tracking_sheet_combo, file_path, self.and_tracking_sheet_name)

    def run_add_new_data(self):
        new_data_file = self.and_new_data_file_path.get()
        new_data_sheet = self.and_new_data_sheet_name.get() # ADDED
        tracking_file = self.and_tracking_file_path.get()
        tracking_sheet = self.and_tracking_sheet_name.get()
        highlight_color = self.and_highlight_color.get()

        if not all([new_data_file, new_data_sheet, tracking_file, tracking_sheet, highlight_color]): # MODIFIED
            messagebox.showwarning("Missing Input", "Please select both files, their sheets, and a highlight color.")
            return

        self.clear_status_log()
        threading.Thread(
            target=add_new_data_to_tracking,
            args=(new_data_file, new_data_sheet, tracking_file, tracking_sheet, highlight_color, self.update_status), # MODIFIED
            daemon=True,
        ).start()

    # ---------- Status Bar ----------
    def create_status_bar(self):
        self.status_frame = ttk.LabelFrame(self.master, text="Status & Log", padding=(10, 10))
        self.status_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.status_text = tk.Text(self.status_frame, height=8, state="disabled", wrap="word", font=('Arial', 9))
        self.status_text.pack(fill="both", expand=True)
        self.update_status("Application started.")