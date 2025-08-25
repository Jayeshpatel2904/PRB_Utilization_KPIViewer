import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
import matplotlib.ticker as mtick

# Define default thresholds for each KPI
KPI_THRESHOLDS = {
    'VoNR Retainability': 0.99,
    'ACC_VoNR_Accessibility': 0.99,
    'UTL_DL PRB utilization': 80
}

# Define the direction of the threshold check for each KPI (below or above)
KPI_CHECK_DIRECTION = {
    'VoNR Retainability': 'below',
    'ACC_VoNR_Accessibility': 'below',
    'UTL_DL PRB utilization': 'above'
}

# Define the predefined ranges for each KPI
KPI_RANGES = {
    'UTL_DL PRB utilization': [(0, 50, '0-50%'), (50, 60, '50-60%'), (60, 70, '60-70%'), (70, 101, '70%+')],
    'VoNR Retainability': [(0, 0.90, '0-90%'), (0.90, 0.95, '90-95%'), (0.95, 0.98, '95-98%'), (0.98, 0.99, '98-99%'), (0.99, 1.01, '99%+')],
    'ACC_VoNR_Accessibility': [(0, 0.90, '0-90%'), (0.90, 0.95, '90-95%'), (0.95, 0.98, '95-98%'), (0.98, 0.99, '98-99%'), (0.99, 1.01, '99%+')],
}

def extract_parts(cellname):
    siteid = cellname[:11]
    sectorid = cellname[12:13]
    band = cellname[14:]
    return pd.Series([siteid, sectorid, band])

class ExcelVisualizer:
    def __init__(self, root):
        self.root = root
        root.state('zoomed')
        self.root.title("Cell Utilization and KPI Viewer V1.2")
        self.tree = None
        self.df = None
        self.last_figure = None
        self.selected_kpi = None
        self.kpi_threshold = None
        self.kpi_check_direction = None
        self.file_path = None
        self.kpi_label = None
        self.progress_bar = None
        self.status_label = None
        self.create_widgets()

    def create_widgets(self):
        # Use a single, consistent grid layout for the entire application
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        
        # Load Data Button
        self.load_btn = tk.Button(self.root, text="Load Data File", bg="#599ACF", fg="white", font=('Arial', 12, 'bold'), command=self.ask_kpi_and_load)
        self.load_btn.grid(row=0, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # Centered KPI label
        self.kpi_label = ttk.Label(self.root, text="", font=('Arial', 12, 'bold'))
        self.kpi_label.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Treeview Frame
        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=2, column=0, sticky="ns", padx=5, pady=5)
        tree_scrollbar = ttk.Scrollbar(tree_frame)
        tree_scrollbar.pack(side="right", fill="y")
        self.treeview = ttk.Treeview(tree_frame, yscrollcommand=tree_scrollbar.set)
        self.treeview.pack(side="left", fill="y")
        tree_scrollbar.config(command=self.treeview.yview)
        self.treeview.bind("<<TreeviewSelect>>", self.on_tree_click)

        # Canvas Frame
        self.canvas_frame = ttk.Frame(self.root)
        self.canvas_frame.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)

        # Footer Frame
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        
        # Use grid for a more reliable footer layout
        footer_frame.grid_columnconfigure(0, weight=1)
        footer_frame.grid_columnconfigure(1, weight=0)
        
        # New label for "Ready" status
        self.status_label = ttk.Label(footer_frame, text="", foreground="#058d10", font=('Arial', 10, 'bold'))
        self.status_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(footer_frame, orient="horizontal", mode="indeterminate")
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        self.progress_bar.grid_remove() # Hide it initially

        contact_label = ttk.Label(footer_frame, text="support contact jayeshkumar.patel@dish.com", foreground="red", font=('Arial', 10, 'bold', 'italic'))
        contact_label.grid(row=0, column=1, padx=10, pady=5, sticky="e")

    def ask_kpi_and_load(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not self.file_path:
            return

        kpi_options = ['UTL_DL PRB utilization', 'VoNR Retainability', 'ACC_VoNR_Accessibility']
        
        select_window = tk.Toplevel(self.root)
        select_window.title("Select KPI")
        
        kpi_var = tk.StringVar(value=kpi_options[0])

        tk.Label(select_window, text="Please select the KPI to analyze:").pack(pady=10, padx=10)

        for option in kpi_options:
            tk.Radiobutton(select_window, text=option, variable=kpi_var, value=option).pack(anchor="w", padx=20)
        
        def on_ok():
            self.selected_kpi = kpi_var.get()
            select_window.destroy()
            self.start_loading_process()

        def on_cancel():
            select_window.destroy()
            messagebox.showinfo("Cancelled", "File loading was cancelled.")

        ok_button = tk.Button(select_window, text="OK", command=on_ok)
        ok_button.pack(side="left", padx=10, pady=10)
        cancel_button = tk.Button(select_window, text="Cancel", command=on_cancel)
        cancel_button.pack(side="right", padx=10, pady=10)
        
        self.root.wait_window(select_window)

    def start_loading_process(self):
        if not self.selected_kpi or not self.file_path:
            return
        
        # Set the KPI label text and ensure it is placed
        self.kpi_label.config(text=f"Selected KPI: {self.selected_kpi}")

        self.kpi_threshold = KPI_THRESHOLDS.get(self.selected_kpi, None)
        self.kpi_check_direction = KPI_CHECK_DIRECTION.get(self.selected_kpi, 'above')

        try:
            if self.selected_kpi == 'UTL_DL PRB utilization':
                threshold_input = simpledialog.askfloat("Threshold", "Enter threshold percentage for UTL_DL PRB utilization (e.g. 80 for 80%)", minvalue=0.0, maxvalue=100.0, initialvalue=self.kpi_threshold, parent=self.root)
                if threshold_input is not None:
                    self.kpi_threshold = threshold_input
            else:
                threshold_input = simpledialog.askfloat("Threshold", f"Enter threshold percentage for {self.selected_kpi} (e.g. 99 for 99%):", minvalue=0.0, maxvalue=100.0, initialvalue=self.kpi_threshold * 100, parent=self.root)
                if threshold_input is not None:
                    self.kpi_threshold = threshold_input / 100
        except Exception:
            messagebox.showwarning("Warning", "Invalid input. Using default threshold.")

        self.treeview.delete(*self.treeview.get_children())
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

        self.load_btn.config(bg="#3d3d3d", text="Data Loading.....", state="disabled")
        # Hide status label and show progress bar
        self.status_label.grid_remove()
        self.progress_bar.grid()
        self.progress_bar.start()

        threading.Thread(target=self.process_file, args=(self.file_path,), daemon=True).start()

    def process_file(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            self.df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
            
            # Sanitize column names
            self.df.columns = self.df.columns.str.strip()
            
            # Check for required columns
            required_cols = ['DATETIME', 'SELECTION_0_NAME', 'CELLNAME', self.selected_kpi]
            
            # Temporarily add SITEID and SECTORID to handle the extract_parts function
            temp_df = self.df.copy()
            if 'SITEID' not in temp_df.columns or 'SECTORID' not in temp_in_df.columns:
                temp_df[['SITEID', 'SECTORID', 'BAND']] = temp_df['CELLNAME'].apply(extract_parts)

            if not all(col in temp_df.columns for col in required_cols):
                missing_cols = [col for col in required_cols if col not in temp_df.columns]
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
            
            self.df = temp_df

            # Data type conversions
            self.df['CELLNAME'] = self.df['CELLNAME'].astype(str).str.strip()
            self.df['DATETIME'] = pd.to_datetime(self.df['DATETIME'], errors='coerce')
            self.df[self.selected_kpi] = pd.to_numeric(self.df[self.selected_kpi], errors='coerce')

            if self.selected_kpi == 'UTL_DL PRB utilization':
                self.df[self.selected_kpi] = self.df[self.selected_kpi] * 100

            self.populate_tree()
            self.load_btn.config(bg="#058d10", text="Data Loaded", state="normal")
            
            # Hide progress bar and show status label
            self.progress_bar.grid_remove()
            self.status_label.config(text="Ready")
            self.status_label.grid()

        except Exception as e:
            self.status_label.config(text=f"Error: {e}")
            self.load_btn.config(bg="red", text="Load Failed", state="normal")
            
            # Hide progress bar on error
            self.progress_bar.grid_remove()
            self.status_label.grid()
            
    def populate_tree(self):
        self.treeview.delete(*self.treeview.get_children())
        
        if self.kpi_check_direction == 'above':
            normal_root = self.treeview.insert("", "end", text="Normal", open=True)
            alert_root = self.treeview.insert("", "end", text="Over Threshold", open=True)
        else:
            normal_root = self.treeview.insert("", "end", text="Normal", open=True)
            alert_root = self.treeview.insert("", "end", text="Below Threshold", open=True)

        normal_selections = {}
        alert_selections = {}

        grouped = self.df.groupby(['SELECTION_0_NAME', 'SITEID', 'SECTORID'])

        for (sel_name, siteid, sectorid), group in grouped:
            group = group.sort_values('DATETIME')
            
            last_15_days = group.iloc[int(len(group) / 2):]
            
            if last_15_days[self.selected_kpi].empty:
                continue

            if self.selected_kpi == 'UTL_DL PRB utilization':
                # Check if any data point in the last 15 days is above the threshold
                condition_met = (last_15_days[self.selected_kpi] > self.kpi_threshold).any()
            elif self.kpi_check_direction == 'above':
                average_value = last_15_days[self.selected_kpi].mean()
                condition_met = average_value > self.kpi_threshold
            else:
                average_value = last_15_days[self.selected_kpi].mean()
                condition_met = average_value < self.kpi_threshold

            root_map = alert_selections if condition_met else normal_selections
            top_node = alert_root if condition_met else normal_root

            if sel_name not in root_map:
                root_map[sel_name] = self.treeview.insert(top_node, "end", text=sel_name, open=True)

            self.treeview.insert(root_map[sel_name], "end", text=f"{siteid} / Sector {sectorid}", open=False)

    def on_tree_click(self, event):
        selected = self.treeview.selection()
        if not selected:
            return

        item_text = self.treeview.item(selected[0], 'text')
        parent_item = self.treeview.parent(selected[0])
        if '/' not in item_text or not parent_item:
            return

        try:
            sel_name = self.treeview.item(parent_item, 'text')
            if sel_name in ["Normal", "Over Threshold", "Below Threshold"]:
                return
            siteid, sector_text = item_text.split(" / ")
            sectorid = sector_text.replace("Sector ", "")
        except Exception:
            messagebox.showerror("Error", "Invalid tree item format.")
            return

        sector_data = self.df[
            (self.df['SELECTION_0_NAME'] == sel_name) &
            (self.df['SITEID'] == siteid) &
            (self.df['SECTORID'] == sectorid)
        ]

        if sector_data.empty:
            messagebox.showerror("Error", f"No data found for {item_text}")
            return

        self.plot_charts(sector_data)

    def plot_charts(self, data):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

        data = data.dropna(subset=['DATETIME', self.selected_kpi])
        cellnames = data['CELLNAME'].unique()

        if len(cellnames) == 0:
            messagebox.showinfo("Info", "No cellnames available for selected sector.")
            return

        fig, axs = plt.subplots(len(cellnames), 1, figsize=(10, 5 * len(cellnames)))
        if len(cellnames) == 1:
            axs = [axs]
        
        for ax, cell in zip(axs, cellnames):
            cell_data = data[data['CELLNAME'] == cell]
            if cell_data.empty:
                continue

            avg_kpi_by_time = cell_data.groupby('DATETIME')[self.selected_kpi].mean()
            ax.plot(avg_kpi_by_time.index, avg_kpi_by_time.values, label=f"{cell}", marker='o')

            if self.selected_kpi == 'UTL_DL PRB utilization':
                threshold_display = f'Threshold: {self.kpi_threshold:.0f}%'
                ranges = KPI_RANGES['UTL_DL PRB utilization']
                ax.set_ylabel("Utilization (%)")
                ax.yaxis.set_major_formatter(mtick.PercentFormatter())
            else:
                threshold_display = f'Threshold: {self.kpi_threshold * 100:.0f}%'
                ranges = KPI_RANGES[self.selected_kpi]
                ax.set_ylabel(f"{self.selected_kpi} (%)")
                ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=1.0))
            
            # Count samples in predefined ranges
            stats_text_parts = []
            for lower, upper, label in ranges:
                if self.selected_kpi in ['VoNR Retainability', 'ACC_VoNR_Accessibility']:
                    count = cell_data[(cell_data[self.selected_kpi] >= lower) & (cell_data[self.selected_kpi] < upper)].shape[0]
                else:
                    count = cell_data[(cell_data[self.selected_kpi] >= lower) & (cell_data[self.selected_kpi] < upper)].shape[0]
                stats_text_parts.append(f"{label}: {count}")
            stats_text = ' | '.join(stats_text_parts)

            ax.axhline(self.kpi_threshold, color='red', linestyle='--', label=threshold_display)
            
            ax.set_title(f"CELLNAME: {cell} [{stats_text}]")
            ax.set_xlabel("Datetime")
            ax.legend()
            ax.grid(True)

        plt.subplots_adjust(hspace=0.5)
        self.last_figure = fig
        canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelVisualizer(root)
    root.mainloop()
