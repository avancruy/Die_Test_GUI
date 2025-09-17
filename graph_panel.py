import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

class GraphPanel:
    def __init__(self):
        self.current_excel_file = None
        self.fig = None
        self.canvas = None
        self.toolbar = None
        self.excel_path_var = None

    def setup_graph_panel(self, main_panel):
        # Graph panel title
        graph_panel_title = tk.Label(main_panel, text="Real-time Data Visualization",
                               font=('Arial', 12, 'bold'), bg='#f5f5f5', fg='#333333')
        graph_panel_title.pack(pady=(0, 10))

        # --- Graph Controls ---
        graph_controls_frame = ttk.LabelFrame(main_panel, text="Graph Controls", padding="8")
        graph_controls_frame.pack(fill='x', pady=(0, 10))

        # Compact controls layout
        controls_grid = tk.Frame(graph_controls_frame)
        controls_grid.pack(fill='x')

        # Excel file selection (smaller)
        tk.Label(controls_grid, text="Excel:", font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=(0, 5),
                                                                               pady=9, sticky=tk.W)

        self.excel_path_var = tk.StringVar()
        excel_entry = ttk.Entry(controls_grid, textvariable=self.excel_path_var, width=20)
        excel_entry.grid(row=0, column=1, padx=(0, 5), sticky=tk.EW)

        browse_excel_button = ttk.Button(controls_grid, text="📁", command=self.browse_excel_file, width=3)
        browse_excel_button.grid(row=0, column=2, padx=(0, 10))

        # Plot and clear buttons
        plot_button = tk.Button(controls_grid, text="📊 Plot", command=self.plot_excel_data,
                                bg='#2196F3', fg='white', font=('Impact', 8, 'normal'), relief='raised', bd=1)
        plot_button.grid(row=0, column=3, padx=(10, 10))

        clear_button = tk.Button(controls_grid, text="🗑 Clear", command=self.clear_plot,
                                 bg='#FF5722', fg='white', font=('Impact', 8, 'normal'), relief='raised', bd=1)
        clear_button.grid(row=0, column=4, padx=(0, 10))

        # Auto-plot checkbox
        self.auto_plot_var = tk.BooleanVar(value=True)
        auto_plot_check = ttk.Checkbutton(controls_grid, text="Auto-plot after test", variable=self.auto_plot_var)
        auto_plot_check.grid(row=0, column=5, padx=5, sticky=tk.W)

        controls_grid.columnconfigure(1, weight=1)

        # --- Compact Status Display ---
        self.graph_status_var = tk.StringVar(value="Ready")
        self.graph_status = tk.Label(main_panel, textvariable=self.graph_status_var,
                                font=('Arial', 10, 'bold'), fg='#333333', bg='#f5f5f5')
        self.graph_status.pack(pady=(0, 10))

        # --- Graph Display Area ---
        graph_display_frame = ttk.LabelFrame(main_panel, text="Live Graph", padding="5")
        graph_display_frame.pack(expand=True, fill='both')

        # Initialize matplotlib figure with smaller size for side panel
        self.fig, self.ax = plt.subplots(figsize=(6, 5))
        self.fig.tight_layout(pad=3.0)  # Increased padding to ensure labels fit

        # Create canvas and toolbar
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_display_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, graph_display_frame)
        self.toolbar.update()

        # Pack toolbar and canvas
        self.toolbar.pack(side='top', fill='x')
        self.canvas.get_tk_widget().pack(side='top', fill=tk.BOTH, expand=True)

        # Initialize with clean empty plot
        self.ax.set_visible(False)
        self.canvas.draw()

    def browse_excel_file(self):
        """Browse for an Excel file to plot"""
        file_path = filedialog.askopenfilename(
            title="Select Excel file to plot",
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.current_excel_file = file_path

    def plot_excel_data(self):
        """Plot data from the selected Excel file"""
        excel_file_path = self.excel_path_var.get()

        if not excel_file_path:
            messagebox.showwarning("No File Selected", "Please select an Excel file first.")
            return

        if not os.path.exists(excel_file_path):
            messagebox.showerror("File Not Found", f"The specified file does not exist:\n{excel_file_path}")
            return

        try:
            # Read the Excel file into a pandas DataFrame
            if excel_file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(excel_file_path)
            elif excel_file_path.endswith(".csv"):
                df = pd.read_csv(excel_file_path)
            else:
                raise Exception
        except Exception as e:
            messagebox.showerror("Error Reading Excel", f"Could not read the Excel file. Error: {e}")
            return

        # Get the filename to determine test type
        filename = os.path.basename(excel_file_path)
        # Determine column mappings based on test type
        if '_LIV_' in filename:
            # Original LIV test columns
            x_col = 'SMU1_Ch2_Laser_Current_Set_mA'
            y1_col = 'SMU1_Ch1_PD_Current_Meas_mA'
            y2_col = 'SMU1_Ch2_Laser_Voltage_Meas_V'
            x_label = 'Laser Current Set (mA)'
            y1_label = 'PD Current (mA)'
            y1_colour = 'red'
            y2_label = 'Laser Voltage (V)'
            y2_colour = 'blue'
            title_prefix = 'Laser Characterization (LIV)'
        elif '_EAM_' in filename:
            # EAM test columns
            x_col = 'SMU2_Ch1_EAM_Voltage_Set_V'
            y1_col = 'SMU1_Ch1_PD_Current_Meas_mA'
            y2_col = 'SMU2_Ch1_EAM_Current_Meas_mA'
            x_label = 'EAM Voltage Set (V)'
            y1_label = 'PD Current (mA)'
            y1_colour = 'red'
            y2_label = 'EAM Current (mA)'
            y2_colour = 'blue'
            title_prefix = 'Laser Characterization (EAM)'
        elif '_55C_80mA_' in filename: # Change name of spectrum files
            x_col = 'Freq'
            y1_col = ' Amplitude'
            y2_col = None
            x_label = 'Wavelength (nm)'
            y1_label = 'Amplitude (dBm)'
            y1_colour = 'blue'
            y2_label = None
            y2_colour = None
            title_prefix = 'Spectrum Plot'
        else:
            messagebox.showerror("Unknown Test Type",
                                 f"Could not determine test type from filename: {filename}\n"
                                 "Expected '_LIV_' or '_EAM_' in filename.")
            return

        # Clear the previous plot and make axes visible
        self.ax.clear()
        self.ax.set_visible(True)

        # Plot PD Current on the primary Y-axis
        self.ax.plot(df[x_col], df[y1_col], marker='o', linestyle='-', markersize=3, color=y1_colour,
                     label=y1_label, linewidth=2)
        self.ax.set_ylabel(y1_label, color=y1_colour, fontweight='bold')
        self.ax.tick_params(axis='y', labelcolor=y1_colour)

        if y2_col is not None:
            # Create a second Y-axis
            ax2 = self.ax.twinx()

            # Plot second parameter on the secondary Y-axis
            ax2.plot(df[x_col], df[y2_col], marker='s', linestyle='--', markersize=3, color=y2_colour,
                     label=y2_label, linewidth=2)
            ax2.set_ylabel(y2_label, color=y2_colour, fontweight='bold')
            ax2.tick_params(axis='y', labelcolor=y2_colour)

            # Combine legends from both axes
            lines, labels = self.ax.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax2.legend(lines + lines2, labels + labels2, loc='upper left', fontsize=8)

        # Set common title and X-axis label
        self.ax.set_title(f'{title_prefix}\n{os.path.basename(excel_file_path)}', fontsize=10, fontweight='bold')
        self.ax.set_xlabel(x_label, fontweight='bold')
        self.ax.grid(True, linestyle='--', alpha=0.3)

        # Adjust layout to prevent labels from being cut off
        self.fig.tight_layout(pad=3.0)

        # Refresh the canvas
        self.canvas.draw()

        self.update_status(f"📊 Plotted: {os.path.basename(excel_file_path)}", "#2e8b57")

    def clear_plot(self):
        """Clear the current plot"""
        self.ax.clear()

        # Check if there's a twin axis and clear it
        other_axes = [ax for ax in self.fig.get_axes() if ax != self.ax]
        for ax in other_axes:
            ax.remove()

        # Hide the axes again until new data is plotted
        self.ax.set_visible(False)
        self.canvas.draw()
        self.update_status("Plot cleared", "#333333")

    def find_latest_excel_file(self, path):
        """Find the most recent Excel file in the data path"""
        try:
            if not os.path.exists(path):
                return None

            excel_files = []
            for file in os.listdir(path):
                if file.endswith(('.xlsx', '.xls', '.csv')):
                    file_path = os.path.join(path, file)
                    excel_files.append((file_path, os.path.getmtime(file_path)))

            if excel_files:
                # Sort by modification time, newest first
                excel_files.sort(key=lambda x: x[1], reverse=True)
                return excel_files[0][0]  # Return the newest file path

        except Exception as e:
            print(f"Error finding latest Excel file: {e}")
        return None

    def update_status(self, message, color="#333333"):
        """Update the status display with a message and color"""
        print("Status: " + message)
        self.graph_status_var.set(message)
        self.graph_status.configure(fg=color)