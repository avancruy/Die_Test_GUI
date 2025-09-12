import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib
from utils import *
from data_extraction import Extraction
"""
Ari Van Cruyningen, Matthew Manjaly
"""
matplotlib.use('TkAgg')  # Use TkAgg backend for tkinter integration

class PulsedGuiApp:
    def __init__(self, root):
        self.root = root
        self.root.title('LIV/EAM/Spectrum Testing with Real-time Graphing')
        self.root.configure(bg='#f5f5f5')

        # Set minimum window size - even wider to accommodate graph
        self.root.minsize(1800, 800)

        smu_resources = {
            'smu1': 'TCPIP0::10.20.0.231::hislip0::INSTR',  # Replace with actual VISA address
            'smu2': 'TCPIP0::10.20.0.230::inst0::INSTR'  # Replace with actual VISA address
        }
        self.liv_controller = LIV(smu_resources)
        self.eam_controller = EAM(smu_resources)
        self.spectrum_controller = Spectrum(smu_resources)
        self.extraction_controller = Extraction(smu_resources)
        self.curr_controller = self.liv_controller
        #self.param_vars = {}  # To store Tkinter variables for parameters
        self.sync_in_progress = False  # Flag to prevent infinite sync loops

        # Graphing related variables
        self.current_excel_file = None
        self.fig = None
        self.canvas = None
        self.toolbar = None

        self.setup_gui()

    def on_focus_in(self, event, placeholder):
        if event.widget.get() == placeholder:
            event.widget.delete(0, tk.END)
            event.widget.config(foreground='black')

    def on_focus_out(self, event, placeholder):
        if not event.widget.get():
            event.widget.insert(0, placeholder)
            event.widget.config(foreground='grey')

    def create_placeholder_entry(self, parent, placeholder, **kwargs):
        entry = ttk.Entry(parent, foreground='grey', **kwargs)
        entry.insert(0, placeholder)
        entry.bind("<FocusIn>", lambda e, p=placeholder: self.on_focus_in(e, p))
        entry.bind("<FocusOut>", lambda e, p=placeholder: self.on_focus_out(e, p))
        return entry

    def browse_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self.path_var.set(directory.replace("\\", "/") + "/")
            self.liv_controller.path = directory
            self.eam_controller.path = directory
            self.spectrum_controller.path = directory
            self.extraction_controller.path = directory


    def setup_gui(self):
        # Create main horizontal paned window
        main_paned = ttk.PanedWindow(self.root, orient='horizontal')
        main_paned.pack(expand=True, fill='both', padx=10, pady=10)

        # Left panel for controls
        left_panel = tk.Frame(main_paned, bg='#f5f5f5')
        # Right panel for graphing
        right_panel = tk.Frame(main_paned, bg='#f5f5f5')

        main_paned.add(left_panel, weight=1)
        main_paned.add(right_panel, weight=3)

        self.setup_control_panel(left_panel)
        self.setup_graph_panel(right_panel)

    def setup_control_panel(self, main_panel):
        # Compact title
        title_label = tk.Label(main_panel, text="Pulsed LIV - EAM Control",
                               font=('Arial', 12, 'bold'), bg='#f5f5f5', fg='#333333')
        title_label.pack(pady=(0, 10))

        # --- Compact Top Controls ---
        controls_frame = ttk.LabelFrame(main_panel, text="Configuration", padding="8")
        controls_frame.pack(fill='x', pady=(0, 10))

        # Single row layout for all controls
        control_grid = tk.Frame(controls_frame)
        control_grid.pack(fill='x')

        # Device ID
        tk.Label(control_grid, text="Device:", font=('Arial', 9, 'bold')) \
            .grid(row=0, column=0, padx=(0, 5), pady=2, sticky=tk.W)
        self.device_entry = self.create_placeholder_entry(control_grid, "W0xx_GF_0102", width=15)
        self.device_entry.grid(row=0, column=1, padx=(0, 15), pady=2, sticky=tk.EW)

        # Temperature
        tk.Label(control_grid, text="Temp (°C):", font=('Arial', 9, 'bold')) \
            .grid(row=0, column=2, padx=(0, 5), pady=2, sticky=tk.W)
        self.temp_entry = self.create_placeholder_entry(control_grid, "25", width=8)
        self.temp_entry.grid(row=0, column=3, padx=(0, 15), pady=2, sticky=tk.W)

        # Data Path (shorter)
        tk.Label(control_grid, text="Path:", font=('Arial', 9, 'bold')).grid(row=0, column=4, padx=(0, 5), pady=2,
                                                                             sticky=tk.W)
        self.path_var = tk.StringVar(value="C:/Users/labaccount.ELPHIC/Documents/TX03_submount_xpt/")
        path_entry = ttk.Entry(control_grid, textvariable=self.path_var, width=25)
        path_entry.grid(row=0, column=5, padx=(0, 5), pady=2, sticky=tk.EW)
        browse_button = ttk.Button(control_grid, text="...", command=self.browse_path, width=3)
        browse_button.grid(row=0, column=6, padx=(0, 15), pady=2)

        # Run button - smaller
        self.run_button = tk.Button(control_grid, text="▶ Run",
                                    bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'),
                                    relief='raised', bd=2, padx=15, pady=5)
        self.run_button.config(command=self.run_test_threaded)
        self.run_button.grid(row=0, column=7, padx=10, pady=2)

        # Configure column weights
        control_grid.columnconfigure(1, weight=1)
        control_grid.columnconfigure(5, weight=2)

        self.path_var.trace_add("write",
                                lambda *args: setattr(self.extraction_controller, 'path', self.path_var.get()))

        # --- Compact Status Display ---
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = tk.Label(main_panel, textvariable=self.status_var,
                                     font=('Arial', 10, 'bold'), fg='#333333', bg='#f5f5f5')
        self.status_label.pack(pady=(0, 8))

        # --- Notebook for Parameters ---
        self.notebook = ttk.Notebook(main_panel)
        self.notebook.pack(expand=True, fill="both", pady=10)

        tab_config = [
            ("LIV", self.liv_controller),
            ("EAM", self.eam_controller),
            ("Spectrum", self.spectrum_controller),
            ("Data Extraction", self.extraction_controller)
        ]

        for name, controller in tab_config:
            tab = ttk.Frame(self.notebook, padding="50")
            self.notebook.add(tab, text=name)
            #self.create_param_entries(tab, params_dict, name)  # Pass name for unique var keys
            #self.create_column_layout(tab)
            controller.setup_tab(tab)

    def setup_graph_panel(self, main_panel):
        # Graph panel title
        graph_title = tk.Label(main_panel, text="Real-time Data Visualization",
                               font=('Arial', 11, 'bold'), bg='#f5f5f5', fg='#333333')
        graph_title.pack(pady=(0, 10))

        # --- Graph Controls ---
        graph_controls_frame = ttk.LabelFrame(main_panel, text="Graph Controls", padding="8")
        graph_controls_frame.pack(fill='x', pady=(0, 10))

        # Compact controls layout
        controls_grid = tk.Frame(graph_controls_frame)
        controls_grid.pack(fill='x')

        # Excel file selection (smaller)
        tk.Label(controls_grid, text="Excel:", font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=(0, 5),
                                                                               sticky=tk.W)
        self.excel_path_var = tk.StringVar()
        excel_entry = ttk.Entry(controls_grid, textvariable=self.excel_path_var, width=20)
        excel_entry.grid(row=0, column=1, padx=(0, 5), sticky=tk.EW)

        browse_excel_button = ttk.Button(controls_grid, text="📁", command=self.browse_excel_file, width=3)
        browse_excel_button.grid(row=0, column=2, padx=(0, 10))

        # Plot and clear buttons
        plot_button = tk.Button(controls_grid, text="📊 Plot", command=self.plot_excel_data,
                                bg='#2196F3', fg='white', font=('Arial', 8, 'bold'), relief='raised', bd=1)
        plot_button.grid(row=0, column=3, padx=(0, 5))

        clear_button = tk.Button(controls_grid, text="🗑 Clear", command=self.clear_plot,
                                 bg='#FF5722', fg='white', font=('Arial', 8, 'bold'), relief='raised', bd=1)
        clear_button.grid(row=0, column=4, padx=(0, 10))

        # Auto-plot checkbox
        self.auto_plot_var = tk.BooleanVar(value=True)
        auto_plot_check = ttk.Checkbutton(controls_grid, text="Auto-plot after test", variable=self.auto_plot_var)
        auto_plot_check.grid(row=0, column=5, padx=5, sticky=tk.W)

        controls_grid.columnconfigure(1, weight=1)

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
            print(df[' Amplitude'])
            print("Done")
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

    def find_latest_excel_file(self):
        """Find the most recent Excel file in the data path"""
        try:
            data_path = self.path_var.get()
            if not os.path.exists(data_path):
                return None

            excel_files = []
            for file in os.listdir(data_path):
                if file.endswith(('.xlsx', '.xls', '.csv')):
                    file_path = os.path.join(data_path, file)
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
        self.status_var.set(message)
        self.status_label.configure(fg=color)

    def run_test_threaded(self):
        device_id = None
        temperature = None
        timestamp = None

        current_tab = self.notebook.index('current')
        self.run_button.config(state='disabled', text="⏳ Running...", bg='#ff9800')

        if current_tab == 0:     # LIV
            self.curr_controller = self.liv_controller
        elif current_tab == 1:   # EAM
            self.curr_controller = self.eam_controller
        elif current_tab == 2:   # Spectrum
            self.curr_controller = self.spectrum_controller
        else:                           # Data Extraction
            # Skip all initialization
            self.curr_controller = self.extraction_controller
            test_thread = threading.Thread(target=self.execute_test_and_reenable_button,
                                           args=(device_id, temperature, timestamp))
            test_thread.daemon = True
            test_thread.start()
            return

        self.update_status(f"Running {self.curr_controller.name} test...", "#4682b4")
        device_id = self.device_entry.get()
        if device_id == "e.g., W0xx_GF_0102" or device_id == "W0xx_GF_0102":  # Accommodate placeholder
            device_id = ""

        temperature = self.temp_entry.get()
        if temperature == "e.g., 25" or temperature == "25":  # Accommodate placeholder
            temperature = ""

        # Validate temperature if entered
        if temperature:
            temp_val = string_to_num(temperature, float)
            if temp_val is None:
                messagebox.showerror("Input Error", f"Invalid temperature value: {temperature}. Please enter a number.")
                self.run_button.config(state='normal', text="▶ Run", bg='#4CAF50')
                self.update_status("Ready")
                return
            temperature = str(temp_val)

        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        print(f"Running {self.curr_controller.name} test for: '{device_id}' at {timestamp} with Temp: '{temperature}°C'")
        if self.curr_controller.name == "LIV" or self.curr_controller.name == "EAM":
            print(f"Photodetector Params: {self.curr_controller.params_photodetector}")
            print(f"Laser Params: {self.curr_controller.params_laser}")
            print(f"EAM Params: {self.curr_controller.params_eam}")
        elif self.curr_controller.name == "Spectrum":
            print(f"Spectrum Params: {self.curr_controller.param_spectrum}")

        print(f"Data Path: {self.path_var.get()}")

        # Run the test in a separate thread
        test_thread = threading.Thread(target=self.execute_test_and_reenable_button,
                                       args=(device_id, temperature, timestamp))
        test_thread.daemon = True
        test_thread.start()

    def execute_test_and_reenable_button(self, device_id, temperature, timestamp):
        try:
            self.curr_controller.run_test(data_path=self.path_var.get(), device_id=device_id, temperature=temperature, timestamp=timestamp)
            print("Test execution finished.")
            self.root.after(0, lambda: self.update_status("Test completed successfully ✓", "#2e8b57"))

            # Auto-plot if enabled
            if self.auto_plot_var.get() and self.curr_controller is not self.extraction_controller:
                latest_file = self.find_latest_excel_file()
                if latest_file:
                    self.root.after(0, lambda: self.excel_path_var.set(latest_file))
                    self.root.after(0, self.plot_excel_data)
                else:
                    self.root.after(0, lambda: self.update_status("Test complete. No Excel file found to auto-plot.",
                                                                  "#ff8c00"))

        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Error during test execution: {e}")
            self.root.after(0, lambda: self.update_status("Test failed ✗", "#dc143c"))
            self.root.after(0, lambda: messagebox.showerror("Test Error", f"An error occurred during the test:\n\n{e}"))
        finally:
            self.root.after(0, lambda: self.run_button.config(state='normal', text="▶ Run", bg='#4CAF50'))

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit? This will close SMU connections if active."):
            print("Closing application, ensuring SMUs are closed...")
            self.curr_controller.close_smus()
            self.root.destroy()

# main----------------------------------------------------

try:
    from pulsed_classes import LIV, EAM, Spectrum
except ImportError:
    print("Error: 'pulsed_classes.py' not found")
    exit()

root = tk.Tk()
app = PulsedGuiApp(root)
root.protocol("WM_DELETE_WINDOW", app.on_closing)
root.mainloop()