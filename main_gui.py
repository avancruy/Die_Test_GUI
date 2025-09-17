import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
import matplotlib
from utils import *
from data_extraction import Extraction
from graph_panel import GraphPanel
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
        self.sync_in_progress = False  # Flag to prevent infinite sync loops

        # Graphing related variables

        self.graph_panel = GraphPanel()

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
        self.graph_panel.setup_graph_panel(right_panel)

    def setup_control_panel(self, main_panel):
        # Compact title
        control_panel_title = tk.Label(main_panel, text="Testing Configuration",
                               font=('Arial', 12, 'bold'), bg='#f5f5f5', fg='#333333')
        control_panel_title.pack(pady=(0, 10))

        # --- Compact Top Controls ---
        config_frame = ttk.LabelFrame(main_panel, text="Configuration", padding="8")
        config_frame.pack(fill='x', pady=(0, 10))

        # Device ID
        tk.Label(config_frame, text="Device:", font=('Arial', 9, 'bold')) \
            .grid(row=0, column=0, padx=(0, 5), pady=2, sticky=tk.W)
        self.device_entry = self.create_placeholder_entry(config_frame, "e.g., AA1234", width=15)
        self.device_entry.grid(row=0, column=1, padx=(0, 15), pady=2, sticky=tk.EW)

        # Temperature
        tk.Label(config_frame, text="Temp (°C):", font=('Arial', 9, 'bold')) \
            .grid(row=0, column=2, padx=(0, 5), pady=2, sticky=tk.W)
        self.temp_entry = self.create_placeholder_entry(config_frame, "55", width=8)
        self.temp_entry.grid(row=0, column=3, padx=(0, 15), pady=2, sticky=tk.W)

        # Data Path (shorter)
        tk.Label(config_frame, text="Path:", font=('Arial', 9, 'bold')).grid(row=0, column=4, padx=(0, 5), pady=2,
                                                                             sticky=tk.W)
        self.path_var = tk.StringVar(value="C:/Users/labaccount.ELPHIC/Documents/TX03_submount_xpt/")
        path_entry = ttk.Entry(config_frame, textvariable=self.path_var, width=25)
        path_entry.grid(row=0, column=5, padx=(0, 5), pady=2, sticky=tk.EW)
        browse_button = ttk.Button(config_frame, text="...", command=self.browse_path, width=3)
        browse_button.grid(row=0, column=6, padx=(0, 15), pady=2)

        # Run button - smaller
        self.run_button = tk.Button(config_frame, text="▶ Run",
                                    bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'),
                                    relief='raised', bd=2, padx=15, pady=5)
        self.run_button.config(command=self.run_test_threaded)
        self.run_button.grid(row=0, column=7, padx=10, pady=2)

        # Configure column weights
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(5, weight=2)

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
            tab = ttk.Frame(self.notebook, padding="10")
            self.notebook.add(tab, text=name)
            controller.setup_tab(tab)

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
        if device_id == "e.g., AA1234":  # Accommodate placeholder
            device_id = ""

        temperature = self.temp_entry.get()

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
        time = f"{timestamp[4:6]}/{timestamp[6:8]}/{timestamp[:4]}, {timestamp[9:11]}:{timestamp[11:13]}:{timestamp[13:15]}0"
        print(f"Running {self.curr_controller.name} test for: '{device_id}' at {time} with Temp: '{temperature}°C'")
        if self.curr_controller.name == "LIV" or self.curr_controller.name == "EAM":
            print(f"Photodetector Params: {self.curr_controller.params_photodetector}")
            print(f"Laser Params: {self.curr_controller.params_laser}")
            print(f"EAM Params: {self.curr_controller.params_eam}")
        elif self.curr_controller.name == "Spectrum":
            print(f"Spectrum Params: {self.curr_controller.params_spectrum}")

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
            if self.graph_panel.auto_plot_var.get() and self.curr_controller is not self.extraction_controller:
                latest_file = self.graph_panel.find_latest_excel_file(self.path_var.get())
                if latest_file:
                    self.graph_panel.clear_plot()
                    self.root.after(0, lambda: self.graph_panel.excel_path_var.set(latest_file))
                    self.root.after(0, self.graph_panel.plot_excel_data)
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
    from test_classes import LIV, EAM, Spectrum

    root = tk.Tk()
    app = PulsedGuiApp(root)

    from PIL import Image, ImageTk

    ico = Image.open('inpho_logo.png')
    photo = ImageTk.PhotoImage(ico)
    root.wm_iconphoto(False, photo)

    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
except ImportError:
    print("Error: 'test_classes.py' not found")

