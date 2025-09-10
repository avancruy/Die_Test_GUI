import tkinter as tk
from tkinter import ttk
import pyvisa
from utils import *
from datetime import datetime
import time

class Base:
    def __init__(self, smu_resources):
        print('Base Test Class Initialized')
        self.smu_resources1_addr = smu_resources['smu1']
        self.smu_resources2_addr = smu_resources['smu2']
        self.smu1 = None
        self.smu2 = None
        self.base_path_config = "C:/Users/labaccount.ELPHIC/Documents/TX03_submount_xpt/"

        self.name = "Base"
        self.params_photodetector = {}
        self.params_laser = {}
        self.params_eam = {}
        self.params_spectrum = {}

        self.param_sets = [
            ("Photodetector (SMU1 Ch1)", self.params_photodetector, '#e8f4fd'),
            ("Laser (SMU1 Ch2)", self.params_laser, '#fff2e8'),
            ("EAM (SMU2 Ch1)", self.params_eam, '#f0f8e8')
        ]

        self.param_vars = {}
        self.sync_in_progress = False

        self.PARAM_METADATA = {
            "smu_channel": ("SMU Channel", int, None),  # Typically fixed per setup
            "source_func": ("Source Function", str, [("Voltage(V)", "volt"), ("Current(mA)", "curr")]),
            "source_shape": ("Source Shape", str, [("DC", "dc"), ("Pulse", "puls")]),
            "source_mode": ("Source Mode", str, [("Fixed", "fix"), ("Sweep", "swe"), ("List", "list")]),
            "start": ("Start Value", float, None),
            "stop": ("Stop Value", float, None),
            "num_points": ("Number of Points", int, None),
            "initval": ("Initial/Base Value", float, None),
            "pulse_delay": ("Pulse Delay (s)", float, None),
            "pulse_width": ("Pulse Width (s)", float, None),
            "sense_func": ("Sense Function", str, [("Current(mA)", "curr"), ("Voltage(V)", "volt")]),
            "sense_range": ("Sense Range", float, None),
            "aperture": ("Aperture Time (s)", float, None),
            "protection": ("Protection/Compliance Level", float, None),
            "trigger_period": ("Trigger Period (s)", float, None),
            "trigger_transition_delay": ("Trigger Transition Delay (s)", float, None),
            "trigger_acquisition_delay": ("Trigger Acquisition Delay (s)", float, None)
        }

        self.SYNCHRONIZED_PARAMS = {
            "num_points", "trigger_period", "pulse_width",
            "trigger_transition_delay", "trigger_acquisition_delay"
        }

        # Easy duty cycle adjustment guide:
        # For 90% duty cycle: pulse_width = 36.0e-3, trigger_period = 40e-3
        # For 75% duty cycle: pulse_width = 30.0e-3, trigger_period = 40e-3
        # For 50% duty cycle: pulse_width = 20.0e-3, trigger_period = 40e-3

        self.base_path_config = "C:/Users/labaccount.ELPHIC/Documents/TX03_submount_xpt/"

    def setup_tab(self, parent):
        # Create main columns container
        print(self.name)
        print(self.param_sets)
        columns_container = tk.Frame(parent, bg='white')
        columns_container.pack(expand=True, fill="both")

        # Create three columns with individual scrollbars
        for i, (name, params_dict, bg_color) in enumerate(self.param_sets):
            # Create column frame
            column_frame = tk.Frame(columns_container, bg=bg_color, relief='solid', bd=1)
            column_frame.grid(row=0, column=i, padx=5, pady=5, sticky='nsew')

            # Section header (fixed at top)
            header_frame = tk.Frame(column_frame, bg=bg_color)
            header_frame.pack(fill=tk.X, padx=8, pady=(8, 3))

            section_label = tk.Label(header_frame, text=name, font=('Arial', 10, 'bold'),
                                     bg=bg_color, fg='#333333')
            section_label.pack(anchor='w')

            # Create scrollable area for parameters
            canvas = tk.Canvas(column_frame, bg=bg_color, highlightthickness=0)
            scrollbar = ttk.Scrollbar(column_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=bg_color)

            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Pack scrollbar and canvas
            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True, padx=8, pady=(0, 8))

            # Bind mousewheel to canvas
            def _on_mousewheel(event, canvas=canvas):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda ev: _on_mousewheel(ev, c)))
            canvas.bind("<Leave>", lambda e, c=canvas: c.unbind_all("<MouseWheel>"))

            # Create parameters in the scrollable frame
            self.create_param_entries_vertical(scrollable_frame, params_dict, name, bg_color)

        # Configure column weights for equal distribution
        for i in range(3):
            columns_container.columnconfigure(i, weight=1)
        columns_container.rowconfigure(0, weight=1)

    def create_param_entries_vertical(self, parent, params_dict_ref, param_set_name, bg_color):
        row_idx = 0

        for key, (label_text, data_type, options) in self.PARAM_METADATA.items():
            if key not in params_dict_ref:
                if key == "smu_channel":  # Show SMU channel as non-editable
                    label_frame = tk.Frame(parent, bg=bg_color)
                    label_frame.pack(fill='x', pady=1)

                    tk.Label(label_frame, text=f"{label_text}:", bg=bg_color,
                             font=('Arial', 8, 'bold')).pack(anchor='w')
                    tk.Label(label_frame, text=str(params_dict_ref.get(key, "N/A")),
                             bg=bg_color, fg='#666666', font=('Arial', 8)).pack(anchor='w')
                continue

            var_key = f"{param_set_name}_{key}"

            # Create parameter frame
            param_frame = tk.Frame(parent, bg=bg_color)
            param_frame.pack(fill='x', pady=1)

            # Parameter label with sync indicator
            label_text_display = label_text
            if key in self.SYNCHRONIZED_PARAMS:
                label_text_display += " 🔗"

            param_label = tk.Label(param_frame, text=f"{label_text_display}:", bg=bg_color,
                                   font=('Arial', 8, 'bold'))
            param_label.pack(anchor='w')

            current_value = params_dict_ref.get(key)

            # Create appropriate widget
            if data_type == bool:  # Checkbutton for boolean
                var = tk.BooleanVar(value=bool(current_value))
                widget = ttk.Checkbutton(param_frame, variable=var)
            elif options:  # Combobox for predefined options
                var = tk.StringVar(value=str(current_value))
                display_value = str(current_value)
                for opt_disp, opt_val in options:
                    if opt_val == current_value:
                        display_value = opt_disp
                        break
                var.set(display_value)
                widget = ttk.Combobox(param_frame, textvariable=var,
                                      values=[opt[0] for opt in options],
                                      state="readonly")
            else:  # Entry for other types
                var = tk.StringVar(value=str(current_value))
                widget = ttk.Entry(param_frame, textvariable=var)

            widget.pack(fill='x', pady=(1, 0))

            # Store var and add trace
            self.param_vars[var_key] = var
            var.trace_add("write", lambda *args, k=key, dt=data_type, ps_dict=params_dict_ref,
                                          v=var, opts=options, psn=param_set_name:
                self.update_parameter(k, dt, ps_dict, v, opts, psn))

            row_idx += 1

    def update_parameter(self, param_key, param_type, param_set_dict, tk_var, options=None, param_set_name=None):
        if self.sync_in_progress:
            return

        new_value_str = tk_var.get()

        if options:  # Combobox: convert display name back to script value
            actual_value = None
            for display_name, script_value in options:
                if display_name == new_value_str:
                    actual_value = script_value
                    break
            if actual_value is None:
                print(f"Error: Could not find script value for combobox display '{new_value_str}'")
                return
            new_value = actual_value
        elif param_type == bool:
            new_value = tk_var.get()
        else:  # Entry widget
            # Here we check for the case where the entry is blank
            if new_value_str == "":  # Empty string should remain empty, not converted
                new_value = None  # Or you could set this to '' depending on your use case
            else:
                new_value = string_to_num(new_value_str, param_type)
                if new_value is None and new_value_str != "":
                    print(
                        f"Warning: Invalid input '{new_value_str}' for {param_key} (expected {param_type}). Not updated.")
                    return

        # Update only if the new value is not None or the string is intentionally empty
        if new_value is not None or (param_type != bool and new_value_str == ""):
            if param_type == bool:
                if new_value_str == "":
                    new_value = False  # Defaulting to False if it's a boolean and empty
            if param_set_dict.get(param_key) != new_value:
                param_set_dict[param_key] = new_value
                print(f"Updated {param_key} in {param_set_dict.get('smu_channel', 'Unknown SMU')}: {new_value}")

                # Synchronize parameter across all devices if it's a synchronized parameter
                if param_key in self.SYNCHRONIZED_PARAMS:
                    self.synchronize_parameter(param_key, new_value, param_set_name)

    def synchronize_parameter(self, param_key, new_value, source_param_set_name):
        """Synchronize a parameter value across all parameter sets"""
        if self.sync_in_progress:
            return

        self.sync_in_progress = True

        try:
            param_sets = [
                ("Photodetector (SMU1 Ch1)", self.params_photodetector),
                ("Laser (SMU1 Ch2)", self.params_laser),
                ("EAM (SMU2 Ch1)", self.params_eam)
            ]

            for param_set_name, params_dict in param_sets:
                if param_set_name == source_param_set_name:
                    continue  # Skip the source parameter set

                if param_key in params_dict:
                    # Update the parameter dictionary
                    params_dict[param_key] = new_value

                    # Update the GUI variable
                    var_key = f"{param_set_name}_{param_key}"
                    if var_key in self.param_vars:
                        self.param_vars[var_key].set(str(new_value))

            print(f"Synchronized {param_key} = {new_value} across all devices")

        finally:
            self.sync_in_progress = False

    def connect_smus(self):
        try:
            if self.smu1 is None or self.smu1.instrument is None:
                print(f"Attempting to connect to SMU1: {self.smu_resources1_addr}")
                self.smu1 = KeysightB2912A(self.smu_resources1_addr)
            if self.smu2 is None or self.smu2.instrument is None:
                print(f"Attempting to connect to SMU2: {self.smu_resources2_addr}")
                self.smu2 = KeysightB2912A(self.smu_resources2_addr)

            if self.smu1.instrument is None or self.smu2.instrument is None:
                raise ConnectionError("One or both SMUs failed to connect.")
            return True
        except Exception as e:
            print(f"Error connecting SMUs: {e}")
            self.smu1 = None
            self.smu2 = None
            return False

    def run_test(self, device_id="", temperature="", timestamp=""):
        settle_time = 0.5

        if not self.connect_smus():
            print("SMU connection failed. Aborting test.")
            return

        try:
            print("\n--- Initializing SMU1 ---")
            self.smu1.reset()
            print("\n--- Initializing SMU2 ---")
            self.smu2.reset()
            time.sleep(settle_time)

            current_timestamp = timestamp or datetime.now().strftime("%Y%m%dT%H%M%S")

            self.smu1.config_pulsed_params(self.params_photodetector)
            self.smu1.config_pulsed_params(self.params_laser)
            self.smu2.config_pulsed_params(self.params_eam)

            if self.params_laser['source_mode'].lower() == 'swe' and self.params_laser['start'] != self.params_laser[
                'stop']:
                test_checker = True
                print("Running LIV Test (Laser sweep)")
                num_measurement_points = self.params_laser['num_points']
                active_trigger_period = self.params_laser['trigger_period']
            elif self.params_eam['source_mode'].lower() == 'swe' and self.params_eam['start'] != self.params_eam[
                'stop']:
                test_checker = False
                print("Running EAM Test (EAM sweep)")
                num_measurement_points = self.params_eam['num_points']
                active_trigger_period = self.params_eam['trigger_period']
            else:
                print("Warning: Neither Laser nor EAM is in sweep mode. Defaulting to EAM num_points for timing.")
                test_checker = False
                num_measurement_points = self.params_eam['num_points']
                active_trigger_period = self.params_eam['trigger_period']

            print("\nTurning on outputs and initiating measurement...")
            self.smu1.output_on(self.params_photodetector['smu_channel'])
            self.smu1.output_on(self.params_laser['smu_channel'])
            self.smu2.output_on(self.params_eam['smu_channel'])
            time.sleep(0.1)

            self.smu1.write(f":init (@{self.params_photodetector['smu_channel']},{self.params_laser['smu_channel']})")
            self.smu2.write(f":init (@{self.params_eam['smu_channel']})")
            print("Measurement initiated. Waiting for completion...")

            total_meas_time = num_measurement_points * active_trigger_period + 2
            print(f"Estimated measurement time: {total_meas_time:.2f} seconds for {num_measurement_points} points.")

            for i in range(int(total_meas_time)):
                print(f"Waiting... {i + 1}/{int(total_meas_time)}s", end='\r')
                time.sleep(1)
            print("\nMeasurement potentially complete. Fetching data.")

            print("\nFetching measurement results...")
            laser_voltage_data = self.smu1.query(f":fetc:arr:volt? (@{self.params_laser['smu_channel']})")
            photodetector_current_data = self.smu1.query(
                f":fetc:arr:curr? (@{self.params_photodetector['smu_channel']})")
            eam_current_data = self.smu2.query(f":fetc:arr:curr? (@{self.params_eam['smu_channel']})")

            print(f"  Laser Fetched (V): {laser_voltage_data[:50]}...")
            print(f"  Detector Fetched (I): {photodetector_current_data[:50]}...")
            print(f"  EAM Fetched (I): {eam_current_data[:50]}...")

            success = create_combined_excel_file(
                laser_voltage_data,
                photodetector_current_data,
                eam_current_data,
                current_timestamp,
                self.params_photodetector,
                self.params_laser,
                self.params_eam,
                test_checker,
                device_id,
                temperature,
                self.base_path_config
            )

            if success:
                print("\n--- Excel file creation completed successfully ---")
            else:
                print("\n--- Excel file creation failed ---")

            print("\n--- Measurement Sequence Finished ---")

        except ConnectionError as e:
            print(f"Connection Error during test: {e}")
        except pyvisa.errors.VisaIOError as e:
            print(f"A VISA communication error occurred: {e}")
        except Exception as e:
            print(f"An unexpected error occurred during the script execution: {e}")
            import traceback
            traceback.print_exc()
        finally:
            # ✨ Modified cleanup procedure
            print("\n--- Cleaning Up ---")
            if self.smu1 and self.smu1.instrument:
                # Turn off outputs first
                self.smu1.output_off(self.params_photodetector['smu_channel'])
                self.smu1.output_off(self.params_laser['smu_channel'])
                print("SMU1 outputs off.")

                # Set channel 1 (PD) to -1V bias with 50mA compliance
                self.smu1.set_source_mode(1, 'VOLT')
                self.smu1.set_voltage(1, -1.0)
                self.smu1.set_current_compliance(1, 0.05)  # 50mA in Amps
                print("SMU1 Channel 1 set to -1V bias with 50mA compliance.")

                # Set channel 2 (Laser) to 80mA bias with 2V compliance
                self.smu1.set_source_mode(2, 'CURR')
                self.smu1.set_current(2, 0.08)  # 80mA in Amps
                self.smu1.set_voltage_compliance(2, 2.0)  # 2V compliance
                print("SMU1 Channel 2 set to 80mA bias with 2V compliance.")

            if self.smu2 and self.smu2.instrument:
                self.smu2.output_off(self.params_eam['smu_channel'])
                print("SMU2 outputs off.")

                # Set channel 3 (EAM) to -2. bias with 80mA compliance
                self.smu2.set_source_mode(1, 'VOLT')
                self.smu2.set_voltage(1, -2.0)
                self.smu2.set_current_compliance(1, 0.08)  # 80mA in Amps
                print("SMU2 Channel 1 set to -2V bias with 80mA compliance.")

    def close_smus(self):
        if self.smu1:
            self.smu1.close()
            self.smu1 = None
        if self.smu2:
            self.smu2.close()
            self.smu2 = None
        print("SMU connections closed.")

class EAM(Base):
    def __init__(self, smu_resources):
        super().__init__(smu_resources)

        self.name = "EAM"

        # Default parameters - these will be updated by the GUI
        self.params_photodetector = {  # SMU1 chan1
            "smu_channel": 1, "source_func": "volt", "source_shape": "dc", "source_mode": "swe",
            "start": -1.0, "stop": -1.0, "num_points": 32, "initval": -1.0,
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "curr", "sense_range": 100, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 50, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.params_laser = {  # SMU1 chan2
            "smu_channel": 2, "source_func": "curr", "source_shape": "puls", "source_mode": "swe",
            "start": 80, "stop": 80, "num_points": 32, "initval": 80,
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "volt", "sense_range": 2.0, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 2.0, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.params_eam = {  # SMU2 chan1
            "smu_channel": 1, "source_func": "volt", "source_shape": "dc", "source_mode": "swe",
            "start": -2.5833, "stop": 0, "num_points": 32, "initval": -2.5833,  # Updated from 0.0 to -1.0
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "curr", "sense_range": 100, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 80, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.param_sets = [
            ("Photodetector (SMU1 Ch1)", self.params_photodetector, '#e8f4fd'),
            ("Laser (SMU1 Ch2)", self.params_laser, '#fff2e8'),
            ("EAM (SMU2 Ch1)", self.params_eam, '#f0f8e8')
        ]

class LIV(Base):
    def __init__(self, smu_resources):
        super().__init__(smu_resources)

        self.name = "LIV"
        # Default parameters - these will be updated by the GUI
        self.params_photodetector = {  # SMU1 chan1
            "smu_channel": 1, "source_func": "volt", "source_shape": "dc", "source_mode": "swe",
            "start": -1.0, "stop": -1.0, "num_points": 21, "initval": -1.0,
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "curr", "sense_range": 100, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 50, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.params_laser = {  # SMU1 chan2
            "smu_channel": 2, "source_func": "curr", "source_shape": "puls", "source_mode": "swe",
            "start": 0, "stop": 100, "num_points": 21, "initval": 0,
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "volt", "sense_range": 2.0, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 2.0, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.params_eam = {  # SMU2 chan1
            "smu_channel": 1, "source_func": "volt", "source_shape": "dc", "source_mode": "fix",
            "start": 0, "stop": 0, "num_points": 21, "initval": 0,  # Updated from 0.0 to -1.0
            "pulse_delay": 0.5e-3, "pulse_width": 200.0e-3,  # 50% duty cycle
            "sense_func": "curr", "sense_range": 100, "aperture": 5e-3,  # Updated aperture from 0.5e-3 to 5e-3
            "protection": 80, "trigger_period": 400e-3,  # 400ms period
            "trigger_transition_delay": 1.5e-3, "trigger_acquisition_delay": 2.9e-3
        }

        self.param_sets = [
            ("Photodetector (SMU1 Ch1)", self.params_photodetector, '#e8f4fd'),
            ("Laser (SMU1 Ch2)", self.params_laser, '#fff2e8'),
            ("EAM (SMU2 Ch1)", self.params_eam, '#f0f8e8')
        ]

class Spectrum(Base):
    def __init__(self, smu_resources):
        super().__init__(smu_resources)

        self.name = "Spectrum"

        self.PARAM_METADATA = {
            "smu_channel": ("SMU Channel", int, None),  # Typically fixed per setup
            "source_func": ("Source Function", str, [("Voltage(V)", "volt"), ("Current(mA)", "curr")]),
            "source_shape": ("Source Shape", str, [("DC", "dc"), ("Pulse", "puls")]),
            "source_mode": ("Source Mode", str, [("Fixed", "fix"), ("Sweep", "swe"), ("List", "list")]),
            "start": ("Start Value", float, None),
            "stop": ("Stop Value", float, None),
            "num_points": ("Number of Points", int, None),
            "initval": ("Initial/Base Value", float, None),
            "centre": ("Centre", float, None),
            "span": ("Span", float, None),
            "res": ("Resolution", float, None),
            "sens": ("Sensitivity", str, None),
            "avg": ("Average", float, None),
            "ref_val": ("Reference Value", float, None)
        }
        self.params_spectrum = {
            "smu_channel": 1, "source_func": "volt", "source_shape": "dc", "source_mode": "swe",
            "start": 0, "stop": 0, "num_points": 21, "initval": 0,
            "centre": 1310, "span": 10, "res": 0.02, "sens": 'High1', "avg": 0, "ref_val": 0
        }

        self.param_sets = [
            ("Spectrum", self.params_spectrum, '#e8f4fd'),
        ]