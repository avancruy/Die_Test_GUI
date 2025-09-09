import pyvisa
import pandas as pd
import numpy as np

def string_to_num(s, target_type=float):
    try:
        return target_type(s)
    except ValueError:
        try:  # Retry with float for potential scientific notation then convert
            val = float(s)
            if target_type == int:
                return int(val)
            return val
        except ValueError:
            return None  # Indicates conversion failure

# --- KeysightB2912A Class Definition (Unchanged) ---
class KeysightB2912A:
    def __init__(self, resource_name):
        self.rm = pyvisa.ResourceManager()
        self.instrument = None
        try:
            self.instrument = self.rm.open_resource(resource_name)
            self.instrument.timeout = 10000  # Increased timeout
            self.instrument.write_termination = '\n'
            self.instrument.read_termination = '\n'
            print(f"Connected to: {self.query('*IDN?')}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Failed to connect to instrument at {resource_name}: {e}")
            self.instrument = None
            self.rm = None

    def query(self, command):
        if self.instrument:
            try:
                response = self.instrument.query(command)
                self._check_instrument_error()
                return response
            except pyvisa.errors.VisaIOError as e:
                print(f"VISA Error during query '{command}': {e}")
                return ""
        return ""

    def write(self, command):
        if self.instrument:
            try:
                self.instrument.write(command)
                self._check_instrument_error()
            except pyvisa.errors.VisaIOError as e:
                print(f"VISA Error during write '{command}': {e}")

    def read(self):
        if self.instrument:
            try:
                response = self.instrument.read()
                return response
            except pyvisa.errors.VisaIOError as e:
                print(f"VISA Error during read: {e}")
                return ""
        return ""

    def _check_instrument_error(self):
        try:
            error_response = self.instrument.query("SYST:ERR?").strip()
            if not error_response.startswith('+0,'):  # Check if error code is not 0
                print(f"Instrument Error: {error_response}")
        except Exception as e:
            print(f"Error checking instrument error: {e}")

    def reset(self):
        self.write('*RST')
        # Using a more descriptive print statement for clarity
        print("Instrument reset to default settings.")

    def set_source_mode(self, channel, mode):
        if mode.upper() in ['VOLT', 'CURR']:
            self.write(f'SOUR{channel}:FUNC:MODE {mode.upper()}')
        else:
            print(f"Invalid source mode: {mode}. Use 'VOLT' or 'CURR'.")

    def set_voltage(self, channel, voltage):
        self.write(f':SOUR{channel}:VOLT {voltage}')

    def set_current(self, channel, current):
        self.write(f':SOUR{channel}:CURR {current}')

    def set_voltage_compliance(self, channel, compliance_voltage):
        self.write(f':SENS{channel}:VOLT:PROT:LEV {compliance_voltage}')

    def set_current_compliance(self, channel, compliance_current):
        self.write(f':SENS{channel}:CURR:PROT:LEV {compliance_current}')

    def read_voltage(self, channel):
        try:
            voltage_str = self.query(f':MEAS:VOLT? (@{channel})')
            return float(voltage_str)
        except (ValueError, TypeError):
            print(f"Could not convert voltage reading to float for channel {channel}: '{voltage_str}'")
            return None
        except pyvisa.errors.VisaIOError:
            return None

    def read_current(self, channel):
        try:
            current_str = self.query(f':MEAS:CURR? (@{channel})')
            return float(current_str)
        except (ValueError, TypeError):
            print(f"Could not convert current reading to float for channel {channel}: '{current_str}'")
            return None
        except pyvisa.errors.VisaIOError:
            return None

    def output_on(self, channel):
        self.write(f':OUTP{channel} ON')

    def output_off(self, channel):
        self.write(f':OUTP{channel} OFF')

    def set_nplc(self, channel, nplc_value):
        self.write(f':SENS{channel}:VOLT:NPLC {nplc_value}')
        self.write(f':SENS{channel}:CURR:NPLC {nplc_value}')

    def config_pulsed_params(self, params):
        channel = params['smu_channel']
        print(f"\nConfiguring SMU Channel {channel} with params: {params}")

        # --- CONVERSION LOGIC ---
        # Assume values are mA if function is 'curr' and convert to Amps for the SMU
        is_current_mode = params['source_func'].lower() == 'curr'
        start_val = params['start'] / 1000.0 if is_current_mode else params['start']
        stop_val = params['stop'] / 1000.0 if is_current_mode else params['stop']
        init_val = params['initval'] / 1000.0 if is_current_mode else params['initval']

        is_sense_current = 'curr' in params['sense_func'].lower()
        sense_range_val = params['sense_range'] / 1000.0 if is_sense_current else params['sense_range']
        protection_val = params['protection'] / 1000.0 if is_sense_current else params['protection']
        # --- END CONVERSION LOGIC ---

        self.write(f":sour{channel}:func:mode {params['source_func']}")
        self.write(f":sour{channel}:func:shap {params['source_shape']}")
        self.write(f":sour{channel}:{params['source_func']}:mode {params['source_mode']}")

        # Use the converted values
        self.write(f":sour{channel}:{params['source_func']}:star {start_val}")
        self.write(f":sour{channel}:{params['source_func']}:stop {stop_val}")
        self.write(f":sour{channel}:{params['source_func']}:poin {params['num_points']}")

        if params['source_shape'].lower() == "puls":
            self.write(f":sour{channel}:puls:del {params['pulse_delay']}")
            self.write(f":sour{channel}:puls:widt {params['pulse_width']}")
            self.write(f":sour{channel}:{params['source_func']} {init_val}")  # Base value
        elif params['source_mode'].lower() == "fix":
            self.write(f":sour{channel}:{params['source_func']} {init_val}")  # Fixed value

        print(f"Source configured for channel {channel}.")

        self.write(f":sens{channel}:func \"{params['sense_func']}\"")
        self.write(f":sens{channel}:{params['sense_func']}:rang:auto off")

        # Use the converted values for sense settings
        self.write(f":sens{channel}:{params['sense_func']}:rang {sense_range_val}")
        self.write(f":sens{channel}:{params['sense_func']}:aper {params['aperture']}")
        self.write(f":sens{channel}:{params['sense_func']}:prot:lev {protection_val}")

        print(f"Sense configured for channel {channel}.")

        self.write(f":trig{channel}:tran:del {params['trigger_transition_delay']}")  # Use from params
        self.write(f":trig{channel}:acq:del {params['trigger_acquisition_delay']}")  # Use from params
        print(f"Trigger timing adjusted for channel {channel}.")

        self.write(f":trig{channel}:sour tim")
        self.write(f":trig{channel}:tim {params['trigger_period']}")
        self.write(f":trig{channel}:coun {params['num_points']}")
        print(f"Triggers configured for channel {channel}.")

    def close(self):
        if self.instrument:
            print("Closing instrument connection.")
            self.instrument.close()
        if self.rm:
            self.rm.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# --- Utility Functions (RESTORED TO ORIGINAL) ---
def parse_measurement_data(data_string):
    try:
        if isinstance(data_string, (list, tuple)):  # Already parsed
            return [float(x) for x in data_string]
        if not data_string: return []
        values = [float(x.strip()) for x in data_string.strip().split(',')]
        return values
    except ValueError as e:
        print(f"Error parsing measurement data string '{data_string}': {e}")
        return []
    except TypeError:
        print(f"Invalid data type for parsing: {data_string}")
        return []


# ✨ Modified function to multiply PD current by 1.69
def create_combined_excel_file(laser_data, detector_data, eam_data, timestamp,
                               detector_params, laser_params, eam_params,
                               test_checker, device_id, temperature, base_path_config):
    try:
        laser_voltages_fetched = parse_measurement_data(laser_data)
        detector_currents_fetched = parse_measurement_data(detector_data)
        eam_currents_fetched = parse_measurement_data(eam_data)

        # Determine num_points from the parameter set that is performing a sweep
        if laser_params['source_mode'].lower() == 'swe':
            num_points_sweep = laser_params['num_points']
        elif eam_params['source_mode'].lower() == 'swe':
            num_points_sweep = eam_params['num_points']
        else:  # Default or if both are fixed (though one should be sweep for a typical test)
            num_points_sweep = max(laser_params['num_points'], eam_params['num_points'], detector_params['num_points'],
                                   1)

        laser_current_setpoints = np.linspace(laser_params['start'], laser_params['stop'], num_points_sweep)
        eam_voltage_setpoints = np.linspace(eam_params['start'], eam_params['stop'], num_points_sweep)

        if detector_params['source_mode'].lower() == 'fix':
            detector_voltage_setpoints = np.full(num_points_sweep, detector_params['initval'])
        else:
            detector_voltage_setpoints = np.linspace(detector_params['start'], detector_params['stop'],
                                                     num_points_sweep)

        def pad_or_truncate(data_list, length):
            if not isinstance(data_list, list): data_list = []
            if len(data_list) < length:
                return data_list + [np.nan] * (length - len(data_list))
            return data_list[:length]

        laser_voltages_final = pad_or_truncate(laser_voltages_fetched, num_points_sweep)
        detector_currents_final = pad_or_truncate(detector_currents_fetched, num_points_sweep)
        eam_currents_final = pad_or_truncate(eam_currents_fetched, num_points_sweep)

        laser_currents_mA = laser_current_setpoints
        # ✨ Apply 1.69 multiplication factor to PD current readings
        detector_currents_mA = np.array(detector_currents_final) * 1000
        eam_currents_mA = np.array(eam_currents_final) * 1000

        combined_data = {
            'SMU2_Ch1_EAM_Voltage_Set_V': eam_voltage_setpoints,
            'SMU2_Ch1_EAM_Current_Meas_mA': eam_currents_mA,
            'SMU1_Ch2_Laser_Current_Set_mA': laser_currents_mA,
            'SMU1_Ch2_Laser_Voltage_Meas_V': laser_voltages_final,
            'SMU1_Ch1_PD_Voltage_Set_V': detector_voltage_setpoints,
            'SMU1_Ch1_PD_Current_Meas_mA': np.abs(detector_currents_mA),
        }

        combined_df = pd.DataFrame(combined_data)

        pulse_width_s = laser_params.get('pulse_width', 0)
        period_s = laser_params.get('trigger_period', 0)
        duty_cycle = 100

        if (laser_params['source_shape'] == 'puls'):
            duty_cycle = (pulse_width_s / period_s) * 100 if period_s > 0 else 0

        # Determine if laser is in pulsed mode (duty cycle < 100% or source_shape is 'puls' with DC < 100%)
        is_pulsed = (laser_params.get('source_shape', '').lower() == 'puls' and duty_cycle < 100)

        base_path = base_path_config
        file_prefix = f"{device_id}_" if device_id else ""
        temp_suffix = f"{temperature}" if temperature else ""

        common_suffix = f"NumPoints{num_points_sweep}_DtyC{duty_cycle:.2f}%_{temp_suffix}°C_{timestamp}.xlsx"

        if test_checker:
            ld_start_mA = int(laser_params['start'])
            ld_stop_mA = int(laser_params['stop'])
            eam_bias_V = eam_params['initval']
            pd_bias_V = detector_params['initval']

            # Only add "pulsed_" prefix if actually in pulsed mode
            mode_prefix = "pulsed_" if is_pulsed else ""
            filename = (
                f"{base_path}{file_prefix}{mode_prefix}LIV_LDBias({ld_start_mA},{ld_stop_mA})mA_"
                f"EAMBias({eam_bias_V})V_PDBias({pd_bias_V})V_{common_suffix}"
            )
        else:
            ld_bias_mA = int(laser_params['initval'])
            eam_start_V = eam_params['start']
            eam_stop_V = eam_params['stop']
            pd_bias_V = detector_params['initval']

            # Only add "pulsed_" prefix if actually in pulsed mode
            mode_prefix = "pulsed_" if is_pulsed else ""
            filename = (
                f"{base_path}{file_prefix}{mode_prefix}EAM_LDBias({ld_bias_mA})mA_"
                f"EAMBias({eam_start_V},{eam_stop_V})V_PDBias({pd_bias_V})V_{common_suffix}"
            )

        combined_df.to_excel(filename, index=False)
        print(f"\nCombined Excel file created successfully: {filename}")
        return True

    except Exception as e:
        print(f"Error creating Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False