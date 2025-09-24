import pyvisa

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

    def set_autorange(self, channel=1, on_off=1):
        """
        Set instrument autorange on (1), off(0).
        """
        self.instrument.write(f':SENS{channel}:RANG:AUTO {on_off}')

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