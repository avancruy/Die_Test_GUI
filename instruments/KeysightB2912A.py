# class KeysightB2912A Precision Source/Measure Unit
# Author: E. Cauchon
# Created: Nov 20, 2024

import pyvisa
from pyvisa import resources

class KeysightB2912A:
    def __init__(self, resource_name: str):
        self.rm = pyvisa.ResourceManager()
        self.instrument = self.rm.open_resource(resource_name)
        #self.instrument.write('*RST')  # Reset the instrument to default settings

    def query(self, command):
        return self.instrument.query(command)

    def write(self, command):
        self.instrument.write(command)

    def read(self):
        return self.instrument.read()

    def get_idn(self):
        """
        query instrument identification (IDN)
        """
        idn = self.instrument.query('*IDN?')
        print(idn)
        return idn

    def set_mode(self, channel, mode):
        """
        Set mode (volts/amps) for a specified channel.
        :param channel: Channel number (1 or 2)
        :param mode (VOLT/CURR)
        """
        valid_modes = ['VOLT', 'CURR']
        if mode not in valid_modes:
            raise ValueError(f"Invalid mode. Choose from: {valid_modes}")
        self.instrument.write(f'SOUR{channel}:FUNC:MODE {mode}')

    def get_mode(self, channel):
        """
        get mode (volts/amps) for a specified channel.
        :param channel: Channel number (1 or 2)
        return mode: 'VOLT' or 'CURR'
        """
        return self.instrument.query(f'SOUR{channel}:FUNC:MODE?').strip()

    def set_voltage(self, channel, voltage):
        """
        Set the voltage for a specified channel.
        :param channel: Channel number (1 or 2)
        :param voltage: Voltage value to set
        """
        self.instrument.write(f':SOUR{channel}:VOLT {voltage}')

    def set_current(self, channel, current):
        """
        Set the currente for a specified channel.
        :param channel: Channel number (1 or 2)
        :param current: current value to set
        """
        self.instrument.write(f':SOUR{channel}:CURR {current}')

    def read_voltage(self, channel):
        """
        Read the voltage on a specified channel.
        :param channel: Channel number (1 or 2)
        :return: Voltage value
        """
        voltage = self.instrument.query(f':MEAS:VOLT? (@{channel})')
        return float(voltage)

    def read_current(self, channel):
        """
        Read the current on a specified channel.
        :param channel: Channel number (1 or 2)
        :return: Current value
        """
        current = self.instrument.query(f':MEAS:CURR? (@{channel})')
        return float(current)

    def output_on(self, channel):
        """
        Turn on the output for a specified channel.
        :param channel: Channel number (1 or 2)
        """
        self.instrument.write(f':OUTP{channel} ON')

    def output_off(self, channel):
        """
        Turn off the output for a specified channel.
        :param channel: Channel number (1 or 2)
        """
        self.instrument.write(f':OUTP{channel} OFF')

    def output_state(self, channel):
        """
        Get the output state for a specified channel.
        :param channel: Channel number (1 or 2)
        :return: 1 if output is on, 0 if output is off
        """
        state = self.instrument.query(f':OUTP{channel}:STATe?').strip()
        return int(state)

    def set_autorange(self, channel=1, on_off=1):
        """
        Set instrument autorage on (1), off(0).
        """
        self.instrument.write(f':SENS{channel}:RANG:AUTO {on_off}')

    def set_current_limit(self, channel, current_limit):
        """
        Set the limit for a specified channel.
        :param channel: Channel number (1 or 2)
        :param current_limit: limit value to set
        """
        self.instrument.write(f':SENS{channel}:FUNC CURR')
        self.instrument.write(f':SENS{channel}:CURR:PROT {current_limit}')

    def set_voltage_limit(self, channel, voltage_limit):
        """
        Set the limit for a specified channel.
        :param channel: Channel number (1 or 2)
        :param voltage_limit: limit value to set
        """
        self.instrument.write(f':SENS{channel}:FUNC VOLT')
        self.instrument.write(f':SENS{channel}:VOLT:PROT {voltage_limit}')



    def close(self):
        """
        Close the connection to the instrument.
        """
        self.instrument.close()
        self.rm.close()
        
        

if __name__ == "__main__":
    # Example usage:
    import time
    ip_addr='TCPIP0::10.20.0.184::hislip0::INSTR'  #'10.30.0.11'
    #smu = KeysightB2912A(f'TCPIP0::{ip_addr}::inst0::INSTR') #  #'TCPIP::K-B2912A-40841::inst0::INSTR'
    smu = KeysightB2912A(ip_addr) #  #'TCPIP::K-B2912A-40841::inst0::INSTR'
    smu.get_idn()
    smu.set_mode(1,'VOLT')
    smu.set_voltage(1, -3.0)
    smu.set_current_limit(1, 0.08)
    smu.output_on(1)
    print(smu.read_voltage(1))
    print(smu.read_current(1))
    time.sleep(1)
    smu.set_current(1, 0.020)
    smu.set_voltage_limit(1,2)
    print(smu.read_voltage(1))
    print(smu.read_current(1))
    time.sleep(1)
    smu.output_off(1)
    smu.set_autorange(1)
    smu.close()
