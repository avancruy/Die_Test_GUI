from turtle import delay

import pyvisa


# # connect to OSA
# rm = pyvisa.ResourceManager()
# # use rm.list_resources() to find OSA

##### a driver to remotely control the Anritsu MS9710C optical spectrum analyzer using RS-232C #####

class AnritsuMS9710CDriver:

    def __init__(self, port):
        self.port = port  # format: ASRL#::INSTR
        self.osa = pyvisa.ResourceManager()
        # match these settings with the OSA (see section 2.2.3 in the MS9710C remote operation manual)
        self.speed = 9600
        self.parity = "none"
        self.stopBit = 1
        self.characterLength = 8

        self.osa.read_termination = '\n'
        self.osa.write_termination = '\n'

    def setAddress(self, port):
        if "::INSTR" not in port:
            raise ConnectionError("ERROR: Please use this format:  ASRL#::INSTR")
        else:
            self.port = port

    # method to connect to OSA
    def open(self):
        self.osa = pyvisa.ResourceManager().open_resource(self.port)

    # method to send query and receive response from OSA
    def query(self, command):

        if "?" not in command:
            print('ERROR: Querys end with "?". Did you mean to use write method?')
            return None
        else:
            return self.osa.query(command)

    # method to send instruction to OSA
    def write(self, command):
        return self.osa.write(command)

    # The methods below can be found in section 3.3.3 in the MS9710C remote control operation manual

    # method to set center (in nm)
    def setCenter(self, center):
        self.write("CNT" + center)
        print("Center: " + self.query("CNT?") + "nm")

    # method to set sweep average
    def setAvg(self, avg):
        self.write("AVS" + avg)
        print("Sweep average: " + self.query("AVS?"))

    # method to set reference level (in dBm)
    def setRefValue(self, ref_value):
        self.write("RLV" + ref_value)
        print("Reference level: " + self.query("RLV?") + "dBm")

    # method to set span of sweep (in nm)
    def setSpan(self, span):
        self.write("SPN" + span)
        print("Span: " + self.query("SPN?") + "nm")

    # methods to set/get sensitivity
    def setSensitivity(self, sens):
        # unsure how to do this
        pass

    def getSensitivity(self):
        # unsure how to do this
        pass

    # methods to set/get sweep speed
    def setSweepSpeed(self, speed):
        # unsure how to do this
        pass

    def getSweepSpeed(self):
        # unsure how to do this
        pass

    # methods to set/get resolution (in nm)
    def setResolution(self, resolution):
        resolutions = ['0.05', '0.07', '0.1', '0.2', '0.5', '1']  # accepted resolutions by MS9710C
        if resolution not in resolutions:
            print("ERROR: Invalid resolution")
        else:
            self.write("RES" + resolution)
            print("Resolution: " + self.query("RES?") + "nm")

    def getResolution(self):
        return (self.query("RES?"))

    # method to perform single sweep
    def singleSweep(self, center=None, span=None):
        if center is not None:
            self.setCenter(center)
        if span is not None:
            self.setSpan(span)
        self.write("SSI")
        # see section 9.76 in operation manual to know when single sweep is complete
        # unsure how to implement that in code
        print("Performing single sweep... please wait...")
        delay(30000)
        print("Single sweep complete.")

    # returns peak wavelength in nm
    def getPeakWavelength(self):
        # find peak wavelength
        self.write("PKS PEAK")  # moves trace marker to peak

        # see section 9.64 in remote operation manual to know when this is complete
        # unsure how to implement that in code
        print("Searching for peak... please wait...")
        delay(30000)
        print("Search complete.")

        # read peak
        return self.query("TMK?")  # reads wavelength at trace marker location

    # returns side mode supression ratio
    def getSMSR(self):
        return self.query("ANA?")

    # method to center peak value on OSA
    def setPeakToCenter(self):
        self.setCenter(self.getPeakWavelength())

    # method to close connection with OSA
    def close(self):
        self.osa.close()

        # send commands to perform functions
        # receive data back
