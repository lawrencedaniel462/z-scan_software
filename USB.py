import pyvisa
from time import sleep

rm = pyvisa.ResourceManager('@py')
print(rm.list_resources())

detector = rm.open_resource('USB0::4883::32882::1904243::0::INSTR')
detector.query_delay = 0.1
detector.write_termination = '\n'
detector.read_termination = '\n'

detector.write("inp:pdi:filt:lpas:stat 0")
print(detector.query("inp:pdi:filt:lpas:stat?"))
print(detector.query("*IDN?"))                          # About the device
print(detector.query("sens:pow:rang:auto?"))            # Checking whether the range is in auto or not
print(detector.query("sens:corr:wav?"))                 # Querying the wave length
# detector.write("sens:corr:wav 6.35E+02")              # Setting the wave length
# detector.write('meas')                                #Start measurement and returns the data
# while True:
#     print(detector.query("read?"))
#     sleep(0.5)
# detector.write('sens:pow:ran:auto 1')
# print(detector.query("sens:pow:ran:auto?"))

# print(detector.query('sens:pow:ran:auto?'))
# print(detector.write('sens:pow:ran:auto?'))
