import serial.tools.list_ports
from time import sleep

ports = serial.tools.list_ports.comports()
Portlist = []

# for one in ports:
#     Portlist.append(str(one))
#     print(str(one))
# com = input("Select COM port for Arduino #: ")

# for i in range(len(Portlist)):
#     if Portlist[i].startswith("COM" + str(com)):
#         use = "COM" + str(com)

mySerial = serial.Serial("COM6", baudrate= 9600, timeout=.1) #port=use

# mySerial.open()

# while True:
#     command = "ON"
#     mySerial.write(command.encode("utf-8"))
#     sleep(2)

    # if command == "exit":
    #     exit()

sleep(3)


command = "300\n"
mySerial.write(command.encode("utf-8"))
times = 0
while True:
    while mySerial.in_waiting > 0:
        msg = str(mySerial.readline(), "utf-8")
        print(msg)
        if msg == "h\r\n":
            print("Hai")
            # sleep(0.5)
            # command = "P,"
            # mySerial.write(command.encode("utf-8"))
            # times= times + 1

