# class DataMaster:
#     def __init__(self):
#         self.enableCommand = "E,"
#         self.disableCommand = "D,"
#         self.clockwiseCommand = "C,"
#         self.antiClockwiseCommand = "A,"
#         self.pulseCommand = "P,"
#         self.enable_ok = "e\r\n"
#         self.disable_ok = "d\r\n"
#         self.clockwise_ok = "c\r\n"
#         self.antiClockwise_ok = "a\r\n"
#         self.pulse_ok = "p\r\n"
#         self.home_ok = "h\r\n"
#         self.end_ok = "n\r\n"
import matplotlib.pyplot as plt

x = [0, 1, 2, 3, 4, 5]
y1 = [1000, 13000, 26000, 42000, 60000, 81000]
y2 = [1000, 13000, 27000, 43000, 63000, 85000]

plt.plot(x, y1)
plt.plot(x, y2, '-.')

plt.xlabel("X-axis data")
plt.ylabel("Y-axis data")
plt.title('multiple plots')
plt.show()

# from matplotlib.figure import Figure
# import numpy as np
#
# fig = Figure(figsize=(12, 4.4), dpi=80, facecolor="#f0f0f0")
# fig.tight_layout()
# ax = fig.add_subplot(111)
# fig.subplots_adjust(left=0.07, right=0.98, top=0.95, bottom=0.09)
# ax.set_xlabel("Distance in mm")
# ax.set_ylabel("Power in W")
# lines = ax.plot([], [])[0]