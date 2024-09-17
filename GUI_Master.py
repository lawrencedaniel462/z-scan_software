from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import threading
from time import sleep
import serial.tools.list_ports
import xml.etree.ElementTree as et
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
from tkinter import filedialog
import xlsxwriter

tree = et.parse("data_file.xml")


class RootGUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("Z Scan")
        self.root.geometry("1200x600")  # self.root.config(bg="white")


class CommGUI:
    def __init__(self, root, usb):
        self.motor_steps = 400
        self.root = root
        self.usb = usb

        self.xml = tree.getroot()

        self.notebook = ttk.Notebook(root)

        self.mainframe = Frame(self.notebook)  # bg="white"
        self.settingsframe = Frame(self.notebook)

        self.Config = LabelFrame(self.settingsframe, text="Configuration", padx=5, pady=5, width=470, height=325)

        self.wavelength_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.wavelength_entry.insert(0, self.xml[0][0].text)
        self.thread_spacing_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.thread_spacing_entry.insert(0, self.xml[0][2].text)
        self.open_aperture_dia_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.open_aperture_dia_entry.insert(0, self.xml[1][6].text)
        self.default_startvalue_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.default_startvalue_entry.insert(0, self.xml[1][0].text)
        self.default_stopvalue_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.default_stopvalue_entry.insert(0, self.xml[1][2].text)
        self.default_stepvalue_entry = Entry(self.Config, width=13, font=('Helvetica', 10))
        self.default_stepvalue_entry.insert(0, self.xml[1][4].text)

        self.wavelength = Label(self.Config, text="Wavelength", anchor="w", width=30, padx=7, pady=8)
        self.bandwidth = Label(self.Config, text="Bandwidth", anchor="w", width=30, padx=7, pady=8)
        self.thread_spacing = Label(self.Config, text="Thread Spacing", anchor="w", width=30, padx=7, pady=8)
        self.open_aperture_dia = Label(self.Config, text="Open Aperture Diameter", anchor="w", width=30, padx=7, pady=8)
        self.default_startvalue = Label(self.Config, text="Default distance from Home to Maxima", anchor="w", width=30,
                                        padx=7, pady=8)
        self.default_stopvalue = Label(self.Config, text="Default Half Sweeping Distance", anchor="w", width=30, padx=7,
                                       pady=8)
        self.default_stepvalue = Label(self.Config, text="Default Step Value", anchor="w", width=30, padx=7, pady=8)

        self.wavelength_unit = Label(self.Config, text="nm", anchor="w", width=12)
        self.open_aperture_unit = Label(self.Config, text="mm", anchor="w", width=12)
        self.thread_spacing_unit = Label(self.Config, text="mm", anchor="w", width=12)

        self.connection_frame = LabelFrame(self.mainframe, text="Communication Manager", padx=5, pady=5, width=520,
                                           height=165)
        self.label_usb = Label(self.connection_frame, text="Available USB Port(s): ", width=17,
                               anchor="w")  # bg="white"
        self.label_com = Label(self.connection_frame, text="Available COM Port(s): ", width=17,
                               anchor="w")  # bg="white"
        self.label_bd = Label(self.connection_frame, text="Baud Rate: ", width=17, anchor="w")  # bg="white"
        self.btn_refresh = Button(self.connection_frame, text="Refresh", width=10, command=self.Comm_Refresh)
        self.btn_usb_refresh = Button(self.connection_frame, text="Refresh", width=10, command=self.USB_Refresh)
        self.com_btn_connect = Button(self.connection_frame, text="Connect", width=10, state="disabled",
                                      command=self.SerialConnect)
        self.com_option_menu()
        self.BaudOptionMenu()
        self.USBOptionMenu()

        self.communication_frame = LabelFrame(self.mainframe, text="Controls Manager", pady=5, padx=5, width=650,
                                              height=165)
        self.homeButton = Button(self.communication_frame, text="Home", width=26, command=self.Home)
        self.endButton = Button(self.communication_frame, text="End", width=26, command=self.End)
        self.startvalue = Label(self.communication_frame, text="Distance from Home to Maxima", anchor="w", width=27)
        self.endvalue = Label(self.communication_frame, text="Sweeping Distance", anchor="w", width=27)
        self.stepvalue = Label(self.communication_frame, text="Step Value", anchor="w", width=27)

        self.startvalueentry = Entry(self.communication_frame, width=16)
        self.startvalueentry.insert(0, self.default_startvalue_entry.get())
        self.endvalueentry = Entry(self.communication_frame, width=16)
        self.endvalueentry.insert(0, self.default_stopvalue_entry.get())
        self.stepvalueentry = Entry(self.communication_frame, width=16)
        self.stepvalueentry.insert(0, self.default_stepvalue_entry.get())

        self.minimumvalue_text = float(self.thread_spacing_entry.get()) / self.motor_steps * 1000
        self.minimumvalue = Label(self.communication_frame,
                                  text=f"Note: The minimum step value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
        self.seperator = ttk.Separator(self.communication_frame, orient="vertical")

        self.graph_frame = LabelFrame(self.mainframe, text="Graph Manager", pady=5, padx=5, width=1180, height=385)
        self.graph_buttons_frame = Frame(self.graph_frame, width=190, height=350)
        self.power_label = Label(self.graph_buttons_frame, text="Power in W", anchor="center", width=26)
        self.power_value = Label(self.graph_buttons_frame, anchor="center", width=12, foreground="red", text="-",
                                 font=('Times New Roman', 18, 'bold'))
        self.close_aperture_entry = Entry(self.graph_buttons_frame, width=30)
        self.close_aperture_entry.insert(0,self.xml[1][7].text)
        self.run_button_without_aperture = Button(self.graph_buttons_frame, text="Start Run without Aperture", width=25, command=self.Run_without_aperture)
        self.run_button_with_aperture = Button(self.graph_buttons_frame, text="Start Run with Aperture", width=25, command=self.Run_with_aperture)
        self.enter_close_aperture_label = Label(self.graph_buttons_frame, text="Enter Close Aperture size in mm", width=25, anchor="center")
        self.enter_file_name_label = Label(self.graph_buttons_frame, text= "Enter file name", width=25, anchor="center")
        self.enter_file_name_entry = Entry(self.graph_buttons_frame, width=30)
        self.save_data_button = Button(self.graph_buttons_frame, text="Save Data", width=25, command=self.save_data_excel)
        self.save_image_data_button = Button(self.graph_buttons_frame, text="Save as image", width=25, command=self.save_as_image)
        self.reset_all_button = Button(self.graph_buttons_frame, text="Reset all", width=25, command=self.reset_all)
        self.seperator2 = ttk.Separator(self.graph_frame, orient="vertical")
        self.seperator3 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator4 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator5 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator6 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator7 = ttk.Separator(self.graph_frame, orient="horizontal")

        self.current_position = 0

        self.value_option_menu()

        self.button_frame = Frame(self.Config)

        self.save_button = Button(self.button_frame, text="Save", width=30, command=self.save_button)
        self.apply_button = Button(self.button_frame, text="Apply", width=30, command=self.apply_button)

        self.initialize_graph()
        self.initialize_motor_control()
        self.initialize_position_finder()

        self.detector = object()
        self.padx = 20
        self.pady = 5
        self.publish()
        self.tester = 0
        self.value_show_threading = False
        self.threading = False
        self.start_Initiation = False
        self.com_list = []
        self.is_home = False
        self.wavelength_value = self.xml[0][0].text
        self.bandwidth_value = self.xml[0][1].text

        self.notebook.hide(1)
        self.initialize_data()

        self.x_data = np.array([])
        self.y_data = np.array([])
        self.without_aperture_x_data = np.array([])
        self.without_aperture_y_data = np.array([])
        self.with_aperture_x_data = np.array([])
        self.with_aperture_y_data = np.array([])

    def initialize_data(self):
        self.enableCommand = "E,"
        self.disableCommand = "D,"
        self.clockwiseCommand = "C,"
        self.antiClockwiseCommand = "A,"
        self.pulseCommand = "P,"
        self.enable_ok = "e\r\n"
        self.disable_ok = "d\r\n"
        self.clockwise_ok = "c\r\n"
        self.antiClockwise_ok = "a\r\n"
        self.pulse_ok = "p\r\n"
        self.home_ok = "h\r\n"
        self.end_ok = "n\r\n"

    def initialize_motor_control(self):
        self.motor_control = LabelFrame(self.settingsframe, text="Manual Control", padx=5, pady=5, width=470,
                                        height=198)
        self.manual_button_frame = Frame(self.motor_control)
        self.manual_power_label = Label(self.manual_button_frame, text="-", anchor="center", width=15, foreground="red",
                                        font=('Times New Roman', 18, 'bold'))
        self.enable_manual_mode = Button(self.manual_button_frame, text="Enable manual mode", width=29,
                                         command=self.manual_mode_enable)
        self.enable_motor_button = Button(self.manual_button_frame, text="Disable Motor", width=29,
                                          command=self.enable_motor)
        self.move_label = Label(self.motor_control, text="Move")
        self.towards_label = Label(self.motor_control, text="towards")
        self.direction_label = Label(self.motor_control, text="direction")
        self.desired_distance = Entry(self.motor_control, width=10)
        options = ["mm", "microns"]
        direction = ["home", "end"]
        self.clicked_desired_unit = StringVar()
        self.clicked_desired_direction = StringVar()
        self.clicked_desired_unit.set(options[0])
        self.clicked_desired_direction.set(direction[0])
        self.drop_desired_unit = OptionMenu(self.motor_control, self.clicked_desired_unit, *options)
        self.drop_desired_direction = OptionMenu(self.motor_control, self.clicked_desired_direction, *direction)
        self.drop_desired_unit.config(width=10)
        self.drop_desired_direction.config(width=7)
        self.move_buttons_frame = Frame(self.motor_control)
        self.cw_button = Button(self.move_buttons_frame, text=">>", width=9, command=self.clockwise_step)
        self.move_button = Button(self.move_buttons_frame, text="Move", width=40, command=self.move)
        self.acw_button = Button(self.move_buttons_frame, text="<<", width=9, command=self.anticlockwise_step)

        self.enable_motor_button["state"] = "disabled"
        self.cw_button["state"] = "disabled"
        self.move_button["state"] = "disabled"
        self.acw_button["state"] = "disabled"
        self.desired_distance["state"] = "disabled"
        self.drop_desired_unit["state"] = "disabled"
        self.drop_desired_direction["state"] = "disabled"

    def initialize_position_finder(self):
        self.position_finder_frame = LabelFrame(self.settingsframe, text="Position Locator", padx=5, pady=5, width=470,
                                                height=117)
        self.find_position = Button(self.position_finder_frame, text="Locate the position", width=62,
                                    command=self.locate)
        self.holder_position = Label(self.position_finder_frame)
        self.text_processor(0)

    def text_processor(self, position):
        if position == 0:
            self.holder_position[
                "text"] = 'Click the "Locate the Position" button\nto know the current location of the holder from home position'
        else:
            self.holder_position[
                "text"] = f"The distance from the initial position to home position is\n{position / self.motor_steps} mm or {position * 1000 / self.motor_steps} microns"

    def locate(self):
        if self.find_position["text"] == "Locate the position":
            if self.homeButton["text"] == "Home" and self.endButton["text"] == "End" and self.run_button_without_aperture[
                "text"] == "Start Run without Aperture" and self.enable_manual_mode["text"] == "Enable manual mode":
                self.enable_manual_mode["state"] = "disabled"
                self.notebook.hide(0)
                self.find_position["text"] = "Stop"
                self.save_button["state"] = "disabled"
                self.apply_button["state"] = "disabled"
                self.t1 = threading.Thread(target=self.locateThreading, daemon=True)
                self.t1.start()
            else:
                messagebox.showerror("Process Error", "Stop all the processes and try again")
        else:
            self.threading = False
            self.find_position["text"] = "Locate the position"
            self.enable_manual_mode["state"] = "active"
            self.notebook.hide(0)
            self.save_button["state"] = "active"
            self.apply_button["state"] = "active"
            self.notebook.add(self.mainframe, text="Controls")
            self.text_processor(0)

    def manual_mode_enable(self):
        if self.enable_manual_mode["text"] == "Enable manual mode":
            if self.homeButton["text"] == "Home" and self.endButton["text"] == "End" and self.run_button_without_aperture[
                "text"] == "Start Run without Aperture" and self.find_position["text"] == "Locate the position":
                self.enable_manual_mode["text"] = "Disable manual mode"
                self.enable_motor_button["state"] = "active"
                self.cw_button["state"] = "active"
                self.move_button["state"] = "active"
                self.acw_button["state"] = "active"
                self.desired_distance["state"] = "normal"
                self.drop_desired_unit["state"] = "active"
                self.drop_desired_direction["state"] = "active"
                self.notebook.hide(0)
                self.save_button["state"] = "disabled"
                self.apply_button["state"] = "disabled"
                self.find_position["state"] = "disabled"

            else:
                messagebox.showerror("Process Error", "Stop all the auto processes and try again")
        else:
            if self.enable_motor_button["text"] == "Disable Motor":
                self.enable_manual_mode["text"] = "Enable manual mode"
                self.enable_motor_button["state"] = "disabled"
                self.cw_button["state"] = "disabled"
                self.move_button["state"] = "disabled"
                self.acw_button["state"] = "disabled"
                self.desired_distance["state"] = "disabled"
                self.drop_desired_unit["state"] = "disabled"
                self.drop_desired_direction["state"] = "disabled"
                self.notebook.add(self.mainframe, text="Controls")
                self.save_button["state"] = "active"
                self.apply_button["state"] = "active"
                self.find_position["state"] = "active"
            else:
                messagebox.showerror("Motor Status Error", "First enable motor to disable manual mode")

    def anticlockwise_step(self):
        if self.enable_motor_button["text"] == "Disable Motor":
            self.sendCommand(self.antiClockwiseCommand)
            if self.receiveCommand() == self.antiClockwise_ok:
                self.sendCommand(self.pulseCommand)
                if self.receiveCommand() == self.pulse_ok:
                    pass
                else:
                    messagebox.showinfo("Position Status", "Home Reached")
        else:
            messagebox.showerror("Motor Status Error", "First enable motor to move 1 step Anticlockwise direction")

    def clockwise_step(self):
        if self.enable_motor_button["text"] == "Disable Motor":
            self.sendCommand(self.clockwiseCommand)
            if self.receiveCommand() == self.clockwise_ok:
                self.sendCommand(self.pulseCommand)
                if self.receiveCommand() == self.pulse_ok:
                    pass
                else:
                    messagebox.showinfo("Position Status", "Reached the End")
        else:
            messagebox.showerror("Motor Status Error", "First enable motor to move 1 step Clockwise direction")

    def enable_motor(self):
        if self.enable_motor_button["text"] == "Disable Motor":
            self.sendCommand(self.disableCommand)
            if self.receiveCommand() == self.disable_ok:
                self.enable_motor_button["text"] = "Enable Motor"
        else:
            self.sendCommand(self.enableCommand)
            if self.receiveCommand() == self.enable_ok:
                self.enable_motor_button["text"] = "Disable Motor"

    def move(self):
        if self.move_button["text"] == "Move":
            if self.enable_motor_button["text"] == "Disable Motor":
                if len(self.desired_distance.get()) > 0:
                    if self.num_validator(self.desired_distance.get()):
                        self.t1 = threading.Thread(target=self.moveThreading, daemon=True)
                        self.t1.start()
                        self.move_button["text"] = "Stop"
                        self.enable_manual_mode["state"] = "disabled"
                        self.enable_motor_button["state"] = "disabled"
                        self.cw_button["state"] = "disabled"
                        self.acw_button["state"] = "disabled"
                        self.desired_distance["state"] = "disabled"
                        self.drop_desired_unit["state"] = "disabled"
                        self.drop_desired_direction["state"] = "disabled"
                    else:
                        messagebox.showerror("Value Error", "Distance field should be a number")
                else:
                    messagebox.showerror("Value Error", "Distance field should not be blank")
            else:
                messagebox.showerror("Motor Status Error", "First enable motor to disable manual mode")
        else:
            self.start_Initiation = False
            self.move_button["text"] = "Move"
            self.enable_manual_mode["state"] = "active"
            self.enable_motor_button["state"] = "active"
            self.cw_button["state"] = "active"
            self.acw_button["state"] = "active"
            self.desired_distance["state"] = "normal"
            self.drop_desired_unit["state"] = "active"
            self.drop_desired_direction["state"] = "active"

    def initialize_graph(self):
        self.fig = Figure(figsize=(12, 4.4), dpi=80, facecolor="#f0f0f0")
        self.fig.tight_layout()
        self.ax = self.fig.add_subplot(111)
        self.fig.subplots_adjust(left=0.07, right=0.98, top=0.95, bottom=0.09)
        self.ax.set_xlabel("Distance in mm")
        self.ax.set_ylabel("Power in W")
        self.lines = self.ax.plot([], [])[0]
        self.lines2 = self.ax.plot([], [])[0]
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, rowspan=5)

    def value_option_menu(self):
        options = ["mm", "microns"]
        bandwidth_options = ["High", "Low"]
        self.clicked_bandwidth = StringVar()
        self.clicked_startvalue = StringVar()
        self.clicked_stopvalue = StringVar()
        self.clicked_stepvalue = StringVar()
        self.default_clicked_startvalue = StringVar()
        self.default_clicked_stopvalue = StringVar()
        self.default_clicked_stepvalue = StringVar()

        if self.xml[0][1].text == "High":
            self.clicked_bandwidth.set(bandwidth_options[0])
        elif self.xml[0][1].text == "Low":
            self.clicked_bandwidth.set(bandwidth_options[1])

        if self.xml[1][1].text == "mm":
            self.default_clicked_startvalue.set(options[0])
        elif self.xml[1][1].text == "microns":
            self.default_clicked_startvalue.set(options[1])

        if self.xml[1][3].text == "mm":
            self.default_clicked_stopvalue.set(options[0])
        elif self.xml[1][3].text == "microns":
            self.default_clicked_stopvalue.set(options[1])

        if self.xml[1][5].text == "mm":
            self.default_clicked_stepvalue.set(options[0])
        elif self.xml[1][5].text == "microns":
            self.default_clicked_stepvalue.set(options[1])

        if self.default_clicked_startvalue.get() == "mm":
            self.clicked_startvalue.set(options[0])
        elif self.default_clicked_startvalue.get() == "microns":
            self.clicked_startvalue.set(options[1])

        if self.default_clicked_stopvalue.get() == "mm":
            self.clicked_stopvalue.set(options[0])
        elif self.default_clicked_stopvalue.get() == "microns":
            self.clicked_stopvalue.set(options[1])

        if self.default_clicked_stepvalue.get() == "mm":
            self.clicked_stepvalue.set(options[0])
        elif self.default_clicked_stepvalue.get() == "microns":
            self.clicked_stepvalue.set(options[1])

        self.drop_bandwidth = OptionMenu(self.Config, self.clicked_bandwidth, *bandwidth_options)
        self.drop_startvalue = OptionMenu(self.communication_frame, self.clicked_startvalue, *options)
        self.drop_stopvalue = OptionMenu(self.communication_frame, self.clicked_stopvalue, *options)
        self.drop_stepvalue = OptionMenu(self.communication_frame, self.clicked_stepvalue, *options)
        self.default_drop_startvalue = OptionMenu(self.Config, self.default_clicked_startvalue, *options)
        self.default_drop_stopvalue = OptionMenu(self.Config, self.default_clicked_stopvalue, *options)
        self.default_drop_stepvalue = OptionMenu(self.Config, self.default_clicked_stepvalue, *options)

        self.drop_bandwidth.config(width=9)
        self.drop_startvalue.config(width=10)
        self.drop_stopvalue.config(width=10)
        self.drop_stepvalue.config(width=10)
        self.default_drop_startvalue.config(width=8)
        self.default_drop_stopvalue.config(width=8)
        self.default_drop_stepvalue.config(width=8)

    def USBOptionMenu(self):
        self.tuple1 = self.usb.list_resources()
        self.list1 = list(self.tuple1)
        self.list1.insert(0, "-")
        self.clicked_usb = StringVar()
        self.clicked_usb.set(self.list1[0])
        self.drop_usb = OptionMenu(self.connection_frame, self.clicked_usb, *self.list1, command=self.Connect_Ctrl)
        self.drop_usb.config(width=35)

    def com_option_menu(self):
        self.get_com_list()
        self.clicked_com = StringVar()
        self.clicked_com.set(self.com_list[0])
        self.drop_com = OptionMenu(self.connection_frame, self.clicked_com, *self.com_list, command=self.Connect_Ctrl)
        self.drop_com.config(width=35)

    def BaudOptionMenu(self):
        bds = ["-", "300", "600", "1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "56000", "57600",
               "115200", "128000", "256000"]
        self.clicked_bd = StringVar()
        self.clicked_bd.set(bds[0])
        self.drop_bd = OptionMenu(self.connection_frame, self.clicked_bd, *bds, command=self.Connect_Ctrl)
        self.drop_bd.config(width=35)

    def Connect_Ctrl(self, Other):
        if "-" in self.clicked_com.get() or "-" in self.clicked_bd.get() or "-" in self.clicked_usb.get():
            self.com_btn_connect["state"] = "disable"
        else:
            self.com_btn_connect["state"] = "active"

    def Comm_Refresh(self):
        self.drop_com.destroy()
        self.com_option_menu()
        self.drop_com.grid(column=2, row=2, padx=self.padx, pady=self.pady)
        logic = []
        self.Connect_Ctrl(logic)

    def USB_Refresh(self):
        self.drop_usb.destroy()
        self.USBOptionMenu()
        self.drop_usb.grid(column=2, row=1, padx=self.padx, pady=self.pady)
        logic = []
        self.Connect_Ctrl(logic)

    def SerialConnect(self):
        if self.com_btn_connect["text"] in "Connect":
            usb_status = False
            try:
                self.detector.query("*IDN?")
                usb_status = True
            except:
                self.detector = self.usb.open_resource(self.clicked_usb.get())
                self.detector.query_delay = 0.1
                self.detector.write_termination = '\n'
                self.detector.read_termination = '\n'
                # self.usb.close()
                try:
                    self.detector.query("*IDN?")
                    usb_status = True
                except:
                    USBErrorMsg = "USB device is not connected"
                    messagebox.showerror("Connection Error", USBErrorMsg)
            self.SerialOpen()
            if self.ser.status:
                pass
            else:
                ErrorMsg = f"Failure to establish UART Connection using {self.clicked_com.get()}"
                messagebox.showerror("Connection Error", ErrorMsg)
            if self.ser.status and usb_status:
                self.com_btn_connect["text"] = "Disconnect"
                self.btn_usb_refresh["state"] = "disabled"
                self.btn_refresh["state"] = "disabled"
                self.drop_usb["state"] = "disabled"
                self.drop_bd["state"] = "disabled"
                self.drop_com["state"] = "disabled"
                Info_msg = f"Successful UART Connection using {self.clicked_com.get()}\nSuccessful USB Connection using {self.clicked_usb.get()}"
                self.t2 = threading.Thread(target=self.display_data, daemon=True)
                self.t2.start()
                self.enable_buttons()
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Connection Successful", Info_msg)
        else:
            self.SerialClose()
            self.value_show_threading = False
            self.threading = False
            self.com_btn_connect["text"] = "Connect"
            self.btn_refresh["state"] = "active"
            self.btn_usb_refresh["state"] = "active"
            self.drop_usb["state"] = "active"
            self.drop_bd["state"] = "active"
            self.drop_com["state"] = "active"
            self.disable_buttons()
            self.notebook.hide(1)

    def get_com_list(self):
        ports = serial.tools.list_ports.comports()
        self.com_list = [com[0] for com in ports]
        self.com_list.insert(0, "-")

    def SerialOpen(self):
        try:
            self.ser.is_open
        except:
            port = self.clicked_com.get()
            baud = self.clicked_bd.get()
            self.ser = serial.Serial()
            self.ser.baudrate = baud
            self.ser.port = port
            self.ser.timeout = 0.5
        try:
            if self.ser.is_open:
                self.ser.status = True
            else:
                port = self.clicked_com.get()
                baud = self.clicked_bd.get()
                self.ser = serial.Serial()
                self.ser.baudrate = baud
                self.ser.port = port
                self.ser.timeout = 0.1
                self.ser.open()
                self.ser.status = True
        except:
            self.ser.status = False

    def SerialClose(self):
        try:
            if self.ser.is_open:
                self.ser.close()
                self.ser.status = False
        except:
            self.ser.status = False

    def display_data(self):
        self.value_show_threading = True
        current_wavelength = self.wavelength_value
        current_bandwidth = self.bandwidth_value
        while self.value_show_threading:
            try:
                if current_wavelength != self.wavelength_value:
                    self.detector.write(f"sens:corr:wav {float(self.wavelength_value)}")
                    current_wavelength = self.wavelength_value
                if current_bandwidth != self.bandwidth_value:
                    if self.bandwidth_value == "High":
                        self.detector.write("inp:pdi:filt:lpas:stat 1")
                        current_bandwidth = self.bandwidth_value
                    elif self.bandwidth_value == "Low":
                        self.detector.write("inp:pdi:filt:lpas:stat 0")
                        current_bandwidth = self.bandwidth_value
                value = self.detector.query("read?")
                self.power_value.config(text=str(value))
                self.manual_power_label.config(text=str(value) + " W")
                sleep(0.5)
                if not self.value_show_threading:
                    self.power_value.config(text="-")
                    break
            except Exception as e:
                self.value_show_threading = False
                if not self.value_show_threading:
                    self.power_value.config(text="-")
                    break

    def sendCommand(self, command):
        self.command = command
        self.ser.reset_output_buffer()
        self.ser.write(self.command.encode("utf-8"))

    def receiveCommand(self):
        self.tester = 0
        while self.tester == 0:
            while self.ser.in_waiting > 0:
                msg = str(self.ser.readline(), "utf-8")
                self.ser.reset_input_buffer()
                self.tester = 1
                return msg

    def homeThreading(self):
        self.threading = True
        self.start_Initiation = True
        self.sendCommand(self.antiClockwiseCommand)
        if self.receiveCommand() == self.antiClockwise_ok:
            self.notebook.hide(1)
            while self.threading & self.start_Initiation:
                try:
                    self.sendCommand(self.pulseCommand)
                    if self.receiveCommand() == self.pulse_ok:
                        self.threading = True
                    else:
                        self.threading = False
                    if self.threading == False:
                        self.notebook.add(self.settingsframe, text="Settings")
                        messagebox.showinfo("Position Status", "Home Reached")
                        self.is_home = True
                        self.status_retainer()
                        break
                except Exception:
                    self.threading = False
                if self.threading == False:
                    break

    def endThreading(self):
        self.threading = True
        self.start_Initiation = True
        self.sendCommand(self.clockwiseCommand)
        if self.receiveCommand() == self.clockwise_ok:
            self.notebook.hide(1)
            while self.threading & self.start_Initiation:
                try:
                    self.sendCommand(self.pulseCommand)
                    if self.receiveCommand() == self.pulse_ok:
                        self.threading = True
                        self.is_home = False
                    else:
                        self.threading = False
                    if self.threading == False:
                        self.notebook.add(self.settingsframe, text="Settings")
                        messagebox.showinfo("Position Status", "Reached the End")
                        self.status_retainer()
                        break
                except Exception:
                    self.threading = False
                if self.threading == False:
                    break

    def moveThreading(self):
        self.threading = True
        self.start_Initiation = True
        pulse_number = 0
        if self.clicked_desired_unit.get() == "mm":
            if ((float(self.desired_distance.get()) * 1000) % self.minimumvalue_text) == 0:
                pulse_number = float(self.desired_distance.get()) * self.motor_steps
            else:
                self.threading = False
                messagebox.showerror("Error",
                                     f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
        elif self.clicked_desired_unit.get() == "microns":
            if (float(self.desired_distance.get()) % self.minimumvalue_text) == 0:
                pulse_number = (float(self.desired_distance.get()) / 1000) * self.motor_steps
            else:
                self.threading = False
                messagebox.showerror("Error",
                                     f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
        if self.desired_distance.get() == "end":
            self.sendCommand(self.clockwiseCommand)
        else:
            self.sendCommand(self.antiClockwiseCommand)
            if self.receiveCommand() == self.antiClockwise_ok:
                while (pulse_number + 1) != 0 and self.threading and self.start_Initiation:
                    self.sendCommand(self.pulseCommand)
                    current_status = self.receiveCommand()
                    if current_status == self.pulse_ok:
                        self.threading = True
                    elif current_status == self.home_ok:
                        self.threading = False
                        messagebox.showinfo("Position Status", "Home Reached")
                    elif current_status == self.end_ok:
                        self.threading = False
                        messagebox.showinfo("Position Status", "Reached the End")
                    if not self.start_Initiation:
                        break
                    if not self.threading:
                        break
                    pulse_number = pulse_number - 1
                self.move_button["text"] = "Move"
                self.enable_manual_mode["state"] = "active"
                self.enable_motor_button["state"] = "active"
                self.cw_button["state"] = "active"
                self.acw_button["state"] = "active"
                self.desired_distance["state"] = "normal"
                self.drop_desired_unit["state"] = "active"
                self.drop_desired_direction["state"] = "active"

    def locateThreading(self):
        self.threading = True
        position = 0
        while self.threading:
            self.sendCommand(self.antiClockwiseCommand)
            if self.receiveCommand() == self.antiClockwise_ok:
                self.sendCommand(self.pulseCommand)
                command = self.receiveCommand()
                if command == self.pulse_ok:
                    position = position + 1
                    if self.threading:
                        self.text_processor(position)
                elif command == self.home_ok:
                    if self.threading:
                        self.text_processor(position)
                    self.threading = False
                    self.enable_manual_mode["state"] = "active"
                    self.notebook.hide(0)
                    self.find_position["text"] = "Locate the position"
                    self.save_button["state"] = "active"
                    self.apply_button["state"] = "active"
                    self.notebook.add(self.mainframe, text="Controls")
            if not self.threading:
                break

    def parameter_checker(self):
        if len(self.startvalueentry.get()) != 0:
            if self.num_validator(self.startvalueentry.get()):
                if len(self.endvalueentry.get()) != 0:
                    if self.num_validator(self.endvalueentry.get()):
                        if len(self.stepvalueentry.get()) != 0:
                            if self.num_validator(self.endvalueentry.get()):
                                home_to_maxima = float(self.startvalueentry.get())
                                sweeping_distance = float(self.endvalueentry.get())
                                step_value = float(self.stepvalueentry.get())
                                if self.clicked_startvalue.get() == "mm":
                                    home_to_maxima = home_to_maxima * 1000
                                if self.clicked_stopvalue.get() == "mm":
                                    sweeping_distance = sweeping_distance * 1000
                                if self.clicked_stepvalue.get() == "mm":
                                    step_value = step_value * 1000
                                if home_to_maxima > sweeping_distance:
                                    if step_value % self.minimumvalue_text == 0:
                                        return True
                                    else:
                                        messagebox.showerror("Error",
                                                             f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                                else:
                                    messagebox.showerror("Value Error",
                                                         'The value in "Distance from Home to Maxima" field should be less than Sweeping distance value')
                            else:
                                messagebox.showerror("Value Error", "Step Value field should be a number")
                        else:
                            messagebox.showerror("Value Error", "Step Value field should not be blank")
                    else:
                        messagebox.showerror("Value Error", "Sweeping Distance field should be a number")
                else:
                    messagebox.showerror("Value Error", "Sweeping Distance field should not be blank")
            else:
                messagebox.showerror("Value Error", '"Distance from Home to Maxima" field should be a number')
        else:
            messagebox.showerror("Value Error", '"Distance from Home to Maxima" field should not be blank')

    def go_home(self):
        while self.threading and self.start_Initiation:
            self.sendCommand(self.antiClockwiseCommand)
            if self.receiveCommand() == self.antiClockwise_ok:
                self.sendCommand(self.pulseCommand)
                command = self.receiveCommand()
                if command == self.pulse_ok:
                    pass
                elif command == self.home_ok:
                    break
                else:
                    self.threading = False
                if not self.threading or not self.start_Initiation:
                    break

    def take_position(self, pulse_number):
        pulse_number = pulse_number * self.motor_steps
        while self.threading and self.start_Initiation and ((pulse_number + 1) > 0):
            self.sendCommand(self.pulseCommand)
            current_status = self.receiveCommand()
            if current_status == self.pulse_ok:
                self.threading = True
            elif current_status == self.end_ok:
                self.threading = False
                messagebox.showinfo("Position Status", "Reached the End")
            if not self.start_Initiation:
                break
            if not self.threading:
                break
            pulse_number = pulse_number - 1

    def run_Threading(self):
        sleep(0.1)
        self.threading = True
        if not self.parameter_checker():
            self.start_Initiation = False
        else:
            self.go_home()
            if self.clicked_startvalue.get() == "mm":
                maxima_point_from_home_in_mm = float(self.startvalueentry.get())
            else:
                maxima_point_from_home_in_mm = (float(self.startvalueentry.get()) / 1000)
            if self.clicked_stopvalue.get() == "mm":
                sweeping_distance_in_mm = float(self.endvalueentry.get())
            else:
                sweeping_distance_in_mm = (float(self.endvalueentry.get()) / 1000)
            if self.clicked_stepvalue.get() == "mm":
                stepvalue_in_mm = float(self.stepvalueentry.get())
            else:
                stepvalue_in_mm = (float(self.stepvalueentry.get()) / 1000)
            self.sendCommand(self.clockwiseCommand)
            if self.receiveCommand() == self.clockwise_ok:
                self.take_position(maxima_point_from_home_in_mm - sweeping_distance_in_mm)
            initial_point = -sweeping_distance_in_mm
            final_point = sweeping_distance_in_mm
            self.ax.set_xlim([initial_point, final_point])
            current_point = initial_point
            self.x_data = np.array([])
            self.y_data = np.array([])
            self.lines2.set_xdata(self.x_data)
            self.lines2.set_ydata(self.y_data)
            self.lines.set_xdata(self.x_data)
            self.lines.set_ydata(self.y_data)
            if self.run_button_without_aperture["text"] == "Stop Run":
                self.lines.set_color("blue")
            elif self.run_button_with_aperture["text"] == "Stop Run":
                self.lines.set_color("red")
            self.canvas.draw()
            while (current_point < final_point) and self.threading and self.start_Initiation:
                self.x_data = np.append(self.x_data, current_point)
                self.y_data = np.append(self.y_data, float(self.detector.query("read?")))
                y_max = self.y_data.max()
                y_min = self.y_data.min()

                y_min = y_min - (y_min * 0.01)
                y_max = y_max + (y_max * 0.01)
                self.ax.set_ylim([y_min, y_max])
                self.lines.set_xdata(self.x_data)
                self.lines.set_ydata(self.y_data)

                self.canvas.draw()

                current_point = current_point + stepvalue_in_mm
                self.take_position(stepvalue_in_mm)

                if not self.start_Initiation:
                    break
                if not self.threading:
                    break
            self.threading = False
        if not self.start_Initiation:
            self.x_data = np.array([])
            self.y_data = np.array([])
            self.lines.set_xdata(self.x_data)
            self.lines.set_ydata(self.y_data)
            self.canvas.draw()
            self.status_retainer()
            sleep(0.2)
            self.t2 = threading.Thread(target=self.display_data, daemon=True)
            self.t2.start()
        elif not self.threading:
            if self.run_button_without_aperture["text"] == "Stop Run":
                self.without_aperture_x_data = self.x_data
                self.without_aperture_y_data = self.y_data
            elif self.run_button_with_aperture["text"] == "Stop Run":
                self.with_aperture_x_data = self.x_data
                self.with_aperture_y_data = self.y_data
            self.x_data = np.array([])
            self.x_data = np.array([])
            self.status_retainer()
            sleep(0.2)
            self.t2 = threading.Thread(target=self.display_data, daemon=True)
            self.t2.start()
            self.run_button_with_aperture["state"] = "active"
            self.close_aperture_entry["state"] = "normal"
            self.run_button_without_aperture["state"] = "active"
            self.save_data_button["state"] = "active"
            self.enter_file_name_entry["state"] = "normal"
            self.save_image_data_button["state"] = "active"
            self.reset_all_button["state"] = "active"

    def reset_all(self):
        self.x_data = np.array([])
        self.y_data = np.array([])
        self.without_aperture_x_data = np.array([])
        self.without_aperture_y_data = np.array([])
        self.with_aperture_x_data = np.array([])
        self.with_aperture_y_data = np.array([])
        self.lines.set_xdata(self.x_data)
        self.lines.set_ydata(self.y_data)
        self.lines2.set_xdata(self.x_data)
        self.lines2.set_ydata(self.y_data)
        self.canvas.draw()
        self.notebook.add(self.settingsframe, text="Settings")
        self.enable_buttons()
        self.run_button_with_aperture["text"] = "Start Run with Aperture"
        self.run_button_without_aperture["text"] = "Start Run without Aperture"
        self.save_data_button["state"] = "disable"
        self.enter_file_name_entry["state"] = "disabled"
        self.reset_all_button["state"] = "disable"
        self.save_image_data_button["state"] = "disable"
        self.com_btn_connect["state"] = "active"

    def save_as_image(self):
        if (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) == 0):
            self.without_aperture_plotter()
        elif (len(self.without_aperture_x_data) == 0) and (len(self.with_aperture_x_data) != 0):
            self.with_aperture_plotter()
        elif (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) != 0):
            self.plot_selector()

    def plot_selector(self):
        self.plot_window = Toplevel()
        self.plot_window.title("Plot Selector")
        self.plot_window.geometry("250x120")
        self.without_aperture_var = IntVar()
        Checkbutton(self.plot_window, text="Plot without aperture", variable=self.without_aperture_var).grid(row=0, column=0, pady=5, sticky=W)
        self.with_aperture_var = IntVar()
        Checkbutton(self.plot_window, text="Plot with aperture", variable=self.with_aperture_var).grid(row=1, column=0, pady=5, sticky=W)
        self.plot_the_graph_button = Button(self.plot_window, text="Plot the Graph", command=self.plot_the_graph)
        self.plot_the_graph_button.grid(row=2, column=0, pady=5)

    def plot_the_graph(self):
        if self.without_aperture_var.get() == 0 and self.with_aperture_var.get() == 0:
            messagebox.showerror("Selection Error", "Select at least one box")
        elif self.without_aperture_var.get() == 1 and self.with_aperture_var.get() == 0:
            self.without_aperture_plotter()
        elif self.without_aperture_var.get() == 0 and self.with_aperture_var.get() == 1:
            self.with_aperture_plotter()
        elif self.without_aperture_var.get() == 1 and self.with_aperture_var.get() == 1:
            plt.plot(self.without_aperture_x_data, self.without_aperture_y_data, color="tab:blue")
            plt.plot(self.with_aperture_x_data, self.with_aperture_y_data, color="tab:red")
            plt.legend(["Open Aperture", "Closed Aperture"], loc="upper right")
            plt.xlabel("Distance in mm")
            plt.ylabel("Power in W")
            plt.show()
        self.plot_window.destroy()

    def without_aperture_plotter(self):
        plt.plot(self.without_aperture_x_data, self.without_aperture_y_data, color="tab:blue")
        plt.xlabel("Distance in mm")
        plt.ylabel("Power in W")
        plt.show()

    def with_aperture_plotter(self):
        plt.plot(self.with_aperture_x_data, self.with_aperture_y_data, color="tab:red")
        plt.xlabel("Distance in mm")
        plt.ylabel("Power in W")
        plt.show()


    def get_file(self):
        # regex = re.compile('.')
        # if (regex.search(file_name_entry.get())==None):
        #     return True
        if len(self.enter_file_name_entry.get()) == 0:
            messagebox.showerror("Naming Error", "File name should not be empty")
            return False
        # elif:
        #     messagebox.showerror("Naming Error", "Special Characters are not allowed")
        #     return False
        else:
            return True

    def excel_selector(self):
        self.excel_window = Toplevel()
        self.excel_window.title("Excel Data Selector")
        self.excel_window.geometry("250x120")
        self.excel_without_aperture_var = IntVar()
        Checkbutton(self.excel_window, text="Plot without aperture", variable=self.excel_without_aperture_var).grid(row=0,
                                                                                                          column=0,
                                                                                                          pady=5,
                                                                                                          sticky=W)
        self.excel_with_aperture_var = IntVar()
        Checkbutton(self.excel_window, text="Plot with aperture", variable=self.excel_with_aperture_var).grid(row=1, column=0,
                                                                                                    pady=5, sticky=W)
        self.export_excel = Button(self.excel_window, text="Prepare Data", command=self.excel_mode_selector)
        self.export_excel.grid(row=2, column=0, pady=5)

    def excel_mode_selector(self):
        if self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 0:
            messagebox.showerror("Selection Error", "Select at least one box")
        elif self.excel_without_aperture_var.get() == 1 and self.excel_with_aperture_var.get() == 0:
            self.without_aperture_plotter()
        elif self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 1:
            self.with_aperture_plotter()
        elif self.excel_without_aperture_var.get() == 1 and self.excel_with_aperture_var.get() == 1:
            if self.get_file():
                file_path = filedialog.askdirectory(title="Select the folder to save")
                file_path = file_path + f"/{self.enter_file_name_entry.get()}.xlsx"
                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet()
                bold = workbook.add_format({'bold': True})
                border = workbook.add_format({'border': 1})
                bold_and_border = workbook.add_format({'border': 1, 'bold': True})
                bold_and_border_and_center = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
                worksheet.set_column_pixels(0, 3, 100)

                worksheet.merge_range("A1:B1", "Wavelength", bold_and_border)
                worksheet.merge_range("A2:B2", "Bandwidth", bold_and_border)
                worksheet.merge_range("A3:B3", "Open Aperture Diameter", bold_and_border)
                worksheet.merge_range("A4:B4", "Close Aperture Diameter", bold_and_border)
                worksheet.merge_range("A5:B5", "Sweeping Distance", bold_and_border)
                worksheet.merge_range("A6:B6", "Step Distance", bold_and_border)

                worksheet.write("C1", f"{self.xml[0][0].text} nm", border)
                worksheet.write("C2", self.xml[0][1].text, border)
                worksheet.write("C3", f"{self.xml[1][6].text} mm", border)
                worksheet.write("C4", f"{self.xml[1][7].text} mm", border)
                worksheet.write("C5", self.endvalueentry.get(), border)
                worksheet.write("C6", self.stepvalueentry.get(), border)

                worksheet.merge_range("A8:B8", "Open Aperture", bold_and_border_and_center)
                worksheet.write("A9", "Distance in mm", bold_and_border)
                worksheet.write("B9", "Power in W", bold_and_border)

                worksheet.merge_range("C8:D8", "Close Aperture", bold_and_border_and_center)
                worksheet.write("C9", "Distance in mm", bold_and_border)
                worksheet.write("D9", "Power in W", bold_and_border)

                row = 9
                column = 0

                for item in self.without_aperture_x_data:
                    worksheet.write(row, column, item, border)
                    row += 1

                row = 9
                column = 1

                for item in self.without_aperture_y_data:
                    worksheet.write(row, column, item, border)
                    row += 1

                row = 9
                column = 2

                for item in self.with_aperture_x_data:
                    worksheet.write(row, column, item, border)
                    row += 1

                row = 9
                column = 3

                for item in self.with_aperture_y_data:
                    worksheet.write(row, column, item, border)
                    row += 1

                workbook.close()
        self.excel_window.destroy()

    def without_aperture_excel(self):
        if self.get_file():
            file_path = filedialog.askdirectory(title="Select the folder to save")
            file_path = file_path + f"/{self.enter_file_name_entry.get()} open aperture.xlsx"
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            border = workbook.add_format({'border': 1})
            bold_and_border = workbook.add_format({'border': 1, 'bold': True})
            bold_and_border_and_center = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
            worksheet.set_column_pixels(0, 3, 100)

            worksheet.merge_range("A1:B1", "Wavelength", bold_and_border)
            worksheet.merge_range("A2:B2", "Bandwidth", bold_and_border)
            worksheet.merge_range("A3:B3", "Open Aperture Diameter", bold_and_border)
            worksheet.merge_range("A4:B4", "Sweeping Distance", bold_and_border)
            worksheet.merge_range("A5:B5", "Step Distance", bold_and_border)

            worksheet.write("C1", f"{self.xml[0][0].text} nm", border)
            worksheet.write("C2", self.xml[0][1].text, border)
            worksheet.write("C3", f"{self.xml[1][6].text} mm", border)
            worksheet.write("C4", self.endvalueentry.get(), border)
            worksheet.write("C5", self.stepvalueentry.get(), border)

            worksheet.merge_range("A8:B8", "Open Aperture", bold_and_border_and_center)
            worksheet.write("A9", "Distance in mm", bold_and_border)
            worksheet.write("B9", "Power in W", bold_and_border)
            row = 9
            column = 0

            for item in self.without_aperture_x_data:
                worksheet.write(row, column, item, border)
                row += 1

            row = 9
            column = 1

            for item in self.without_aperture_y_data:
                worksheet.write(row, column, item, border)
                row += 1

            workbook.close()

    def with_aperture_excel(self):
        if self.get_file():
            file_path = filedialog.askdirectory(title="Select the folder to save")
            file_path = file_path + f"/{self.enter_file_name_entry.get()} close aperture.xlsx"
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            border = workbook.add_format({'border': 1})
            bold_and_border = workbook.add_format({'border': 1, 'bold': True})
            bold_and_border_and_center = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
            worksheet.set_column_pixels(0, 3, 100)

            worksheet.merge_range("A1:B1", "Wavelength", bold_and_border)
            worksheet.merge_range("A2:B2", "Bandwidth", bold_and_border)
            worksheet.merge_range("A3:B3", "Close Aperture Diameter", bold_and_border)
            worksheet.merge_range("A4:B4", "Sweeping Distance", bold_and_border)
            worksheet.merge_range("A5:B5", "Step Distance", bold_and_border)

            worksheet.write("C1", f"{self.xml[0][0].text} nm", border)
            worksheet.write("C2", self.xml[0][1].text, border)
            worksheet.write("C3", f"{self.xml[1][7].text} mm", border)
            worksheet.write("C4", self.endvalueentry.get(), border)
            worksheet.write("C5", self.stepvalueentry.get(), border)

            worksheet.merge_range("A8:B8", "Close Aperture", bold_and_border_and_center)
            worksheet.write("A9", "Distance in mm", bold_and_border)
            worksheet.write("B9", "Power in W", bold_and_border)
            row = 9
            column = 0

            for item in self.with_aperture_x_data:
                worksheet.write(row, column, item, border)
                row += 1

            row = 9
            column = 1

            for item in self.with_aperture_y_data:
                worksheet.write(row, column, item, border)
                row += 1

            workbook.close()

    def save_data_excel(self):
        if (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) == 0):
            self.without_aperture_excel()
        elif (len(self.without_aperture_x_data) == 0) and (len(self.with_aperture_x_data) != 0):
            self.with_aperture_excel()
        elif (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) != 0):
            self.excel_selector()

    def status_updater(self):
        if self.homeButton["text"] == "Home":
            self.homeButton["state"] = "disabled"
        if self.endButton["text"] == "End":
            self.endButton["state"] = "disabled"
        if self.run_button_without_aperture["text"] == "Start Run without Aperture":
            self.run_button_without_aperture["state"] = "disabled"
        if self.run_button_with_aperture["text"] == "Start Run with Aperture":
            self.run_button_with_aperture["state"] = "disabled"
            self.close_aperture_entry["state"] = "disabled"
        self.startvalueentry["state"] = "disabled"
        self.endvalueentry["state"] = "disabled"
        self.stepvalueentry["state"] = "disabled"
        self.com_btn_connect["state"] = "disabled"
        self.drop_startvalue["state"] = "disabled"
        self.drop_stopvalue["state"] = "disabled"
        self.drop_stepvalue["state"] = "disabled"

    def status_retainer(self):
        if self.homeButton["text"] == "Stop":
            self.homeButton["text"] = "Home"
        else:
            self.homeButton["state"] = "active"
        if self.endButton["text"] == "Stop":
            self.endButton["text"] = "End"
        else:
            self.endButton["state"] = "active"
        if self.run_button_without_aperture["text"] == "Stop Run":
            self.run_button_without_aperture["text"] = "Start Run without Aperture"
        else:
            self.run_button_without_aperture["state"] = "active"
        if self.run_button_with_aperture["text"] == "Stop Run":
            self.run_button_with_aperture["text"] = "Start Run with Aperture"
        else:
            self.run_button_with_aperture["state"] = "active"
            self.close_aperture_entry["state"] = "normal"
        self.startvalueentry["state"] = "normal"
        self.endvalueentry["state"] = "normal"
        self.stepvalueentry["state"] = "normal"
        self.com_btn_connect["state"] = "active"
        self.drop_startvalue["state"] = "normal"
        self.drop_stopvalue["state"] = "normal"
        self.drop_stepvalue["state"] = "normal"
        self.notebook.add(self.settingsframe, text="Settings")

    def Home(self):
        if self.homeButton["text"] == "Home":
            self.t1 = threading.Thread(target=self.homeThreading, daemon=True)
            self.t1.start()
            self.homeButton["text"] = "Stop"
            self.status_updater()
        else:
            self.start_Initiation = False
            self.notebook.add(self.settingsframe, text="Settings")
            self.status_retainer()

    def End(self):
        if self.endButton["text"] == "End":
            self.t1 = threading.Thread(target=self.endThreading, daemon=True)
            self.t1.start()
            self.endButton["text"] = "Stop"
            self.status_updater()
        else:
            self.start_Initiation = False
            self.notebook.add(self.settingsframe, text="Settings")
            self.status_retainer()

    def Run_with_aperture(self):
        if self.run_button_with_aperture["text"] == "Start Run with Aperture":
            if len(self.close_aperture_entry.get()) == 0:
                self.xml[1][7].text = " "
            else:
                self.xml[1][7].text = self.close_aperture_entry.get()
            tree.write("data_file.xml")
            self.value_show_threading = False
            sleep(0.1)
            self.start_Initiation = True
            self.t1 = threading.Thread(target=self.run_Threading, daemon=True)
            self.t1.start()
            self.run_button_with_aperture["text"] = "Stop Run"
            self.notebook.hide(1)
            self.status_updater()
            self.save_data_button["state"] = "disabled"
            self.enter_file_name_entry["state"] = "disabled"
            self.reset_all_button["state"] = "disabled"
            self.save_image_data_button["state"] = "disabled"
        else:
            self.start_Initiation = False

    def Run_without_aperture(self):
        if self.run_button_without_aperture["text"] == "Start Run without Aperture":
            self.value_show_threading = False
            sleep(0.1)
            self.start_Initiation = True
            self.t1 = threading.Thread(target=self.run_Threading, daemon=True)
            self.t1.start()
            self.run_button_without_aperture["text"] = "Stop Run"
            self.notebook.hide(1)
            self.status_updater()
            self.save_data_button["state"] = "disabled"
            self.enter_file_name_entry["state"] = "disabled"
            self.reset_all_button["state"] = "disabled"
            self.save_image_data_button["state"] = "disabled"
        else:
            self.start_Initiation = False

    def num_validator(self, received_data):
        self.data = received_data
        try:
            float(self.data)
            return True
        except:
            return False

    def values_validator(self, start, stop, step):
        self.start = start
        self.stop = stop
        self.step = step
        if len(self.start) > 0:
            if len(self.stop) > 0:
                if len(self.step) > 0:
                    if self.num_validator(self.start):
                        if self.num_validator(self.stop):
                            if self.num_validator(self.step):
                                return True
                            else:
                                messagebox.showerror("Value Error", "The Step value must be a number")
                                return False
                        else:
                            messagebox.showerror("Value Error", "The Ending Point value must be a number")
                            return False
                    else:
                        messagebox.showerror("Value Error", "The Starting Point value must be a number")
                        return False
                else:
                    messagebox.showerror("Value Error", "The Step value field should not be blank")
                    return False
            else:
                messagebox.showerror("Value Error", "The Ending Point field should not be blank")
                return False
        else:
            messagebox.showerror("Value Error", "The Starting Point value field should not be blank")
            return False

    def save_button(self):
        wavelength = self.wavelength_entry.get()
        thread_spacing = self.thread_spacing_entry.get()
        start = self.default_startvalue_entry.get()
        stop = self.default_stopvalue_entry.get()
        step = self.default_stepvalue_entry.get()
        if len(wavelength) > 0:
            if self.num_validator(wavelength):
                if float(wavelength) >= 400 and float(wavelength) <= 1100:
                    if len(thread_spacing) > 0:
                        if self.num_validator(thread_spacing):
                            if self.values_validator(start, stop, step):
                                self.xml[0][0].text = str(wavelength)
                                self.xml[0][1].text = self.clicked_bandwidth.get()
                                self.xml[0][2].text = str(thread_spacing)
                                self.xml[1][0].text = str(start)
                                self.xml[1][1].text = self.default_clicked_startvalue.get()
                                self.xml[1][2].text = str(stop)
                                self.xml[1][3].text = self.default_clicked_stopvalue.get()
                                self.xml[1][4].text = str(step)
                                self.xml[1][5].text = self.default_clicked_stepvalue.get()
                                if len(self.open_aperture_dia_entry.get()) == 0:
                                    self.xml[1][6].text = " "
                                else:
                                    self.xml[1][6].text = self.open_aperture_dia_entry.get()

                                tree.write("data_file.xml")
                                messagebox.showinfo("Save status", "Data saved successfully")
                        else:
                            messagebox.showerror("Value Error", "Thread Spacing value must be a number")
                    else:
                        messagebox.showerror("Value Error", "Thread Spacing value should not be blank")
                else:
                    messagebox.showerror("Value Error", "Wavelength value must in between 400 nm to 1100 nm")
            else:
                messagebox.showerror("Value Error", "Wavelength value must be a number")
        else:
            messagebox.showerror("Value Error", "Wavelength value should not be blank")

    def apply_button(self):
        self.wavelength = self.wavelength_entry.get()
        thread_spacing = self.thread_spacing_entry.get()
        if len(self.wavelength) > 0:
            if self.num_validator(self.wavelength):
                if float(self.wavelength) >= 400 and float(self.wavelength) <= 1100:
                    if len(thread_spacing) > 0:
                        if self.num_validator(thread_spacing):
                            self.wavelength_value = self.wavelength
                            if self.clicked_bandwidth.get() == "High":
                                self.bandwidth_value = "High"
                            elif self.clicked_bandwidth.get() == "Low":
                                self.bandwidth_value = "Low"
                            self.minimumvalue_text = float(self.thread_spacing_entry.get()) / self.motor_steps * 1000
                            self.minimumvalue.config(
                                text=f"Note: The minimum step value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                            messagebox.showinfo("Apply status", "Data applied successfully")
                        else:
                            messagebox.showerror("Value Error", "Thread Spacing value must be a number")
                    else:
                        messagebox.showerror("Value Error", "Thread Spacing value should not be blank")
                else:
                    messagebox.showerror("Value Error", "Wavelength value must in between 400 nm to 1100 nm")
            else:
                messagebox.showerror("Value Error", "Wavelength value must be a number")
        else:
            messagebox.showerror("Value Error", "Wavelength value should not be blank")

    def enable_buttons(self):
        self.homeButton["state"] = "active"
        self.endButton["state"] = "active"
        self.run_button_without_aperture["state"] = "active"
        self.run_button_with_aperture["state"] = "active"
        self.close_aperture_entry["state"] = "normal"
        self.startvalueentry["state"] = "normal"
        self.endvalueentry["state"] = "normal"
        self.stepvalueentry["state"] = "normal"
        self.drop_startvalue["state"] = "active"
        self.drop_stopvalue["state"] = "active"
        self.drop_stepvalue["state"] = "active"
        # self.save_data_button["state"] = "active"
        # self.reset_all_button["state"] = "active"
        # self.save_image_data_button["state"] = "active"

    def disable_buttons(self):
        self.homeButton["state"] = "disabled"
        self.endButton["state"] = "disabled"
        self.run_button_without_aperture["state"] = "disabled"
        self.run_button_with_aperture["state"] = "disabled"
        self.close_aperture_entry["state"] = "disabled"
        self.startvalueentry["state"] = "disabled"
        self.endvalueentry["state"] = "disabled"
        self.stepvalueentry["state"] = "disabled"
        self.drop_startvalue["state"] = "disable"
        self.drop_stopvalue["state"] = "disable"
        self.drop_stepvalue["state"] = "disable"
        self.save_data_button["state"] = "disabled"
        self.enter_file_name_entry["state"] = "disabled"
        self.reset_all_button["state"] = "disabled"
        self.save_image_data_button["state"] = "disabled"

    def publish(self):
        self.notebook.grid(row=0, column=0)
        self.notebook.add(self.mainframe, text="Controls")

        self.mainframe.grid(row=0, column=0)
        self.settingsframe.grid(row=0, column=0)

        self.connection_frame.grid(row=0, column=0, columnspan=3, rowspan=3, padx=5, pady=5)
        self.connection_frame.grid_propagate(False)

        self.notebook.add(self.mainframe, text="Controls")
        self.notebook.add(self.settingsframe, text="Settings")

        self.label_usb.grid(column=1, row=1)
        self.drop_usb.grid(column=2, row=1, padx=self.padx, pady=self.pady)
        self.label_com.grid(column=1, row=2)
        self.drop_com.grid(column=2, row=2, padx=self.padx, pady=self.pady)
        self.label_bd.grid(column=1, row=3)
        self.drop_bd.grid(column=2, row=3, padx=self.padx, pady=self.pady)

        self.btn_usb_refresh.grid(column=3, row=1)
        self.btn_refresh.grid(column=3, row=2)
        self.com_btn_connect.grid(column=3, row=3)

        self.communication_frame.grid_propagate(False)
        self.communication_frame.grid(row=0, column=4, rowspan=3, columnspan=5, padx=5, pady=5)

        self.homeButton.grid(row=0, column=4, pady=5, padx=(10, 0))
        self.endButton.grid(row=1, column=4, pady=5, padx=(10, 0))
        # self.run_button_without_aperture.grid(row=2, column=4, pady=5, padx=(10, 0))
        self.startvalue.grid(row=0, column=0)
        self.endvalue.grid(row=1, column=0)
        self.stepvalue.grid(row=2, column=0)
        self.startvalueentry.grid(row=0, column=1)
        self.endvalueentry.grid(row=1, column=1)
        self.stepvalueentry.grid(row=2, column=1)
        self.drop_startvalue.grid(row=0, column=2, padx=15)
        self.drop_stopvalue.grid(row=1, column=2, padx=15)
        self.drop_stepvalue.grid(row=2, column=2, padx=15)
        self.seperator.place(relx=0.67, rely=0, relwidth=0.001, relheight=0.8)

        self.graph_frame.grid_propagate(False)
        self.graph_frame.grid(row=3, column=0, columnspan=8, padx=5, pady=5)
        self.graph_buttons_frame.grid(row=0, column=1, padx=(15, 0))
        self.graph_buttons_frame.grid_propagate(False)
        self.power_label.grid(row=0, column=1)
        self.power_value.grid(row=1, column=1)
        self.run_button_without_aperture.grid(row=2, column=1, pady=(20, 7.4))
        self.enter_close_aperture_label.grid(row=3, column=1, pady=2)
        self.close_aperture_entry.grid(row=4, column=1, pady=2)
        self.run_button_with_aperture.grid(row=5, column=1, pady=7.4)
        self.enter_file_name_label.grid(row=6, column=1, pady=2)
        self.enter_file_name_entry.grid(row=7, column=1, pady=2)
        self.save_data_button.grid(row=8, column=1, pady=7.4)
        self.save_image_data_button.grid(row=9, column=1, pady=7.4)
        self.reset_all_button.grid(row=10, column=1, pady=2)
        self.seperator2.place(relx=0.83, rely=-0.04, relwidth=0.001, relheight=1.055)
        self.seperator3.place(relx=0.83, rely=0.18, relwidth=0.174, relheight=0.001)
        self.seperator4.place(relx=0.83, rely=0.297, relwidth=0.174, relheight=0.001)
        self.seperator5.place(relx=0.83, rely=0.547, relwidth=0.174, relheight=0.001)
        self.seperator6.place(relx=0.83, rely=0.793, relwidth=0.174, relheight=0.001)
        # self.seperator7.place(relx=0.83, rely=0.896, relwidth=0.174, relheight=0.001)

        self.Config.grid(row=0, column=0, rowspan=2, padx=5, pady=5)
        self.Config.grid_propagate(False)

        self.wavelength.grid(row=0, column=0)
        self.bandwidth.grid(row=1, column=0)
        self.thread_spacing.grid(row=2, column=0)
        self.open_aperture_dia.grid(row=3, column=0)
        self.default_startvalue.grid(row=4, column=0)
        self.default_stopvalue.grid(row=5, column=0)
        self.default_stepvalue.grid(row=6, column=0)

        self.wavelength_entry.grid(row=0, column=1, padx=(0, 7))
        self.drop_bandwidth.grid(row=1, column=1, padx=(0, 7))
        self.thread_spacing_entry.grid(row=2, column=1, padx=(0, 7))
        self.open_aperture_dia_entry.grid(row=3, column=1, padx=(0, 7))
        self.default_startvalue_entry.grid(row=4, column=1, padx=(0, 7))
        self.default_stopvalue_entry.grid(row=5, column=1, padx=(0, 7))
        self.default_stepvalue_entry.grid(row=6, column=1, padx=(0, 7))

        self.wavelength_unit.grid(row=0, column=2)
        self.open_aperture_unit.grid(row=3, column=2)
        self.thread_spacing_unit.grid(row=2, column=2)
        self.default_drop_startvalue.grid(row=4, column=2)
        self.default_drop_stopvalue.grid(row=5, column=2)
        self.default_drop_stepvalue.grid(row=6, column=2)

        self.save_button.grid(row=0, column=1, padx=5)
        self.apply_button.grid(row=0, column=2, padx=5)

        self.button_frame.grid(row=7, column=0, columnspan=3, pady=20)

        self.motor_control.grid(row=0, column=1, padx=5, pady=5)
        self.motor_control.grid_propagate(False)
        self.manual_button_frame.grid(row=0, column=0, columnspan=6)
        self.enable_manual_mode.grid(row=1, column=0, padx=5, pady=(0, 15))
        self.enable_motor_button.grid(row=1, column=1, padx=5, pady=(0, 15))
        self.manual_power_label.grid(row=0, column=0, columnspan=2, pady=(5, 10))
        self.move_label.grid(row=1, column=0, padx=5, pady=5)
        self.towards_label.grid(row=1, column=3, padx=5, pady=5)
        self.direction_label.grid(row=1, column=5, padx=5, pady=5)
        self.desired_distance.grid(row=1, column=1, padx=5, pady=5)
        self.drop_desired_unit.grid(row=1, column=2, padx=5, pady=5)
        self.drop_desired_direction.grid(row=1, column=4, padx=5, pady=5)
        self.move_buttons_frame.grid(row=2, column=0, columnspan=6)
        self.cw_button.grid(row=0, column=2)
        self.move_button.grid(row=0, column=1)
        self.acw_button.grid(row=0, column=0)

        self.position_finder_frame.grid(row=1, column=1, padx=5, pady=5)
        self.position_finder_frame.grid_propagate(False)
        self.find_position.grid(row=0, column=0, padx=5, pady=5)
        self.holder_position.grid(row=1, column=0, padx=5, pady=10)

        self.minimumvalue.grid(row=3, column=0, columnspan=5, pady=(5, 0))
        self.disable_buttons()


if __name__ == "__main__":
    RootGUI()
    CommGUI()
