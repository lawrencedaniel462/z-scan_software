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


# bg = button_background, fg = button_foreground, bd = button_border_size, highlightbackground = button_highlight_background, highlightcolor = button_highlight_foreground, disabledforeground = button_disabled_foreground
class CommGUI:
    def __init__(self, root, usb, style):
        self.xml = tree.getroot()

        self.motor_steps = float(self.xml[1][9].text)
        self.root = root
        self.usb = usb
        self.styles = style

        self.style = ttk.Style()

        self.style.theme_create("custom_theme", parent="alt", settings={
            "TLabelframe": {
                "configure": {  # "padx": 5,
                    # "pady": 5,
                    "background": "#FFE5B4",
                    "bordercolor": "black",#"#CD7F32",
                    "borderwidth": 13,
                    "darkcolor": "#CD7F32",
                    "labelmargins": 1,
                    "labeloutside": 1,
                    "lightcolor": "#FFE5B4"}
            },
            "TLabelframe.Label": {
                "configure": {"background": self.styles.bg_color}
            },
            "TSeparator": {
                "configure": {"background": "#CD7F32"}
            },
            "TNotebook": {
                "configure": {"background": self.styles.bg_color}
            },
            "TNotebook.Tab": {
                "configure": {"background": self.styles.bg_color},
                "map": {"background": [("selected", self.styles.bg_color), ('!active', self.styles.light_bg_color), ('active', "#FAC898")]}
            }
        })
        self.style.configure("TLabelframe",
                             background="#FFE5B4",
                             bordercolor="#CD7F32",
                             borderwidth=13,
                             darkcolor="#CD7F32",
                             labelmargins=1,
                             labeloutside=1,
                             lightcolor="#FFE5B4")
        self.style.theme_use("custom_theme")

        self.notebook = ttk.Notebook(root)

        self.mainframe = Frame(self.notebook,
                               bg=self.styles.bg_color)
        self.settingsframe = Frame(self.notebook,
                                   bg=self.styles.bg_color)

        self.Config = LabelFrame(self.settingsframe,
                                 text="Configuration",
                                 padx=5,
                                 pady=5,
                                 width=880,
                                 height=260,
                                 bg=self.styles.bg_color)

        self.wavelength_entry = Entry(self.Config,
                                      width=13,
                                      font=('Helvetica', 10),
                                      bg=self.styles.entry_background_color,
                                      highlightbackground=self.styles.entry_border_color,
                                      highlightthickness=self.styles.entry_border_size,
                                      highlightcolor=self.styles.active_entry_border_color,
                                      disabledbackground=self.styles.disabled_background_entry,
                                      disabledforeground=self.styles.disabled_foreground_entry)
        self.wavelength_entry.insert(0, self.xml[0][0].text)

        self.thread_spacing_entry = Entry(self.Config,
                                          width=13,
                                          font=('Helvetica', 10),
                                          bg=self.styles.entry_background_color,
                                          highlightbackground=self.styles.entry_border_color,
                                          highlightthickness=self.styles.entry_border_size,
                                          highlightcolor=self.styles.active_entry_border_color,
                                          disabledbackground=self.styles.disabled_background_entry,
                                          disabledforeground=self.styles.disabled_foreground_entry)
        self.thread_spacing_entry.insert(0, self.xml[0][2].text)

        self.pulse_width_entry = Entry(self.Config,
                                       width=13,
                                       font=('Helvetica', 10),
                                       bg=self.styles.entry_background_color,
                                       highlightbackground=self.styles.entry_border_color,
                                       highlightthickness=self.styles.entry_border_size,
                                       highlightcolor=self.styles.active_entry_border_color,
                                       disabledbackground=self.styles.disabled_background_entry,
                                       disabledforeground=self.styles.disabled_foreground_entry)
        self.pulse_width_entry.insert(0, self.xml[1][8].text)

        self.steps_per_rotation_entry = Entry(self.Config,
                                              width=13,
                                              font=('Helvetica', 10),
                                              bg=self.styles.entry_background_color,
                                              highlightbackground=self.styles.entry_border_color,
                                              highlightthickness=self.styles.entry_border_size,
                                              highlightcolor=self.styles.active_entry_border_color,
                                              disabledbackground=self.styles.disabled_background_entry,
                                              disabledforeground=self.styles.disabled_foreground_entry)
        self.steps_per_rotation_entry.insert(0, self.xml[1][9].text)

        self.open_aperture_dia_entry = Entry(self.Config,
                                             width=13,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.open_aperture_dia_entry.insert(0, self.xml[1][6].text)

        self.close_aperture_entry = Entry(self.Config,
                                          width=13,
                                          font=('Helvetica', 10),
                                          bg=self.styles.entry_background_color,
                                          highlightbackground=self.styles.entry_border_color,
                                          highlightthickness=self.styles.entry_border_size,
                                          highlightcolor=self.styles.active_entry_border_color,
                                          disabledbackground=self.styles.disabled_background_entry,
                                          disabledforeground=self.styles.disabled_foreground_entry)
        self.close_aperture_entry.insert(0, self.xml[1][7].text)

        self.default_startvalue_entry = Entry(self.Config,
                                              width=13,
                                              font=('Helvetica', 10),
                                              bg=self.styles.entry_background_color,
                                              highlightbackground=self.styles.entry_border_color,
                                              highlightthickness=self.styles.entry_border_size,
                                              highlightcolor=self.styles.active_entry_border_color,
                                              disabledbackground=self.styles.disabled_background_entry,
                                              disabledforeground=self.styles.disabled_foreground_entry)
        self.default_startvalue_entry.insert(0, self.xml[1][0].text)

        self.default_stopvalue_entry = Entry(self.Config,
                                             width=13,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.default_stopvalue_entry.insert(0, self.xml[1][2].text)

        self.default_stepvalue_entry = Entry(self.Config,
                                             width=13,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.default_stepvalue_entry.insert(0, self.xml[1][4].text)

        self.wavelength = Label(self.Config, text="Wavelength", anchor="w", width=25, padx=7, pady=8, bg=self.styles.bg_color)
        self.bandwidth = Label(self.Config, text="Bandwidth", anchor="w", width=25, padx=7, pady=8, bg=self.styles.bg_color)
        self.thread_spacing = Label(self.Config, text="Thread Spacing", anchor="w", width=25, padx=7, pady=8,
                                    bg=self.styles.bg_color)
        self.pulse_width_label = Label(self.Config, text="Pulse width", anchor="w", width=25, padx=7, pady=8,
                                       bg=self.styles.bg_color)
        self.steps_per_rotation_label = Label(self.Config, text="No of steps per rotation", anchor="w", width=25,
                                              padx=7, pady=8, bg=self.styles.bg_color)
        self.open_aperture_dia = Label(self.Config, text="Open Aperture Dia", anchor="w", width=35, padx=7, pady=8,
                                       bg=self.styles.bg_color)
        self.enter_close_aperture_label = Label(self.Config, text="Close Aperture Dia",
                                                anchor="w", width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.default_startvalue = Label(self.Config, text="Default distance from Home to Maxima", anchor="w", width=35,
                                        padx=7, pady=8, bg=self.styles.bg_color)
        self.default_stopvalue = Label(self.Config, text="Default Half Sweeping Distance", anchor="w", width=35, padx=7,
                                       pady=8, bg=self.styles.bg_color)
        self.default_stepvalue = Label(self.Config, text="Default Step Value", anchor="w", width=35, padx=7, pady=8,
                                       bg=self.styles.bg_color)

        self.wavelength_unit = Label(self.Config, text="nm", anchor="w", width=12, bg=self.styles.bg_color)
        self.thread_spacing_unit = Label(self.Config, text="mm", anchor="w", width=12, bg=self.styles.bg_color)
        self.pulse_width_unit = Label(self.Config, text="µs", anchor="w", width=12, bg=self.styles.bg_color)
        self.open_aperture_unit = Label(self.Config, text="mm", anchor="w", width=12, bg=self.styles.bg_color)
        self.close_aperture_unit = Label(self.Config, text="mm", anchor="w", width=12, bg=self.styles.bg_color)

        self.seperator2 = ttk.Separator(self.Config, orient="horizontal")

        self.connection_frame = LabelFrame(self.mainframe,
                                           text="Communication Manager",
                                           padx=5,
                                           pady=5,
                                           width=520,
                                           height=165,
                                           bg=self.styles.bg_color)
        self.label_usb = Label(self.connection_frame, text="Available USB Port(s): ", width=17,
                               anchor="w", bg=self.styles.bg_color)
        self.label_com = Label(self.connection_frame, text="Available COM Port(s): ", width=17,
                               anchor="w", bg=self.styles.bg_color)
        self.label_bd = Label(self.connection_frame, text="Baud Rate: ", width=17, anchor="w", bg=self.styles.bg_color)
        self.btn_refresh = Button(self.connection_frame,
                                  text="Refresh",
                                  width=9,
                                  command=self.Comm_Refresh,
                                  bg=self.styles.button_background,
                                  fg=self.styles.button_foreground,
                                  bd=self.styles.button_border_size,
                                  highlightbackground=self.styles.button_highlight_background,
                                  highlightcolor=self.styles.button_highlight_foreground,
                                  disabledforeground=self.styles.button_disabled_foreground,
                                  activebackground=self.styles.button_active_background,
                                  activeforeground=self.styles.button_active_foreground,
                                  font=self.styles.bold_font)
        self.btn_usb_refresh = Button(self.connection_frame,
                                      text="Refresh",
                                      width=9,
                                      command=self.USB_Refresh,
                                      bg=self.styles.button_background,
                                      fg=self.styles.button_foreground,
                                      bd=self.styles.button_border_size,
                                      highlightbackground=self.styles.button_highlight_background,
                                      highlightcolor=self.styles.button_highlight_foreground,
                                      disabledforeground=self.styles.button_disabled_foreground,
                                      activebackground=self.styles.button_active_background,
                                      activeforeground=self.styles.button_active_foreground,
                                      font=self.styles.bold_font)
        self.com_btn_connect = Button(self.connection_frame,
                                      text="Connect",
                                      width=9,
                                      state="disabled",
                                      command=self.SerialConnect,
                                      bg=self.styles.button_background,
                                      fg=self.styles.button_foreground,
                                      bd=self.styles.button_border_size,
                                      highlightbackground=self.styles.button_highlight_background,
                                      highlightcolor=self.styles.button_highlight_foreground,
                                      disabledforeground=self.styles.button_disabled_foreground,
                                      activebackground=self.styles.button_active_background,
                                      activeforeground=self.styles.button_active_foreground,
                                      font=self.styles.bold_font)
        self.com_option_menu()
        self.BaudOptionMenu()
        self.USBOptionMenu()

        self.communication_frame = LabelFrame(self.mainframe, text="Controls Manager", pady=5, padx=5, width=650,
                                              height=165, bg=self.styles.bg_color)
        self.home_button = Button(self.communication_frame,
                                  text="Home", width=23,
                                  command=self.Home,
                                  bg=self.styles.button_background,
                                  fg=self.styles.button_foreground,
                                  bd=self.styles.button_border_size,
                                  highlightbackground=self.styles.button_highlight_background,
                                  highlightcolor=self.styles.button_highlight_foreground,
                                  disabledforeground=self.styles.button_disabled_foreground,
                                  activebackground=self.styles.button_active_background,
                                  activeforeground=self.styles.button_active_foreground,
                                  font=self.styles.bold_font)
        self.end_button = Button(self.communication_frame,
                                 text="End",
                                 width=23,
                                 command=self.End,
                                 bg=self.styles.button_background,
                                 fg=self.styles.button_foreground,
                                 bd=self.styles.button_border_size,
                                 highlightbackground=self.styles.button_highlight_background,
                                 highlightcolor=self.styles.button_highlight_foreground,
                                 disabledforeground=self.styles.button_disabled_foreground,
                                 activebackground=self.styles.button_active_background,
                                 activeforeground=self.styles.button_active_foreground,
                                 font=self.styles.bold_font)
        self.startvalue = Label(self.communication_frame, text="Distance from Home to Maxima", anchor="w", width=27,
                                bg=self.styles.bg_color)
        self.endvalue = Label(self.communication_frame, text="Half Sweeping Distance", anchor="w", width=27,
                              bg=self.styles.bg_color)
        self.stepvalue = Label(self.communication_frame, text="Step Value", anchor="w", width=27, bg=self.styles.bg_color)

        self.startvalueentry = Entry(self.communication_frame,
                                     width=15,
                                     bg=self.styles.entry_background_color,
                                     highlightbackground=self.styles.entry_border_color,
                                     highlightthickness=self.styles.entry_border_size,
                                     highlightcolor=self.styles.active_entry_border_color,
                                     disabledbackground=self.styles.disabled_background_entry,
                                     disabledforeground=self.styles.disabled_foreground_entry)
        self.startvalueentry.insert(0, self.default_startvalue_entry.get())
        self.endvalueentry = Entry(self.communication_frame,
                                   width=15,
                                   bg=self.styles.entry_background_color,
                                   highlightbackground=self.styles.entry_border_color,
                                   highlightthickness=self.styles.entry_border_size,
                                   highlightcolor=self.styles.active_entry_border_color,
                                   disabledbackground=self.styles.disabled_background_entry,
                                   disabledforeground=self.styles.disabled_foreground_entry)
        self.endvalueentry.insert(0, self.default_stopvalue_entry.get())
        self.stepvalueentry = Entry(self.communication_frame,
                                    width=15,
                                    bg=self.styles.entry_background_color,
                                    highlightbackground=self.styles.entry_border_color,
                                    highlightthickness=self.styles.entry_border_size,
                                    highlightcolor=self.styles.active_entry_border_color,
                                    disabledbackground=self.styles.disabled_background_entry,
                                    disabledforeground=self.styles.disabled_foreground_entry)
        self.stepvalueentry.insert(0, self.default_stepvalue_entry.get())

        self.minimumvalue_text = float(self.thread_spacing_entry.get()) / self.motor_steps * 1000
        self.minimumvalue = Label(self.communication_frame,
                                  text=f"Note: The minimum step value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm",
                                  bg=self.styles.bg_color)
        self.seperator = ttk.Separator(self.communication_frame, orient="vertical")

        self.graph_frame = LabelFrame(self.mainframe, text="Graph Manager", pady=5, padx=5, width=1180, height=385,
                                      bg=self.styles.bg_color)

        self.graph_buttons_frame = Frame(self.graph_frame, width=190, height=350, bg=self.styles.bg_color)
        self.power_label = Label(self.graph_buttons_frame, text="Power in W", anchor="center", width=26, bg=self.styles.bg_color)
        self.power_value = Label(self.graph_buttons_frame, anchor="center", width=12, foreground="red", text="-",
                                 font=('Times New Roman', 18, 'bold'), bg=self.styles.bg_color)

        self.run_button_without_aperture = Button(self.graph_buttons_frame,
                                                  text="Start Run without Aperture",
                                                  width=21,
                                                  command=self.Run_without_aperture,
                                                  bg=self.styles.button_background,
                                                  fg=self.styles.button_foreground,
                                                  bd=self.styles.button_border_size,
                                                  highlightbackground=self.styles.button_highlight_background,
                                                  highlightcolor=self.styles.button_highlight_foreground,
                                                  disabledforeground=self.styles.button_disabled_foreground,
                                                  activebackground=self.styles.button_active_background,
                                                  activeforeground=self.styles.button_active_foreground,
                                                  font=self.styles.bold_font)
        self.run_button_with_aperture = Button(self.graph_buttons_frame,
                                               text="Start Run with Aperture",
                                               width=21,
                                               command=self.Run_with_aperture,
                                               bg=self.styles.button_background,
                                               fg=self.styles.button_foreground,
                                               bd=self.styles.button_border_size,
                                               highlightbackground=self.styles.button_highlight_background,
                                               highlightcolor=self.styles.button_highlight_foreground,
                                               disabledforeground=self.styles.button_disabled_foreground,
                                               activebackground=self.styles.button_active_background,
                                               activeforeground=self.styles.button_active_foreground,
                                               font=self.styles.bold_font)
        self.enter_file_name_label = Label(self.graph_buttons_frame, text="Enter file name", width=21, anchor="center",
                                           bg=self.styles.bg_color)
        self.enter_file_name_entry = Entry(self.graph_buttons_frame,
                                           width=25,
                                           bg=self.styles.entry_background_color,
                                           highlightbackground=self.styles.entry_border_color,
                                           highlightthickness=self.styles.entry_border_size,
                                           highlightcolor=self.styles.active_entry_border_color,
                                           disabledbackground=self.styles.disabled_background_entry,
                                           disabledforeground=self.styles.disabled_foreground_entry)
        self.save_data_button = Button(self.graph_buttons_frame,
                                       text="Save Data",
                                       width=21,
                                       command=self.save_data_excel,
                                       bg=self.styles.button_background,
                                       fg=self.styles.button_foreground,
                                       bd=self.styles.button_border_size,
                                       highlightbackground=self.styles.button_highlight_background,
                                       highlightcolor=self.styles.button_highlight_foreground,
                                       disabledforeground=self.styles.button_disabled_foreground,
                                       activebackground=self.styles.button_active_background,
                                       activeforeground=self.styles.button_active_foreground,
                                       font=self.styles.bold_font)
        self.save_image_data_button = Button(self.graph_buttons_frame,
                                             text="Save as image",
                                             width=21,
                                             command=self.save_as_image,
                                             bg=self.styles.button_background,
                                             fg=self.styles.button_foreground,
                                             bd=self.styles.button_border_size,
                                             highlightbackground=self.styles.button_highlight_background,
                                             highlightcolor=self.styles.button_highlight_foreground,
                                             disabledforeground=self.styles.button_disabled_foreground,
                                             activebackground=self.styles.button_active_background,
                                             activeforeground=self.styles.button_active_foreground,
                                             font=self.styles.bold_font)
        self.reset_all_button = Button(self.graph_buttons_frame,
                                       text="Reset all",
                                       width=21,
                                       command=self.reset_all,
                                       bg=self.styles.button_background,
                                       fg=self.styles.button_foreground,
                                       bd=self.styles.button_border_size,
                                       highlightbackground=self.styles.button_highlight_background,
                                       highlightcolor=self.styles.button_highlight_foreground,
                                       disabledforeground=self.styles.button_disabled_foreground,
                                       activebackground=self.styles.button_active_background,
                                       activeforeground=self.styles.button_active_foreground,
                                       font=self.styles.bold_font)
        self.seperator3 = ttk.Separator(self.graph_frame, orient="vertical")
        self.seperator4 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator5 = ttk.Separator(self.graph_frame, orient="horizontal")
        self.seperator6 = ttk.Separator(self.graph_frame, orient="horizontal")
        # self.seperator7 = ttk.Separator(self.graph_frame, orient="horizontal")
        # self.seperator8 = ttk.Separator(self.graph_frame, orient="horizontal")

        self.initialize_graph()
        self.steps_remaining = Label(self.graph_frame, text="", anchor="w", width=17, bg=self.styles.bg_color)
        self.steps_completed = Label(self.graph_frame, text="", anchor="w", width=18, bg=self.styles.bg_color)
        self.total_step = Label(self.graph_frame, text="", anchor="w", width=13, bg=self.styles.bg_color)

        self.steps_remaining_value = Label(self.graph_frame, text="", anchor="w", width=17, foreground="Red",
                                     bg=self.styles.bg_color)
        self.steps_completed_value = Label(self.graph_frame, text="", anchor="w", width=18, foreground="Red",
                                     bg=self.styles.bg_color)
        self.total_step_value = Label(self.graph_frame, text="", anchor="w", width=13, foreground="Red", bg=self.styles.bg_color)

        self.current_position = 0

        self.value_option_menu()

        self.button_frame = Frame(self.Config, bg=self.styles.bg_color)

        self.save_button = Button(self.button_frame,
                                  text="Save",
                                  width=26,
                                  command=self.save_button,
                                  bg=self.styles.button_background,
                                  fg=self.styles.button_foreground,
                                  bd=self.styles.button_border_size,
                                  highlightbackground=self.styles.button_highlight_background,
                                  highlightcolor=self.styles.button_highlight_foreground,
                                  disabledforeground=self.styles.button_disabled_foreground,
                                  activebackground=self.styles.button_active_background,
                                  activeforeground=self.styles.button_active_foreground,
                                  font=self.styles.bold_font)
        self.apply_button = Button(self.button_frame,
                                   text="Apply",
                                   width=26,
                                   command=self.apply_button,
                                   bg=self.styles.button_background,
                                   fg=self.styles.button_foreground,
                                   bd=self.styles.button_border_size,
                                   highlightbackground=self.styles.button_highlight_background,
                                   highlightcolor=self.styles.button_highlight_foreground,
                                   disabledforeground=self.styles.button_disabled_foreground,
                                   activebackground=self.styles.button_active_background,
                                   activeforeground=self.styles.button_active_foreground,
                                   font=self.styles.bold_font)

        self.initialize_motor_control()
        self.initialize_position_finder()

        self.detector = object()
        self.publish()
        self.tester = 0
        self.overall_threading = False
        self.value_show_threading = False
        self.run_threading_status = False
        self.com_list = []
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
        self.enableCommand = "E\n"
        self.disableCommand = "D\n"
        self.clockwiseCommand = "C\n"
        self.antiClockwiseCommand = "A\n"
        self.pulseCommand = "P\n"
        self.homeCommand = "H\n"
        self.endCommand = "N\n"
        self.locateCommand = "L\n"
        self.stopCommand = "S\n"
        self.enable_ok = "e\r\n"
        self.disable_ok = "d\r\n"
        self.clockwise_ok = "c\r\n"
        self.antiClockwise_ok = "a\r\n"
        self.pulse_ok = "p\r\n"
        self.home_ok = "h\r\n"
        self.end_ok = "n\r\n"
        self.stop_ok = "s\r\n"
        self.move_ok = "m\r\n"

    def initialize_motor_control(self):
        self.motor_control = LabelFrame(self.settingsframe, text="Manual Control", padx=5, pady=5, width=470,
                                        height=198, bg=self.styles.bg_color)
        self.manual_button_frame = Frame(self.motor_control, bg=self.styles.bg_color)
        self.manual_power_label = Label(self.manual_button_frame, text="-", anchor="center", width=15, foreground="red",
                                        font=('Times New Roman', 18, 'bold'), bg=self.styles.bg_color)
        self.enable_manual_mode = Button(self.manual_button_frame,
                                         text="Enable manual mode",
                                         width=25,
                                         command=self.manual_mode_enable,
                                         bg=self.styles.button_background,
                                         fg=self.styles.button_foreground,
                                         bd=self.styles.button_border_size,
                                         highlightbackground=self.styles.button_highlight_background,
                                         highlightcolor=self.styles.button_highlight_foreground,
                                         disabledforeground=self.styles.button_disabled_foreground,
                                         activebackground=self.styles.button_active_background,
                                         activeforeground=self.styles.button_active_foreground,
                                         font=self.styles.bold_font)
        self.enable_motor_button = Button(self.manual_button_frame,
                                          text="Disable Motor",
                                          width=25,
                                          command=self.enable_motor,
                                          bg=self.styles.button_background,
                                          fg=self.styles.button_foreground,
                                          bd=self.styles.button_border_size,
                                          highlightbackground=self.styles.button_highlight_background,
                                          highlightcolor=self.styles.button_highlight_foreground,
                                          disabledforeground=self.styles.button_disabled_foreground,
                                          activebackground=self.styles.button_active_background,
                                          activeforeground=self.styles.button_active_foreground,
                                          font=self.styles.bold_font)
        self.move_label = Label(self.motor_control, text="Move", bg=self.styles.bg_color)
        self.towards_label = Label(self.motor_control, text="towards", bg=self.styles.bg_color)
        self.direction_label = Label(self.motor_control, text="direction", bg=self.styles.bg_color)
        self.desired_distance = Entry(self.motor_control,
                                      width=10,
                                      bg=self.styles.entry_background_color,
                                      highlightbackground=self.styles.entry_border_color,
                                      highlightthickness=self.styles.entry_border_size,
                                      highlightcolor=self.styles.active_entry_border_color,
                                      disabledbackground=self.styles.disabled_background_entry,
                                      disabledforeground=self.styles.disabled_foreground_entry)
        options = ["mm", "microns"]
        direction = ["home", "end"]
        self.clicked_desired_unit = StringVar()
        self.clicked_desired_direction = StringVar()
        self.clicked_desired_unit.set(options[0])
        self.clicked_desired_direction.set(direction[0])
        self.drop_desired_unit = OptionMenu(self.motor_control, self.clicked_desired_unit, *options)
        self.drop_desired_direction = OptionMenu(self.motor_control, self.clicked_desired_direction, *direction)
        self.drop_desired_unit.config(width=10,
                                      bg=self.styles.drop_background_color,
                                      activebackground=self.styles.active_drop_border_color,
                                      highlightbackground=self.styles.highlight_drop_color,
                                      highlightthickness=self.styles.drop_border_size,
                                      disabledforeground=self.styles.disabled_foreground_drop_color)
        self.drop_desired_direction.config(width=7,
                                           bg=self.styles.drop_background_color,
                                           activebackground=self.styles.active_drop_border_color,
                                           highlightbackground=self.styles.highlight_drop_color,
                                           highlightthickness=self.styles.drop_border_size,
                                           disabledforeground=self.styles.disabled_foreground_drop_color)
        self.move_buttons_frame = Frame(self.motor_control, bg=self.styles.bg_color)
        self.cw_button = Button(self.move_buttons_frame,
                                text=">>",
                                width=7,
                                command=self.clockwise_step,
                                bg=self.styles.button_background,
                                fg=self.styles.button_foreground,
                                bd=self.styles.button_border_size,
                                highlightbackground=self.styles.button_highlight_background,
                                highlightcolor=self.styles.button_highlight_foreground,
                                disabledforeground=self.styles.button_disabled_foreground,
                                activebackground=self.styles.button_active_background,
                                activeforeground=self.styles.button_active_foreground,
                                font=self.styles.bold_font)
        self.move_button = Button(self.move_buttons_frame,
                                  text="Move",
                                  width=37,
                                  command=self.move,
                                  bg=self.styles.button_background,
                                  fg=self.styles.button_foreground,
                                  bd=self.styles.button_border_size,
                                  highlightbackground=self.styles.button_highlight_background,
                                  highlightcolor=self.styles.button_highlight_foreground,
                                  disabledforeground=self.styles.button_disabled_foreground,
                                  activebackground=self.styles.button_active_background,
                                  activeforeground=self.styles.button_active_foreground,
                                  font=self.styles.bold_font)
        self.acw_button = Button(self.move_buttons_frame,
                                 text="<<",
                                 width=7,
                                 command=self.anticlockwise_step,
                                 bg=self.styles.button_background,
                                 fg=self.styles.button_foreground,
                                 bd=self.styles.button_border_size,
                                 highlightbackground=self.styles.button_highlight_background,
                                 highlightcolor=self.styles.button_highlight_foreground,
                                 disabledforeground=self.styles.button_disabled_foreground,
                                 activebackground=self.styles.button_active_background,
                                 activeforeground=self.styles.button_active_foreground,
                                 font=self.styles.bold_font)

        self.enable_motor_button["state"] = "disabled"
        self.cw_button["state"] = "disabled"
        self.move_button["state"] = "disabled"
        self.acw_button["state"] = "disabled"
        self.desired_distance["state"] = "disabled"
        self.drop_desired_unit["state"] = "disabled"
        self.drop_desired_direction["state"] = "disabled"

    def initialize_position_finder(self):
        self.position_finder_frame = LabelFrame(self.settingsframe, text="Position Locator", padx=5, pady=5, width=400,
                                                height=198, bg=self.styles.bg_color)
        self.find_position = Button(self.position_finder_frame,
                                    text="Locate the position",
                                    width=45,
                                    command=self.locate,
                                    bg=self.styles.button_background,
                                    fg=self.styles.button_foreground,
                                    bd=self.styles.button_border_size,
                                    highlightbackground=self.styles.button_highlight_background,
                                    highlightcolor=self.styles.button_highlight_foreground,
                                    disabledforeground=self.styles.button_disabled_foreground,
                                    activebackground=self.styles.button_active_background,
                                    activeforeground=self.styles.button_active_foreground,
                                    font=self.styles.bold_font)
        self.holder_position = Label(self.position_finder_frame, bg=self.styles.bg_color)
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
            if self.home_button["text"] == "Home" and self.end_button["text"] == "End" and \
                    self.run_button_without_aperture[
                        "text"] == "Start Run without Aperture" and self.enable_manual_mode[
                        "text"] == "Enable manual mode":
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
            self.sendCommand(self.stopCommand)

    def manual_mode_enable(self):
        if self.enable_manual_mode["text"] == "Enable manual mode":
            if self.home_button["text"] == "Home" and self.end_button["text"] == "End" and \
                    self.run_button_without_aperture[
                        "text"] == "Start Run without Aperture" and self.find_position["text"] == "Locate the position":
                self.enable_manual_mode["text"] = "Disable manual mode"
                self.enable_motor_button["state"] = "normal"
                self.cw_button["state"] = "normal"
                self.move_button["state"] = "normal"
                self.acw_button["state"] = "normal"
                self.desired_distance["state"] = "normal"
                self.drop_desired_unit["state"] = "normal"
                self.drop_desired_direction["state"] = "normal"
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
                self.save_button["state"] = "normal"
                self.apply_button["state"] = "normal"
                self.find_position["state"] = "normal"
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
                    else:
                        messagebox.showerror("Value Error", "Distance field should be a number")
                else:
                    messagebox.showerror("Value Error", "Distance field should not be blank")
            else:
                messagebox.showerror("Motor Status Error", "First enable motor to move")
        else:
            self.sendCommand(self.stopCommand)

    def initialize_graph(self):
        self.fig = Figure(figsize=(12, 4.4), dpi=80, facecolor=self.styles.bg_color)
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

        self.drop_bandwidth.config(width=9,
                                   bg=self.styles.drop_background_color,
                                   activebackground=self.styles.active_drop_border_color,
                                   highlightbackground=self.styles.highlight_drop_color,
                                   highlightthickness=self.styles.drop_border_size,
                                   disabledforeground=self.styles.disabled_foreground_drop_color)
        self.drop_startvalue.config(width=9,
                                    bg=self.styles.drop_background_color,
                                    activebackground=self.styles.active_drop_border_color,
                                    highlightbackground=self.styles.highlight_drop_color,
                                    highlightthickness=self.styles.drop_border_size,
                                    disabledforeground=self.styles.disabled_foreground_drop_color)
        self.drop_stopvalue.config(width=9,
                                   bg=self.styles.drop_background_color,
                                   activebackground=self.styles.active_drop_border_color,
                                   highlightbackground=self.styles.highlight_drop_color,
                                   highlightthickness=self.styles.drop_border_size,
                                   disabledforeground=self.styles.disabled_foreground_drop_color)
        self.drop_stepvalue.config(width=9,
                                   bg=self.styles.drop_background_color,
                                   activebackground=self.styles.active_drop_border_color,
                                   highlightbackground=self.styles.highlight_drop_color,
                                   highlightthickness=self.styles.drop_border_size,
                                   disabledforeground=self.styles.disabled_foreground_drop_color)
        self.default_drop_startvalue.config(width=8,
                                            bg=self.styles.drop_background_color,
                                            activebackground=self.styles.active_drop_border_color,
                                            highlightbackground=self.styles.highlight_drop_color,
                                            highlightthickness=self.styles.drop_border_size,
                                            disabledforeground=self.styles.disabled_foreground_drop_color)
        self.default_drop_stopvalue.config(width=8,
                                           bg=self.styles.drop_background_color,
                                           activebackground=self.styles.active_drop_border_color,
                                           highlightbackground=self.styles.highlight_drop_color,
                                           highlightthickness=self.styles.drop_border_size,
                                           disabledforeground=self.styles.disabled_foreground_drop_color)
        self.default_drop_stepvalue.config(width=8,
                                           bg=self.styles.drop_background_color,
                                           activebackground=self.styles.active_drop_border_color,
                                           highlightbackground=self.styles.highlight_drop_color,
                                           highlightthickness=self.styles.drop_border_size,
                                           disabledforeground=self.styles.disabled_foreground_drop_color)

    def USBOptionMenu(self):
        self.tuple1 = self.usb.list_resources()
        self.list1 = list(self.tuple1)
        self.list1.insert(0, "-")
        self.clicked_usb = StringVar()
        self.clicked_usb.set(self.list1[0])
        self.drop_usb = OptionMenu(self.connection_frame, self.clicked_usb, *self.list1, command=self.Connect_Ctrl)
        self.drop_usb.config(width=35,
                             bg=self.styles.drop_background_color,
                             activebackground=self.styles.active_drop_border_color,
                             highlightbackground=self.styles.highlight_drop_color,
                             highlightthickness=self.styles.drop_border_size,
                             disabledforeground=self.styles.disabled_foreground_drop_color)

    def com_option_menu(self):
        self.get_com_list()
        self.clicked_com = StringVar()
        self.clicked_com.set(self.com_list[0])
        self.drop_com = OptionMenu(self.connection_frame, self.clicked_com, *self.com_list, command=self.Connect_Ctrl)
        self.drop_com.config(width=35,
                             bg=self.styles.drop_background_color,
                             activebackground=self.styles.active_drop_border_color,
                             highlightbackground=self.styles.highlight_drop_color,
                             highlightthickness=self.styles.drop_border_size,
                             disabledforeground=self.styles.disabled_foreground_drop_color)

    def BaudOptionMenu(self):
        bds = ["-", "300", "600", "1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "56000", "57600",
               "115200", "128000", "256000"]
        self.clicked_bd = StringVar()
        self.clicked_bd.set(bds[0])
        self.drop_bd = OptionMenu(self.connection_frame, self.clicked_bd, *bds, command=self.Connect_Ctrl)
        self.drop_bd.config(width=35,
                            bg=self.styles.drop_background_color,
                            activebackground=self.styles.active_drop_border_color,
                            highlightbackground=self.styles.highlight_drop_color,
                            highlightthickness=self.styles.drop_border_size,
                            disabledforeground=self.styles.disabled_foreground_drop_color)

    def Connect_Ctrl(self, Other):
        if "-" in self.clicked_com.get() or "-" in self.clicked_bd.get() or "-" in self.clicked_usb.get():
            self.com_btn_connect["state"] = "disable"
        else:
            self.com_btn_connect["state"] = "normal"

    def Comm_Refresh(self):
        self.drop_com.destroy()
        self.com_option_menu()
        self.drop_com.grid(column=2, row=2, padx=20, pady=5)
        logic = []
        self.Connect_Ctrl(logic)

    def USB_Refresh(self):
        self.drop_usb.destroy()
        self.USBOptionMenu()
        self.drop_usb.grid(column=2, row=1, padx=20, pady=5)
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
                sleep(2)
                self.com_btn_connect["text"] = "Disconnect"
                self.btn_usb_refresh["state"] = "disabled"
                self.btn_refresh["state"] = "disabled"
                self.drop_usb["state"] = "disabled"
                self.drop_bd["state"] = "disabled"
                self.drop_com["state"] = "disabled"
                command = "pulse:" + self.xml[1][8].text + "\n"
                self.sendCommand(command)
                Info_msg = f"Successful UART Connection using {self.clicked_com.get()}\nSuccessful USB Connection using {self.clicked_usb.get()}"
                self.overall_threading = True
                self.value_show_threading = True
                self.t2 = threading.Thread(target=self.display_data, daemon=True)
                self.t2.start()
                self.enable_buttons()
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Connection Successful", Info_msg)
        else:
            self.SerialClose()
            self.overall_threading = False
            self.value_show_threading = False
            self.com_btn_connect["text"] = "Connect"
            self.btn_refresh["state"] = "normal"
            self.btn_usb_refresh["state"] = "normal"
            self.drop_usb["state"] = "normal"
            self.drop_bd["state"] = "normal"
            self.drop_com["state"] = "normal"
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
        current_wavelength = self.wavelength_value
        current_bandwidth = self.bandwidth_value
        while self.overall_threading:
            if self.value_show_threading:
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
            else:
                self.power_value.config(text="-")

            if self.run_threading_status:
                self.value_show_threading = False
                program_complete = False
                if self.parameter_checker() and self.run_threading_status:
                    if self.go_home() and self.run_threading_status:
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
                            if self.take_position(maxima_point_from_home_in_mm - sweeping_distance_in_mm):
                                self.steps_remaining["text"] = "No of Steps remaining: "
                                self.steps_completed["text"] = "No of Steps completed: "
                                self.total_step["text"] = "Total no of Steps: "
                                initial_point = sweeping_distance_in_mm
                                final_point = -sweeping_distance_in_mm
                                self.ax.set_xlim([final_point, initial_point])
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
                                self.no_of_total_steps = round(2*sweeping_distance_in_mm/stepvalue_in_mm)
                                counter = self.no_of_total_steps
                                self.total_step_value["text"] = str(self.no_of_total_steps)
                                while current_point > final_point and self.run_threading_status:
                                    self.steps_remaining_value["text"] = str(round(counter))
                                    self.steps_completed_value["text"] = str(round(self.no_of_total_steps-counter))
                                    self.x_data = np.append(self.x_data, current_point)
                                    value = self.detector.query("read?")
                                    self.y_data = np.append(self.y_data, float(value))
                                    self.power_value.config(text=str(value))
                                    if len(self.y_data) == 1:
                                        pass
                                    else:
                                        y_max = self.y_data.max()
                                        y_min = self.y_data.min()

                                        y_min = y_min - (y_min * 0.01)
                                        y_max = y_max + (y_max * 0.01)
                                        self.ax.set_ylim([y_min, y_max])

                                    self.lines.set_xdata(self.x_data)
                                    self.lines.set_ydata(self.y_data)

                                    self.canvas.draw()

                                    current_point = current_point - stepvalue_in_mm
                                    if self.run_threading_status:
                                        self.take_position(stepvalue_in_mm)
                                        counter = counter-1

                                    if current_point <= final_point:
                                        program_complete = True

                                self.x_data = np.append(self.x_data, current_point)
                                value = self.detector.query("read?")
                                self.y_data = np.append(self.y_data, float(value))
                                self.power_value.config(text=str(value))
                                y_max = self.y_data.max()
                                y_min = self.y_data.min()

                                y_min = y_min - (y_min * 0.01)
                                y_max = y_max + (y_max * 0.01)
                                self.ax.set_ylim([y_min, y_max])

                                self.lines.set_xdata(self.x_data)
                                self.lines.set_ydata(self.y_data)

                                self.canvas.draw()

                if not self.run_threading_status:
                    self.x_data = np.array([])
                    self.y_data = np.array([])
                    self.lines.set_xdata(self.x_data)
                    self.lines.set_ydata(self.y_data)
                    self.canvas.draw()
                    self.status_retainer()
                    self.value_show_threading = True
                    self.steps_remaining["text"] = ""
                    self.steps_completed["text"] = ""
                    self.total_step["text"] = ""
                    self.steps_remaining_value["text"] = ""
                    self.steps_completed_value["text"] = ""
                    self.total_step_value["text"] = ""
                if program_complete:
                    if self.run_button_without_aperture["text"] == "Stop Run":
                        self.without_aperture_x_data = self.x_data
                        self.without_aperture_y_data = self.y_data
                    elif self.run_button_with_aperture["text"] == "Stop Run":
                        self.with_aperture_x_data = self.x_data
                        self.with_aperture_y_data = self.y_data
                    self.x_data = np.array([])
                    self.x_data = np.array([])
                    self.status_retainer()
                    self.value_show_threading = True
                    self.run_threading_status = False
                    self.run_button_with_aperture["state"] = "normal"
                    self.run_button_without_aperture["state"] = "normal"
                    self.save_data_button["state"] = "normal"
                    self.enter_file_name_entry["state"] = "normal"
                    self.save_image_data_button["state"] = "normal"
                    self.reset_all_button["state"] = "normal"
                    self.steps_remaining["text"] = ""
                    self.steps_completed["text"] = ""
                    self.total_step["text"] = ""
                    self.steps_remaining_value["text"] = ""
                    self.steps_completed_value["text"] = ""
                    self.total_step_value["text"] = ""

    def sendCommand(self, command):
        self.command = command
        #print("S= " + command)
        self.ser.reset_output_buffer()
        self.ser.write(self.command.encode("utf-8"))

    def receiveCommand(self):
        self.tester = 0
        while self.tester == 0:
            while self.ser.in_waiting > 0:
                msg = str(self.ser.readline(), "utf-8")
                self.ser.reset_input_buffer()
                self.tester = 1
                #print("R= " + msg)
                return msg

    def homeThreading(self):
        self.sendCommand(self.antiClockwiseCommand)
        if self.receiveCommand() == self.antiClockwise_ok:
            self.notebook.hide(1)
            self.sendCommand(self.homeCommand)
            received_command = self.receiveCommand()
            if received_command == self.home_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Position Status", "Home Reached")
                self.status_retainer()
            elif received_command == self.stop_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                self.status_retainer()

    def endThreading(self):
        self.sendCommand(self.clockwiseCommand)
        if self.receiveCommand() == self.clockwise_ok:
            self.notebook.hide(1)
            self.sendCommand(self.endCommand)
            received_command = self.receiveCommand()
            if received_command == self.end_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Position Status", "Reached the End")
                self.status_retainer()
            elif received_command == self.stop_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                self.status_retainer()

    def moveThreading(self):
        received_command = ""
        pulse_number = 1
        validator = 0
        if self.clicked_desired_unit.get() == "mm":
            if ((float(self.desired_distance.get()) * 1000) % self.minimumvalue_text) == 0:
                pulse_number = float(self.desired_distance.get()) * self.motor_steps
                validator = 1
            else:
                messagebox.showerror("Error",
                                     f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                self.move_retainer()
        elif self.clicked_desired_unit.get() == "microns":
            if (float(self.desired_distance.get()) % self.minimumvalue_text) == 0:
                pulse_number = (float(self.desired_distance.get()) / 1000) * self.motor_steps
                validator = 1
            else:
                messagebox.showerror("Error",
                                     f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                self.move_retainer()
        if pulse_number <= 0:
            validator = 0
            messagebox.showerror("Value error", "The value must be more than 0")
            self.move_retainer()
        if validator == 1:
            if self.clicked_desired_direction.get() == "end":
                self.sendCommand(self.clockwiseCommand)
                if self.receiveCommand() == self.clockwise_ok:
                    command = "move:" + str(round(pulse_number)) + "\n"
                    self.sendCommand(command)
                    received_command = self.receiveCommand()
            else:
                self.sendCommand(self.antiClockwiseCommand)
                if self.receiveCommand() == self.antiClockwise_ok:
                    command = "move:" + str(round(pulse_number)) + "\n"
                    self.sendCommand(command)
                    received_command = self.receiveCommand()
            if received_command == self.home_ok:
                messagebox.showinfo("Position Status", "Home Reached")
            elif received_command == self.end_ok:
                messagebox.showinfo("Position Status", "End Reached")
            elif received_command == self.stop_ok or received_command == self.move_ok:
                pass
            self.move_retainer()

    def move_retainer(self):
        self.move_button["text"] = "Move"
        self.enable_manual_mode["state"] = "normal"
        self.enable_motor_button["state"] = "normal"

    def locateThreading(self):
        self.sendCommand(self.antiClockwiseCommand)
        if self.receiveCommand() == self.antiClockwise_ok:
            self.sendCommand(self.locateCommand)
            self.holder_position["text"] = 'Please wait...'
            command = self.receiveCommand()
            if command == self.stop_ok:
                self.text_processor(0)
            else:
                command = ''.join(letter for letter in command if letter.isalnum())
                self.text_processor(int(command))
            self.enable_manual_mode["state"] = "normal"
            self.notebook.hide(0)
            self.find_position["text"] = "Locate the position"
            self.save_button["state"] = "normal"
            self.apply_button["state"] = "normal"
            self.notebook.add(self.mainframe, text="Controls")

    def parameter_checker(self):
        if len(self.startvalueentry.get()) != 0:
            if self.num_validator(self.startvalueentry.get()):
                if round(float(self.startvalueentry.get()), 4) > 0:
                    if len(self.endvalueentry.get()) != 0:
                        if self.num_validator(self.endvalueentry.get()):
                            if round(float(self.endvalueentry.get()), 4) > 0:
                                if len(self.stepvalueentry.get()) != 0:
                                    if self.num_validator(self.endvalueentry.get()):
                                        if round(float(self.stepvalueentry.get()), 4) > 0:
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
                                                    self.run_threading_status = False
                                                    messagebox.showerror("Error",
                                                                         f"The minimum value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                                            else:
                                                self.run_threading_status = False
                                                messagebox.showerror("Value Error",
                                                                     'The value in "Distance from Home to Maxima" field should be less than Sweeping distance value')
                                        else:
                                            self.run_threading_status = False
                                            messagebox.showerror("Value Error",
                                                                 "Step Value field should be greater than 0")
                                    else:
                                        self.run_threading_status = False
                                        messagebox.showerror("Value Error", "Step Value field should be a number")
                                else:
                                    self.run_threading_status = False
                                    messagebox.showerror("Value Error", "Step Value field should not be blank")
                            else:
                                self.run_threading_status = False
                                messagebox.showerror("Value Error",
                                                     "Half Sweeping Distance field should be greater than 0")
                        else:
                            self.run_threading_status = False
                            messagebox.showerror("Value Error", "Sweeping Distance field should be a number")
                    else:
                        self.run_threading_status = False
                        messagebox.showerror("Value Error", "Sweeping Distance field should not be blank")
                else:
                    self.run_threading_status = False
                    messagebox.showerror("Value Error", '"Distance from Home to Maxima" field should be greater than 0')
            else:
                self.run_threading_status = False
                messagebox.showerror("Value Error", '"Distance from Home to Maxima" field should be a number')
        else:
            self.run_threading_status = False
            messagebox.showerror("Value Error", '"Distance from Home to Maxima" field should not be blank')

    def go_home(self):
        self.sendCommand(self.antiClockwiseCommand)
        if self.receiveCommand() == self.antiClockwise_ok:
            self.sendCommand(self.homeCommand)
            command = self.receiveCommand()
            if command == self.home_ok:
                return True
            elif command == self.stop_ok:
                return False

    def take_position(self, pulse_number):
        pulse_number = pulse_number * self.motor_steps
        command = "move:" + str(round(pulse_number)) + "\n"
        self.sendCommand(command)
        received_command = self.receiveCommand()
        if received_command == self.move_ok:
            return True
        elif received_command == self.stop_ok:
            return False
        elif received_command == self.end_ok:
            messagebox.showinfo("Position Status", "Reached the End")
            return False

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
        self.save_data_button["state"] = "disabled"
        self.enter_file_name_entry["state"] = "disabled"
        self.reset_all_button["state"] = "disabled"
        self.save_image_data_button["state"] = "disabled"
        self.com_btn_connect["state"] = "normal"

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
        Checkbutton(self.plot_window, text="Plot without aperture", variable=self.without_aperture_var).grid(row=0,
                                                                                                             column=0,
                                                                                                             pady=5,
                                                                                                             sticky=W)
        self.with_aperture_var = IntVar()
        Checkbutton(self.plot_window, text="Plot with aperture", variable=self.with_aperture_var).grid(row=1, column=0,
                                                                                                       pady=5, sticky=W)
        self.plot_the_graph_button = Button(self.plot_window,
                                            text="Plot the Graph",
                                            command=self.plot_the_graph,
                                            bg=self.styles.button_background,
                                            fg=self.styles.button_foreground,
                                            bd=self.styles.button_border_size,
                                            highlightbackground=self.styles.button_highlight_background,
                                            highlightcolor=self.styles.button_highlight_foreground,
                                            disabledforeground=self.styles.button_disabled_foreground,
                                            activebackground=self.styles.button_active_background,
                                            activeforeground=self.styles.button_active_foreground,
                                            font=self.styles.bold_font)
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
        else:
            return True

    def excel_selector(self):
        self.excel_window = Toplevel()
        self.excel_window.title("Excel Data Selector")
        self.excel_window.geometry("250x120")
        self.excel_without_aperture_var = IntVar()
        Checkbutton(self.excel_window, text="Plot without aperture", variable=self.excel_without_aperture_var).grid(
            row=0,
            column=0,
            pady=5,
            sticky=W)
        self.excel_with_aperture_var = IntVar()
        Checkbutton(self.excel_window, text="Plot with aperture", variable=self.excel_with_aperture_var).grid(row=1,
                                                                                                              column=0,
                                                                                                              pady=5,
                                                                                                              sticky=W)
        self.export_excel = Button(self.excel_window,
                                   text="Prepare Data",
                                   command=self.excel_mode_selector,
                                   bg=self.styles.button_background,
                                   fg=self.styles.button_foreground,
                                   bd=self.styles.button_border_size,
                                   highlightbackground=self.styles.button_highlight_background,
                                   highlightcolor=self.styles.button_highlight_foreground,
                                   disabledforeground=self.styles.button_disabled_foreground,
                                   activebackground=self.styles.button_active_background,
                                   activeforeground=self.styles.button_active_foreground,
                                   font=self.styles.bold_font)
        self.export_excel.grid(row=2, column=0, pady=5)

    def excel_mode_selector(self):
        if self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 0:
            messagebox.showerror("Selection Error", "Select at least one box")
        elif self.excel_without_aperture_var.get() == 1 and self.excel_with_aperture_var.get() == 0:
            self.without_aperture_excel()
        elif self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 1:
            self.with_aperture_excel()
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
        if self.home_button["text"] == "Home":
            self.home_button["state"] = "disabled"
        if self.end_button["text"] == "End":
            self.end_button["state"] = "disabled"
        if self.run_button_without_aperture["text"] == "Start Run without Aperture":
            self.run_button_without_aperture["state"] = "normal"
        if self.run_button_with_aperture["text"] == "Start Run with Aperture":
            self.run_button_with_aperture["state"] = "disabled"
        self.startvalueentry["state"] = "disabled"
        self.endvalueentry["state"] = "disabled"
        self.stepvalueentry["state"] = "disabled"
        self.com_btn_connect["state"] = "disabled"
        self.drop_startvalue["state"] = "disabled"
        self.drop_stopvalue["state"] = "disabled"
        self.drop_stepvalue["state"] = "disabled"

    def status_retainer(self):
        if self.home_button["text"] == "Stop":
            self.home_button["text"] = "Home"
        else:
            self.home_button["state"] = "normal"
        if self.end_button["text"] == "Stop":
            self.end_button["text"] = "End"
        else:
            self.end_button["state"] = "normal"
        if self.run_button_without_aperture["text"] == "Stop Run":
            self.run_button_without_aperture["text"] = "Start Run without Aperture"
        else:
            self.run_button_without_aperture["state"] = "normal"
        if self.run_button_with_aperture["text"] == "Stop Run":
            self.run_button_with_aperture["text"] = "Start Run with Aperture"
        else:
            self.run_button_with_aperture["state"] = "normal"
        self.startvalueentry["state"] = "normal"
        self.endvalueentry["state"] = "normal"
        self.stepvalueentry["state"] = "normal"
        self.com_btn_connect["state"] = "normal"
        self.drop_startvalue["state"] = "normal"
        self.drop_stopvalue["state"] = "normal"
        self.drop_stepvalue["state"] = "normal"
        self.notebook.add(self.settingsframe, text="Settings")

    def Home(self):
        if self.home_button["text"] == "Home":
            self.t1 = threading.Thread(target=self.homeThreading, daemon=True)
            self.t1.start()
            self.home_button["text"] = "Stop"
            self.status_updater()
        else:
            self.sendCommand(self.stopCommand)

    def End(self):
        if self.end_button["text"] == "End":
            self.t1 = threading.Thread(target=self.endThreading, daemon=True)
            self.t1.start()
            self.end_button["text"] = "Stop"
            self.status_updater()
        else:
            self.sendCommand(self.stopCommand)

    def Run_with_aperture(self):
        if self.run_button_with_aperture["text"] == "Start Run with Aperture":
            self.run_threading_status = True
            self.run_button_with_aperture["text"] = "Stop Run"
            self.notebook.hide(1)
            self.status_updater()
            self.save_data_button["state"] = "disabled"
            self.enter_file_name_entry["state"] = "disabled"
            self.reset_all_button["state"] = "disabled"
            self.save_image_data_button["state"] = "disabled"
        else:
            self.sendCommand(self.stopCommand)
            self.run_threading_status = False

    def Run_without_aperture(self):
        if self.run_button_without_aperture["text"] == "Start Run without Aperture":
            self.value_show_threading = False
            sleep(0.1)
            self.run_threading_status = True
            self.run_button_without_aperture["text"] = "Stop Run"
            self.notebook.hide(1)
            self.status_updater()
            self.save_data_button["state"] = "disabled"
            self.enter_file_name_entry["state"] = "disabled"
            self.reset_all_button["state"] = "disabled"
            self.save_image_data_button["state"] = "disabled"
        else:
            self.sendCommand(self.stopCommand)
            self.run_threading_status = False

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
                if 400 <= float(wavelength) <= 1100:
                    if len(thread_spacing) > 0:
                        if self.num_validator(thread_spacing):
                            if len(self.pulse_width_entry.get()) > 0:
                                if self.num_validator(self.pulse_width_entry.get()):
                                    if 180 <= int(self.pulse_width_entry.get()) <= 32769:
                                        if len(self.steps_per_rotation_entry.get()) > 0:
                                            if self.num_validator(self.steps_per_rotation_entry.get()):
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
                                                    if len(self.close_aperture_entry.get()) == 0:
                                                        self.xml[1][7].text = " "
                                                    else:
                                                        self.xml[1][7].text = self.close_aperture_entry.get()
                                                    self.xml[1][8].text = self.pulse_width_entry.get()
                                                    self.xml[1][9].text = self.steps_per_rotation_entry.get()

                                                    tree.write("data_file.xml")
                                                    messagebox.showinfo("Save status", "Data saved successfully")
                                            else:
                                                messagebox.showerror("Value Error",
                                                                     "Steps per rotation value must be a number")
                                        else:
                                            messagebox.showerror("Value Error",
                                                                 "Steps per rotation value should not be blank")
                                    else:
                                        messagebox.showerror("Value Error",
                                                             "Pulse Width value must in between 180 µs (Fast) to 32769 µs (Slow)")
                                else:
                                    messagebox.showerror("Value Error", "Pulse Width value must be a number")
                            else:
                                messagebox.showerror("Value Error", "Pulse Width value should not be blank")
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
                            if len(self.pulse_width_entry.get()) > 0:
                                if self.num_validator(self.pulse_width_entry.get()):
                                    if 180 <= int(self.pulse_width_entry.get()) <= 32769:
                                        if len(self.steps_per_rotation_entry.get()) > 0:
                                            if self.num_validator(self.steps_per_rotation_entry.get()):
                                                self.wavelength_value = self.wavelength
                                                if self.clicked_bandwidth.get() == "High":
                                                    self.bandwidth_value = "High"
                                                elif self.clicked_bandwidth.get() == "Low":
                                                    self.bandwidth_value = "Low"
                                                self.motor_steps = int(self.steps_per_rotation_entry.get())
                                                command = "pulse:" + self.pulse_width_entry.get() + "\n"
                                                self.sendCommand(command)
                                                self.minimumvalue_text = float(
                                                    self.thread_spacing_entry.get()) / self.motor_steps * 1000
                                                self.minimumvalue.config(
                                                    text=f"Note: The minimum step value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm")
                                                messagebox.showinfo("Apply status", "Data applied successfully")
                                            else:
                                                messagebox.showerror("Value Error",
                                                                     "Steps per rotation value must be a number")
                                        else:
                                            messagebox.showerror("Value Error",
                                                                 "Steps per rotation value should not be blank")
                                    else:
                                        messagebox.showerror("Value Error",
                                                             "Pulse Width value must in between 180 µs (Fast) to 32769 µs (Slow)")
                                else:
                                    messagebox.showerror("Value Error", "Pulse Width value must be a number")
                            else:
                                messagebox.showerror("Value Error", "Pulse Width value should not be blank")
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
        self.home_button["state"] = "normal"
        self.end_button["state"] = "normal"
        self.run_button_without_aperture["state"] = "normal"
        self.run_button_with_aperture["state"] = "normal"
        self.startvalueentry["state"] = "normal"
        self.endvalueentry["state"] = "normal"
        self.stepvalueentry["state"] = "normal"
        self.drop_startvalue["state"] = "normal"
        self.drop_stopvalue["state"] = "normal"
        self.drop_stepvalue["state"] = "normal"

    def disable_buttons(self):
        self.home_button["state"] = "disabled"
        self.end_button["state"] = "disabled"
        self.run_button_without_aperture["state"] = "disabled"
        self.run_button_with_aperture["state"] = "disabled"
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
        self.drop_usb.grid(column=2, row=1, padx=20, pady=5)
        self.label_com.grid(column=1, row=2)
        self.drop_com.grid(column=2, row=2, padx=20, pady=5)
        self.label_bd.grid(column=1, row=3)
        self.drop_bd.grid(column=2, row=3, padx=20, pady=5)

        self.btn_usb_refresh.grid(column=3, row=1)
        self.btn_refresh.grid(column=3, row=2)
        self.com_btn_connect.grid(column=3, row=3)

        self.communication_frame.grid_propagate(False)
        self.communication_frame.grid(row=0, column=4, rowspan=3, columnspan=5, padx=5, pady=5)

        self.home_button.grid(row=0, column=4, pady=5, padx=(10, 0))
        self.end_button.grid(row=1, column=4, pady=5, padx=(10, 0))
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
        self.run_button_without_aperture.grid(row=2, column=1, pady=(20, 8))
        self.run_button_with_aperture.grid(row=5, column=1, pady=8)
        self.enter_file_name_label.grid(row=6, column=1, pady=2)
        self.enter_file_name_entry.grid(row=7, column=1, pady=2)
        self.save_data_button.grid(row=8, column=1, pady=8)
        self.save_image_data_button.grid(row=9, column=1, pady=8)
        self.reset_all_button.grid(row=10, column=1, pady=8)

        self.seperator3.place(relx=0.83, rely=-0.04, relwidth=0.001, relheight=1.055)
        self.seperator4.place(relx=0.83, rely=0.18, relwidth=0.174, relheight=0.001)
        self.seperator5.place(relx=0.83, rely=0.425, relwidth=0.174, relheight=0.001)
        self.seperator6.place(relx=0.83, rely=0.705, relwidth=0.174, relheight=0.001)
        # self.seperator7.place(relx=0.83, rely=0.793, relwidth=0.174, relheight=0.001)
        # self.seperator7.place(relx=0.83, rely=0.896, relwidth=0.174, relheight=0.001)

        self.steps_remaining.place(relx=0.1, rely=-0.02)
        self.steps_completed.place(relx=0.35, rely=-0.02)
        self.total_step.place(relx=0.6, rely=-0.02)

        self.steps_remaining_value.place(relx=0.21, rely=-0.02)
        self.steps_completed_value.place(relx=0.46, rely=-0.02)
        self.total_step_value.place(relx=0.68, rely=-0.02)

        self.Config.grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        self.Config.grid_propagate(False)

        self.wavelength.grid(row=0, column=0)
        self.bandwidth.grid(row=1, column=0)
        self.thread_spacing.grid(row=2, column=0)
        self.pulse_width_label.grid(row=3, column=0)
        self.steps_per_rotation_label.grid(row=4, column=0)
        self.open_aperture_dia.grid(row=0, column=4)
        self.enter_close_aperture_label.grid(row=1, column=4)
        self.default_startvalue.grid(row=2, column=4)
        self.default_stopvalue.grid(row=3, column=4)
        self.default_stepvalue.grid(row=4, column=4)

        self.wavelength_entry.grid(row=0, column=1, padx=(0, 7))
        self.drop_bandwidth.grid(row=1, column=1, padx=(0, 7))
        self.thread_spacing_entry.grid(row=2, column=1, padx=(0, 7))
        self.pulse_width_entry.grid(row=3, column=1, padx=(0, 7))
        self.steps_per_rotation_entry.grid(row=4, column=1, padx=(0, 7))
        self.open_aperture_dia_entry.grid(row=0, column=5, padx=(0, 7))
        self.close_aperture_entry.grid(row=1, column=5, pady=2)
        self.default_startvalue_entry.grid(row=2, column=5, padx=(0, 7))
        self.default_stopvalue_entry.grid(row=3, column=5, padx=(0, 7))
        self.default_stepvalue_entry.grid(row=4, column=5, padx=(0, 7))

        self.wavelength_unit.grid(row=0, column=2)
        self.thread_spacing_unit.grid(row=2, column=2)
        self.pulse_width_unit.grid(row=3, column=2)
        self.open_aperture_unit.grid(row=0, column=6)
        self.close_aperture_unit.grid(row=1, column=6)
        self.default_drop_startvalue.grid(row=2, column=6)
        self.default_drop_stopvalue.grid(row=3, column=6)
        self.default_drop_stepvalue.grid(row=4, column=6)

        self.seperator2.place(relx=0.42, rely=0, relwidth=0.003, relheight=0.8)

        self.save_button.grid(row=0, column=1, padx=5)
        self.apply_button.grid(row=0, column=2, padx=5)

        self.button_frame.grid(row=5, column=0, columnspan=7, pady=20)

        self.motor_control.grid(row=1, column=0, padx=5, pady=5)
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
    CommGUI()
