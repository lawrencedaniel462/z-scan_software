from tkinter import *
from tkinter import ttk, messagebox
import threading
from time import sleep
import serial.tools.list_ports
import xml.etree.ElementTree as et
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np

tree = et.parse("data_file.xml")


# bg = button_background, fg = button_foreground, bd = button_border_size, highlightbackground = button_highlight_background, highlightcolor = button_highlight_foreground, disabledforeground = button_disabled_foreground
class CommGUI:
    def __init__(self, root, usb, style, command, excel, graph, view, calculation):
        self.xml = tree.getroot()

        self.motor_steps = float(self.xml[1][9].text)
        self.root = root
        self.usb = usb
        self.styles = style
        self.command = command
        self.excel = excel
        self.graph = graph
        self.view = view
        self.calculation = calculation

        self.style = ttk.Style()

        self.style.theme_create("custom_theme", parent="alt", settings={
            "TLabelframe": {
                "configure": {  # "padx": 5,
                    # "pady": 5,
                    "background": "#FFE5B4",
                    "bordercolor": "black",  # "#CD7F32",
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
                "map": {"background": [("selected", self.styles.bg_color), ('!active', self.styles.light_bg_color),
                                       ('active', "#FAC898")]}
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
        self.graph_view_frame = Frame(self.notebook,
                                      bg=self.styles.bg_color)
        self.calculation_frame = Frame(self.notebook,
                                      bg=self.styles.bg_color)

        self.settingsframe_content()
        self.mainframe_content()
        self.initialize_graph_viewer()
        self.calculation.initialize_calculation(self)

        self.current_position = 0

        self.value_option_menu()

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

        self.x_data = np.array([])
        self.y_data = np.array([])
        self.without_aperture_x_data = np.array([])
        self.without_aperture_y_data = np.array([])
        self.with_aperture_x_data = np.array([])
        self.with_aperture_y_data = np.array([])

        self.title_value = ""
        self.sweeping_distance_value = ""
        self.step_distance_value = ""
        self.bandwidth_value = ""

    def settingsframe_content(self):
        self.Config = LabelFrame(self.settingsframe,
                                 text="Configuration",
                                 padx=5,
                                 pady=5,
                                 width=880,
                                 height=350,
                                 bg=self.styles.bg_color)
        self.wavelength_entry = Entry(self.Config,
                                      width=10,
                                      font=('Helvetica', 10),
                                      bg=self.styles.entry_background_color,
                                      highlightbackground=self.styles.entry_border_color,
                                      highlightthickness=self.styles.entry_border_size,
                                      highlightcolor=self.styles.active_entry_border_color,
                                      disabledbackground=self.styles.disabled_background_entry,
                                      disabledforeground=self.styles.disabled_foreground_entry)
        self.wavelength_entry.insert(0, self.xml[0][0].text)
        self.laser_beam_input_power_entry = Entry(self.Config,
                                                  width=10,
                                                  font=('Helvetica', 10),
                                                  bg=self.styles.entry_background_color,
                                                  highlightbackground=self.styles.entry_border_color,
                                                  highlightthickness=self.styles.entry_border_size,
                                                  highlightcolor=self.styles.active_entry_border_color,
                                                  disabledbackground=self.styles.disabled_background_entry,
                                                  disabledforeground=self.styles.disabled_foreground_entry)
        self.laser_beam_input_power_entry.insert(0, self.xml[0][3].text)
        self.laser_beam_diameter_entry = Entry(self.Config,
                                               width=10,
                                               font=('Helvetica', 10),
                                               bg=self.styles.entry_background_color,
                                               highlightthickness=self.styles.entry_border_size,
                                               highlightcolor=self.styles.active_entry_border_color,
                                               disabledbackground=self.styles.disabled_background_entry,
                                               disabledforeground=self.styles.disabled_foreground_entry)
        self.laser_beam_diameter_entry.insert(0, self.xml[0][4].text)
        self.transmittance_entry = Entry(self.Config,
                                         width=10,
                                         font=('Helvetica', 10),
                                         bg=self.styles.entry_background_color,
                                         highlightthickness=self.styles.entry_border_size,
                                         highlightcolor=self.styles.active_entry_border_color,
                                         disabledbackground=self.styles.disabled_background_entry,
                                         disabledforeground=self.styles.disabled_foreground_entry)
        self.transmittance_entry.insert(0, self.xml[0][5].text)
        self.transmittance_thickness_entry = Entry(self.Config,
                                                   width=10,
                                                   font=('Helvetica', 10),
                                                   bg=self.styles.entry_background_color,
                                                   highlightthickness=self.styles.entry_border_size,
                                                   highlightcolor=self.styles.active_entry_border_color,
                                                   disabledbackground=self.styles.disabled_background_entry,
                                                   disabledforeground=self.styles.disabled_foreground_entry)
        self.transmittance_thickness_entry.insert(0, self.xml[0][6].text)
        self.thread_spacing_entry = Entry(self.Config,
                                          width=10,
                                          font=('Helvetica', 10),
                                          bg=self.styles.entry_background_color,
                                          highlightbackground=self.styles.entry_border_color,
                                          highlightthickness=self.styles.entry_border_size,
                                          highlightcolor=self.styles.active_entry_border_color,
                                          disabledbackground=self.styles.disabled_background_entry,
                                          disabledforeground=self.styles.disabled_foreground_entry)
        self.thread_spacing_entry.insert(0, self.xml[0][2].text)
        self.pulse_width_entry = Entry(self.Config,
                                       width=10,
                                       font=('Helvetica', 10),
                                       bg=self.styles.entry_background_color,
                                       highlightbackground=self.styles.entry_border_color,
                                       highlightthickness=self.styles.entry_border_size,
                                       highlightcolor=self.styles.active_entry_border_color,
                                       disabledbackground=self.styles.disabled_background_entry,
                                       disabledforeground=self.styles.disabled_foreground_entry)
        self.pulse_width_entry.insert(0, self.xml[1][8].text)
        self.steps_per_rotation_entry = Entry(self.Config,
                                              width=10,
                                              font=('Helvetica', 10),
                                              bg=self.styles.entry_background_color,
                                              highlightbackground=self.styles.entry_border_color,
                                              highlightthickness=self.styles.entry_border_size,
                                              highlightcolor=self.styles.active_entry_border_color,
                                              disabledbackground=self.styles.disabled_background_entry,
                                              disabledforeground=self.styles.disabled_foreground_entry)
        self.steps_per_rotation_entry.insert(0, self.xml[1][9].text)
        self.focal_length_entry = Entry(self.Config,
                                             width=10,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.focal_length_entry.insert(0, self.xml[1][6].text)
        self.close_aperture_entry = Entry(self.Config,
                                          width=10,
                                          font=('Helvetica', 10),
                                          bg=self.styles.entry_background_color,
                                          highlightbackground=self.styles.entry_border_color,
                                          highlightthickness=self.styles.entry_border_size,
                                          highlightcolor=self.styles.active_entry_border_color,
                                          disabledbackground=self.styles.disabled_background_entry,
                                          disabledforeground=self.styles.disabled_foreground_entry)
        self.close_aperture_entry.insert(0, self.xml[1][7].text)
        self.beam_rad_at_aperture_entry = Entry(self.Config,
                                                width=10,
                                                font=('Helvetica', 10),
                                                bg=self.styles.entry_background_color,
                                                highlightbackground=self.styles.entry_border_color,
                                                highlightthickness=self.styles.entry_border_size,
                                                highlightcolor=self.styles.active_entry_border_color,
                                                disabledbackground=self.styles.disabled_background_entry,
                                                disabledforeground=self.styles.disabled_foreground_entry)
        self.beam_rad_at_aperture_entry.insert(0, self.xml[0][7].text)
        self.z_scan_sample_thickness_entry = Entry(self.Config,
                                                   width=10,
                                                   font=('Helvetica', 10),
                                                   bg=self.styles.entry_background_color,
                                                   highlightbackground=self.styles.entry_border_color,
                                                   highlightthickness=self.styles.entry_border_size,
                                                   highlightcolor=self.styles.active_entry_border_color,
                                                   disabledbackground=self.styles.disabled_background_entry,
                                                   disabledforeground=self.styles.disabled_foreground_entry)
        self.z_scan_sample_thickness_entry.insert(0, self.xml[0][8].text)
        self.linear_refractive_index_entry = Entry(self.Config,
                                                   width=10,
                                                   font=('Helvetica', 10),
                                                   bg=self.styles.entry_background_color,
                                                   highlightbackground=self.styles.entry_border_color,
                                                   highlightthickness=self.styles.entry_border_size,
                                                   highlightcolor=self.styles.active_entry_border_color,
                                                   disabledbackground=self.styles.disabled_background_entry,
                                                   disabledforeground=self.styles.disabled_foreground_entry)
        self.linear_refractive_index_entry.insert(0, self.xml[0][9].text)
        self.default_home_to_maxima_entry = Entry(self.Config,
                                              width=10,
                                              font=('Helvetica', 10),
                                              bg=self.styles.entry_background_color,
                                              highlightbackground=self.styles.entry_border_color,
                                              highlightthickness=self.styles.entry_border_size,
                                              highlightcolor=self.styles.active_entry_border_color,
                                              disabledbackground=self.styles.disabled_background_entry,
                                              disabledforeground=self.styles.disabled_foreground_entry)
        self.default_home_to_maxima_entry.insert(0, self.xml[1][0].text)
        self.default_stopvalue_entry = Entry(self.Config,
                                             width=10,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.default_stopvalue_entry.insert(0, self.xml[1][2].text)
        self.default_step_value_entry = Entry(self.Config,
                                             width=10,
                                             font=('Helvetica', 10),
                                             bg=self.styles.entry_background_color,
                                             highlightbackground=self.styles.entry_border_color,
                                             highlightthickness=self.styles.entry_border_size,
                                             highlightcolor=self.styles.active_entry_border_color,
                                             disabledbackground=self.styles.disabled_background_entry,
                                             disabledforeground=self.styles.disabled_foreground_entry)
        self.default_step_value_entry.insert(0, self.xml[1][4].text)

        self.wavelength = Label(self.Config, text="Wavelength of the Laser (λ)", anchor="w", width=35, padx=7, pady=8,
                                bg=self.styles.bg_color)
        self.laser_beam_input_power = Label(self.Config, text="Power of the Laser (E\u209A)", anchor="w", width=35,
                                            padx=7, pady=8,bg=self.styles.bg_color)
        self.laser_beam_diameter = Label(self.Config, text="Diameter of the Laser beam (d)", anchor="w", width=35,
                                         padx=7, pady=8, bg=self.styles.bg_color)
        self.transmittance = Label(self.Config, text="Transmittance (T)", anchor="w", width=35,
                                   padx=7, pady=8, bg=self.styles.bg_color)
        self.transmittance_thickness = Label(self.Config, text="Thickness at transmittance measurement (t)", anchor="w",
                                             width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.bandwidth = Label(self.Config, text="Bandwidth", anchor="w", width=35, padx=7, pady=8,
                               bg=self.styles.bg_color)
        self.thread_spacing = Label(self.Config, text="Thread Spacing", anchor="w", width=35, padx=7, pady=8,
                                    bg=self.styles.bg_color)
        self.pulse_width_label = Label(self.Config, text="Pulse width", anchor="w", width=35, padx=7, pady=8,
                                       bg=self.styles.bg_color)
        self.steps_per_rotation_label = Label(self.Config, text="No of steps per rotation", anchor="w", width=35,
                                              padx=7, pady=8, bg=self.styles.bg_color)
        self.focal_length = Label(self.Config, text="Focal length of the lens (f)", anchor="w", width=35, padx=7,
                                  pady=8, bg=self.styles.bg_color)
        self.close_aperture_label = Label(self.Config, text="Radius of the Aperture (r\u2090)",
                                          anchor="w", width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.beam_rad_at_aperture_label = Label(self.Config, text="Radius of the beam at aperture (w\u2090)",
                                                anchor="w", width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.z_scan_sample_thickness_label = Label(self.Config, text="Z-scan sample thickness (L)",
                                                   anchor="w", width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.linear_refractive_index_label = Label(self.Config, text="Linear refractive index (n\u2080)",
                                                   anchor="w", width=35, padx=7, pady=8, bg=self.styles.bg_color)
        self.default_home_to_maxima = Label(self.Config, text="Default distance from Home to Maxima", anchor="w", width=35,
                                        padx=7, pady=8, bg=self.styles.bg_color)
        self.default_stopvalue = Label(self.Config, text="Default Half Sweeping Distance", anchor="w", width=35, padx=7,
                                       pady=8, bg=self.styles.bg_color)
        self.default_step_value = Label(self.Config, text="Default Step Value", anchor="w", width=35, padx=7, pady=8,
                                       bg=self.styles.bg_color)

        self.wavelength_unit = Label(self.Config, text="nm", anchor="w", width=7, bg=self.styles.bg_color)
        self.laser_beam_input_power_unit = Label(self.Config, text="W", anchor="w", width=7, bg=self.styles.bg_color)
        self.laser_beam_diameter_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.transmittance_unit = Label(self.Config, text="%", anchor="w", width=7, bg=self.styles.bg_color)
        self.transmittance_thickness_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.thread_spacing_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.pulse_width_unit = Label(self.Config, text="µs", anchor="w", width=7, bg=self.styles.bg_color)
        self.focal_length_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.close_aperture_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.beam_rad_at_aperture_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)
        self.z_scan_sample_thickness_unit = Label(self.Config, text="mm", anchor="w", width=7, bg=self.styles.bg_color)

        self.seperator2 = ttk.Separator(self.Config, orient="horizontal")

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

    def mainframe_content(self):
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
        self.intensity_button = Button(self.communication_frame,
                                  text="Start Run", width=23,
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
        self.home_to_maxima = Label(self.communication_frame, text="Distance from Home to Maxima", anchor="w", width=27,
                                bg=self.styles.bg_color)
        self.sweeping_distance = Label(self.communication_frame, text="Half Sweeping Distance", anchor="w", width=27,
                              bg=self.styles.bg_color)
        self.step_value = Label(self.communication_frame, text="Step Value", anchor="w", width=27,
                               bg=self.styles.bg_color)

        self.home_to_maxima_entry = Entry(self.communication_frame,
                                     width=15,
                                     bg=self.styles.entry_background_color,
                                     highlightbackground=self.styles.entry_border_color,
                                     highlightthickness=self.styles.entry_border_size,
                                     highlightcolor=self.styles.active_entry_border_color,
                                     disabledbackground=self.styles.disabled_background_entry,
                                     disabledforeground=self.styles.disabled_foreground_entry)
        self.home_to_maxima_entry.insert(0, self.default_home_to_maxima_entry.get())
        self.sweeping_distance_entry = Entry(self.communication_frame,
                                   width=15,
                                   bg=self.styles.entry_background_color,
                                   highlightbackground=self.styles.entry_border_color,
                                   highlightthickness=self.styles.entry_border_size,
                                   highlightcolor=self.styles.active_entry_border_color,
                                   disabledbackground=self.styles.disabled_background_entry,
                                   disabledforeground=self.styles.disabled_foreground_entry)
        self.sweeping_distance_entry.insert(0, self.default_stopvalue_entry.get())
        self.step_value_entry = Entry(self.communication_frame,
                                    width=15,
                                    bg=self.styles.entry_background_color,
                                    highlightbackground=self.styles.entry_border_color,
                                    highlightthickness=self.styles.entry_border_size,
                                    highlightcolor=self.styles.active_entry_border_color,
                                    disabledbackground=self.styles.disabled_background_entry,
                                    disabledforeground=self.styles.disabled_foreground_entry)
        self.step_value_entry.insert(0, self.default_step_value_entry.get())

        self.minimumvalue_text = float(self.thread_spacing_entry.get()) / self.motor_steps * 1000
        self.minimumvalue = Label(self.communication_frame,
                                  text=f"Note: The minimum step value should be a multiple(s) of {str(self.minimumvalue_text)} microns or {str(self.minimumvalue_text / 1000)} mm",
                                  bg=self.styles.bg_color)
        self.seperator = ttk.Separator(self.communication_frame, orient="vertical")

        self.graph_frame = LabelFrame(self.mainframe, text="Graph Manager", pady=5, padx=5, width=1180, height=385,
                                      bg=self.styles.bg_color)

        self.graph_buttons_frame = Frame(self.graph_frame, width=190, height=350, bg=self.styles.bg_color)
        self.power_label = Label(self.graph_buttons_frame, text="Power in W", anchor="center", width=26,
                                 bg=self.styles.bg_color)
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
                                       command=lambda: self.excel.save_data_excel(self),
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
                                             command=lambda: self.graph.save_as_image(self.with_aperture_x_data,
                                                                                      self.without_aperture_x_data,
                                                                                      self.with_aperture_y_data,
                                                                                      self.without_aperture_y_data),
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
        self.total_step_value = Label(self.graph_frame, text="", anchor="w", width=13, foreground="Red",
                                      bg=self.styles.bg_color)

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
            if self.enable_manual_mode["text"] == "Enable manual mode":
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
            self.sendCommand(self.command.stopCommand)

    def manual_mode_enable(self):
        if self.enable_manual_mode["text"] == "Enable manual mode":
            if self.find_position["text"] == "Locate the position":
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
            self.sendCommand(self.command.antiClockwiseCommand)
            if self.receiveCommand() == self.command.antiClockwise_ok:
                self.sendCommand(self.command.pulseCommand)
                if self.receiveCommand() == self.command.pulse_ok:
                    pass
                else:
                    messagebox.showinfo("Position Status", "Home Reached")
        else:
            messagebox.showerror("Motor Status Error", "First enable motor to move 1 step Anticlockwise direction")

    def clockwise_step(self):
        if self.enable_motor_button["text"] == "Disable Motor":
            self.sendCommand(self.command.clockwiseCommand)
            if self.receiveCommand() == self.command.clockwise_ok:
                self.sendCommand(self.command.pulseCommand)
                if self.receiveCommand() == self.command.pulse_ok:
                    pass
                else:
                    messagebox.showinfo("Position Status", "Reached the End")
        else:
            messagebox.showerror("Motor Status Error", "First enable motor to move 1 step Clockwise direction")

    def enable_motor(self):
        if self.enable_motor_button["text"] == "Disable Motor":
            self.sendCommand(self.command.disableCommand)
            if self.receiveCommand() == self.command.disable_ok:
                self.enable_motor_button["text"] = "Enable Motor"
        else:
            self.sendCommand(self.command.enableCommand)
            if self.receiveCommand() == self.command.enable_ok:
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
            self.sendCommand(self.command.stopCommand)

    def initialize_graph(self):
        self.fig = Figure(figsize=(12, 4.4), dpi=80, facecolor=self.styles.bg_color)
        self.fig.tight_layout()
        self.ax = self.fig.add_subplot(111)
        self.fig.subplots_adjust(left=0.07, right=0.98, top=0.95, bottom=0.09)
        self.ax.set_xlabel("Distance in mm")
        self.ax.set_ylabel("Power in W")
        self.lines = self.ax.plot([], [])[0]
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, rowspan=5)

    def initialize_graph_viewer(self):
        self.load_button = Button(self.graph_view_frame,
                                  text="Import Data",
                                  width=26,
                                  command=lambda: self.view.file_loader(self),
                                  bg=self.styles.button_background,
                                  fg=self.styles.button_foreground,
                                  bd=self.styles.button_border_size,
                                  highlightbackground=self.styles.button_highlight_background,
                                  highlightcolor=self.styles.button_highlight_foreground,
                                  disabledforeground=self.styles.button_disabled_foreground,
                                  activebackground=self.styles.button_active_background,
                                  activeforeground=self.styles.button_active_foreground,
                                  font=self.styles.bold_font)
        self.load_button.grid(row=0, column=0, pady=5)

        self.fig1 = Figure(figsize=(14.7, 6.4), dpi=80, facecolor=self.styles.bg_color)
        # self.fig1.tight_layout()
        self.fig1.subplots_adjust(left=0.06, right=1, top=0.95, bottom=0.09, hspace=0.5, wspace=0.15)

        self.ax1 = self.fig1.add_subplot(221)
        self.ax1.set_xlabel("Distance in mm")
        self.ax1.set_ylabel("Power in W")
        self.ax1.set_title("Open Aperture")
        self.lines1 = self.ax1.plot([], [])[0]

        self.ax2 = self.fig1.add_subplot(222)
        self.ax2.set_xlabel("Distance in mm")
        self.ax2.set_ylabel("Power in W")
        self.ax2.set_title("Normalized Open Aperture")
        self.lines2 = self.ax2.plot([], [])[0]

        self.ax3 = self.fig1.add_subplot(223)
        self.ax3.set_xlabel("Distance in mm")
        self.ax3.set_ylabel("Power in W")
        self.ax3.set_title("Close Aperture")
        self.lines3 = self.ax3.plot([], [])[0]

        self.ax4 = self.fig1.add_subplot(224)
        self.ax4.set_xlabel("Distance in mm")
        self.ax4.set_ylabel("Power in W")
        self.ax4.set_title("Normalized Close Aperture")
        self.lines4 = self.ax4.plot([], [])[0]

        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=self.graph_view_frame)
        self.canvas1.get_tk_widget().grid(row=1, column=0)

        self.seperator7 = ttk.Separator(self.graph_view_frame, orient="horizontal")
        self.seperator7.place(relx=0, rely=0.535, relwidth=1, relheight=0.001)

        self.normalize_button_1 = Button(self.graph_view_frame,
                                         text="Normalize",
                                         width=15,
                                         command=lambda: self.view.normalize(self, 1),
                                         bg=self.styles.button_background,
                                         fg=self.styles.button_foreground,
                                         bd=self.styles.button_border_size,
                                         highlightbackground=self.styles.button_highlight_background,
                                         highlightcolor=self.styles.button_highlight_foreground,
                                         disabledforeground=self.styles.button_disabled_foreground,
                                         activebackground=self.styles.button_active_background,
                                         activeforeground=self.styles.button_active_foreground,
                                         font=self.styles.bold_font,
                                         state="disabled")
        self.normalize_button_1.place(relx=0.89, rely=0.485)

        self.normalize_button_2 = Button(self.graph_view_frame,
                                         text="Normalize",
                                         width=15,
                                         command=lambda: self.view.normalize(self, 2),
                                         bg=self.styles.button_background,
                                         fg=self.styles.button_foreground,
                                         bd=self.styles.button_border_size,
                                         highlightbackground=self.styles.button_highlight_background,
                                         highlightcolor=self.styles.button_highlight_foreground,
                                         disabledforeground=self.styles.button_disabled_foreground,
                                         activebackground=self.styles.button_active_background,
                                         activeforeground=self.styles.button_active_foreground,
                                         font=self.styles.bold_font,
                                         state="disabled")
        self.normalize_button_2.place(relx=0.89, rely=0.949)

        self.viewer_open_x = []
        self.viewer_close_x = []
        self.viewer_normalized_open_x = []
        self.viewer_normalized_close_x = []
        self.viewer_open_y = []
        self.viewer_close_y = []
        self.viewer_normalized_open_y = []
        self.viewer_normalized_close_y = []
        self.initial_slider_value_1 = 0
        self.initial_slider_value_2 = 0

    def value_option_menu(self):
        options = ["mm", "microns"]
        bandwidth_options = ["High", "Low"]
        self.clicked_bandwidth = StringVar()
        self.clicked_home_to_maxima = StringVar()
        self.clicked_sweeping_distance = StringVar()
        self.clicked_step_value = StringVar()
        self.default_clicked_home_to_maxima = StringVar()
        self.default_clicked_sweeping_distance = StringVar()
        self.default_clicked_step_value = StringVar()

        if self.xml[0][1].text == "High":
            self.clicked_bandwidth.set(bandwidth_options[0])
        elif self.xml[0][1].text == "Low":
            self.clicked_bandwidth.set(bandwidth_options[1])

        if self.xml[1][1].text == "mm":
            self.default_clicked_home_to_maxima.set(options[0])
        elif self.xml[1][1].text == "microns":
            self.default_clicked_home_to_maxima.set(options[1])

        if self.xml[1][3].text == "mm":
            self.default_clicked_sweeping_distance.set(options[0])
        elif self.xml[1][3].text == "microns":
            self.default_clicked_sweeping_distance.set(options[1])

        if self.xml[1][5].text == "mm":
            self.default_clicked_step_value.set(options[0])
        elif self.xml[1][5].text == "microns":
            self.default_clicked_step_value.set(options[1])

        if self.default_clicked_home_to_maxima.get() == "mm":
            self.clicked_home_to_maxima.set(options[0])
        elif self.default_clicked_home_to_maxima.get() == "microns":
            self.clicked_home_to_maxima.set(options[1])

        if self.default_clicked_sweeping_distance.get() == "mm":
            self.clicked_sweeping_distance.set(options[0])
        elif self.default_clicked_sweeping_distance.get() == "microns":
            self.clicked_sweeping_distance.set(options[1])

        if self.default_clicked_step_value.get() == "mm":
            self.clicked_step_value.set(options[0])
        elif self.default_clicked_step_value.get() == "microns":
            self.clicked_step_value.set(options[1])

        self.drop_bandwidth = OptionMenu(self.Config, self.clicked_bandwidth, *bandwidth_options)
        self.drop_home_to_maxima = OptionMenu(self.communication_frame, self.clicked_home_to_maxima, *options)
        self.drop_stopvalue = OptionMenu(self.communication_frame, self.clicked_sweeping_distance, *options)
        self.drop_step_value = OptionMenu(self.communication_frame, self.clicked_step_value, *options)
        self.default_drop_home_to_maxima = OptionMenu(self.Config, self.default_clicked_home_to_maxima, *options)
        self.default_drop_stopvalue = OptionMenu(self.Config, self.default_clicked_sweeping_distance, *options)
        self.default_drop_step_value = OptionMenu(self.Config, self.default_clicked_step_value, *options)

        self.drop_bandwidth.config(width=6,
                                   bg=self.styles.drop_background_color,
                                   activebackground=self.styles.active_drop_border_color,
                                   highlightbackground=self.styles.highlight_drop_color,
                                   highlightthickness=self.styles.drop_border_size,
                                   disabledforeground=self.styles.disabled_foreground_drop_color)
        self.drop_home_to_maxima.config(width=9,
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
        self.drop_step_value.config(width=9,
                                   bg=self.styles.drop_background_color,
                                   activebackground=self.styles.active_drop_border_color,
                                   highlightbackground=self.styles.highlight_drop_color,
                                   highlightthickness=self.styles.drop_border_size,
                                   disabledforeground=self.styles.disabled_foreground_drop_color)
        self.default_drop_home_to_maxima.config(width=8,
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
        self.default_drop_step_value.config(width=8,
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
            try:
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
            except Exception as e:
                messagebox.showerror("Connection Error", e)
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
                        if self.clicked_home_to_maxima.get() == "mm":
                            maxima_point_from_home_in_mm = float(self.home_to_maxima_entry.get())
                        else:
                            maxima_point_from_home_in_mm = (float(self.home_to_maxima_entry.get()) / 1000)
                        if self.clicked_sweeping_distance.get() == "mm":
                            sweeping_distance_in_mm = float(self.sweeping_distance_entry.get())
                        else:
                            sweeping_distance_in_mm = (float(self.sweeping_distance_entry.get()) / 1000)
                        if self.clicked_step_value.get() == "mm":
                            step_value_in_mm = float(self.step_value_entry.get())
                        else:
                            step_value_in_mm = (float(self.step_value_entry.get()) / 1000)
                        self.sendCommand(self.command.clockwiseCommand)
                        if self.receiveCommand() == self.command.clockwise_ok:
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
                                self.lines.set_xdata(self.x_data)
                                self.lines.set_ydata(self.y_data)
                                if self.run_button_without_aperture["text"] == "Stop Run":
                                    self.lines.set_color("blue")
                                elif self.run_button_with_aperture["text"] == "Stop Run":
                                    self.lines.set_color("red")
                                self.canvas.draw()
                                self.no_of_total_steps = round(2 * sweeping_distance_in_mm / step_value_in_mm)
                                counter = self.no_of_total_steps
                                self.total_step_value["text"] = str(self.no_of_total_steps)
                                while current_point > final_point and self.run_threading_status:
                                    self.steps_remaining_value["text"] = str(round(counter))
                                    self.steps_completed_value["text"] = str(round(self.no_of_total_steps - counter))
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

                                    current_point = current_point - step_value_in_mm
                                    if self.run_threading_status:
                                        self.take_position(step_value_in_mm)
                                        counter = counter - 1

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
                        self.viewer_open_x = self.x_data
                        self.viewer_open_y = self.y_data

                        self.lines1.set_xdata([])
                        self.lines1.set_ydata([])
                        self.canvas1.draw()

                        self.lines1.set_xdata(self.viewer_open_x)
                        self.lines1.set_ydata(self.viewer_open_y)
                        self.ax1.set_ylim([min(self.viewer_open_y) - (min(self.viewer_open_y) * 0.01),
                                           max(self.viewer_open_y) + (max(self.viewer_open_y) * 0.01)])
                        self.ax1.set_xlim([min(self.viewer_open_x), max(self.viewer_open_x)])
                        self.lines2.set_xdata([])
                        self.lines2.set_ydata([])
                        self.normalize_button_1["state"] = "normal"
                        self.canvas1.draw()

                    elif self.run_button_with_aperture["text"] == "Stop Run":
                        self.with_aperture_x_data = self.x_data
                        self.with_aperture_y_data = self.y_data
                        self.viewer_close_x = self.x_data
                        self.viewer_close_y = self.y_data

                        self.lines3.set_xdata([])
                        self.lines3.set_ydata([])
                        self.canvas1.draw()

                        self.lines3.set_xdata(self.viewer_close_x)
                        self.lines3.set_ydata(self.viewer_close_y)
                        self.ax3.set_ylim([min(self.viewer_close_y) - (min(self.viewer_close_y) * 0.05),
                                           max(self.viewer_close_y) + (max(self.viewer_close_y) * 0.05)])
                        self.ax3.set_xlim([min(self.viewer_close_x), max(self.viewer_close_x)])
                        self.lines4.set_xdata([])
                        self.lines4.set_ydata([])
                        self.normalize_button_2["state"] = "normal"
                        self.canvas1.draw()

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

                    self.calculation.update_values(self)
                    self.calculation.update_button["state"] = "normal"
                    self.Home()

    def sendCommand(self, command):
        # print("S= " + command)
        self.ser.reset_output_buffer()
        self.ser.write(command.encode("utf-8"))

    def receiveCommand(self):
        self.tester = 0
        while self.tester == 0:
            while self.ser.in_waiting > 0:
                msg = str(self.ser.readline(), "utf-8")
                self.ser.reset_input_buffer()
                self.tester = 1
                # print("R= " + msg)
                return msg

    def homeThreading(self):
        self.sendCommand(self.command.antiClockwiseCommand)
        if self.receiveCommand() == self.command.antiClockwise_ok:
            self.notebook.hide(1)
            self.sendCommand(self.command.homeCommand)
            received_command = self.receiveCommand()
            if received_command == self.command.home_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Position Status", "Home Reached")
                self.status_retainer()
            elif received_command == self.command.stop_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                self.status_retainer()

    def endThreading(self):
        self.sendCommand(self.command.clockwiseCommand)
        if self.receiveCommand() == self.command.clockwise_ok:
            self.notebook.hide(1)
            self.sendCommand(self.command.endCommand)
            received_command = self.receiveCommand()
            if received_command == self.command.end_ok:
                self.notebook.add(self.settingsframe, text="Settings")
                messagebox.showinfo("Position Status", "Reached the End")
                self.status_retainer()
            elif received_command == self.command.stop_ok:
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
                self.sendCommand(self.command.clockwiseCommand)
                if self.receiveCommand() == self.command.clockwise_ok:
                    command = "move:" + str(round(pulse_number)) + "\n"
                    self.sendCommand(command)
                    received_command = self.receiveCommand()
            else:
                self.sendCommand(self.command.antiClockwiseCommand)
                if self.receiveCommand() == self.command.antiClockwise_ok:
                    command = "move:" + str(round(pulse_number)) + "\n"
                    self.sendCommand(command)
                    received_command = self.receiveCommand()
            if received_command == self.command.home_ok:
                messagebox.showinfo("Position Status", "Home Reached")
            elif received_command == self.command.end_ok:
                messagebox.showinfo("Position Status", "End Reached")
            elif received_command == self.command.stop_ok or received_command == self.command.move_ok:
                pass
            self.move_retainer()

    def move_retainer(self):
        self.move_button["text"] = "Move"
        self.enable_manual_mode["state"] = "normal"
        self.enable_motor_button["state"] = "normal"

    def locateThreading(self):
        self.sendCommand(self.command.antiClockwiseCommand)
        if self.receiveCommand() == self.command.antiClockwise_ok:
            self.sendCommand(self.command.locateCommand)
            self.holder_position["text"] = 'Please wait...'
            command = self.receiveCommand()
            if command == self.command.stop_ok:
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
        if len(self.home_to_maxima_entry.get()) != 0:
            if self.num_validator(self.home_to_maxima_entry.get()):
                if round(float(self.home_to_maxima_entry.get()), 4) > 0:
                    if len(self.sweeping_distance_entry.get()) != 0:
                        if self.num_validator(self.sweeping_distance_entry.get()):
                            if round(float(self.sweeping_distance_entry.get()), 4) > 0:
                                if len(self.step_value_entry.get()) != 0:
                                    if self.num_validator(self.sweeping_distance_entry.get()):
                                        if round(float(self.step_value_entry.get()), 4) > 0:
                                            home_to_maxima = float(self.home_to_maxima_entry.get())
                                            sweeping_distance = float(self.sweeping_distance_entry.get())
                                            step_value = float(self.step_value_entry.get())
                                            if self.clicked_home_to_maxima.get() == "mm":
                                                home_to_maxima = home_to_maxima * 1000
                                            if self.clicked_sweeping_distance.get() == "mm":
                                                sweeping_distance = sweeping_distance * 1000
                                            if self.clicked_step_value.get() == "mm":
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
        self.sendCommand(self.command.antiClockwiseCommand)
        if self.receiveCommand() == self.command.antiClockwise_ok:
            self.sendCommand(self.command.homeCommand)
            command = self.receiveCommand()
            if command == self.command.home_ok:
                return True
            elif command == self.command.stop_ok:
                return False

    def take_position(self, pulse_number):
        pulse_number = pulse_number * self.motor_steps
        command = "move:" + str(round(pulse_number)) + "\n"
        self.sendCommand(command)
        received_command = self.receiveCommand()
        if received_command == self.command.move_ok:
            return True
        elif received_command == self.command.stop_ok:
            return False
        elif received_command == self.command.end_ok:
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

    def status_updater(self):
        if self.home_button["text"] == "Home":
            self.home_button["state"] = "disabled"
        if self.end_button["text"] == "End":
            self.end_button["state"] = "disabled"
        if self.run_button_without_aperture["text"] == "Start Run without Aperture":
            self.run_button_without_aperture["state"] = "disabled"
        if self.run_button_with_aperture["text"] == "Start Run with Aperture":
            self.run_button_with_aperture["state"] = "disabled"
        self.home_to_maxima_entry["state"] = "disabled"
        self.sweeping_distance_entry["state"] = "disabled"
        self.step_value_entry["state"] = "disabled"
        self.com_btn_connect["state"] = "disabled"
        self.drop_home_to_maxima["state"] = "disabled"
        self.drop_stopvalue["state"] = "disabled"
        self.drop_step_value["state"] = "disabled"

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
        self.home_to_maxima_entry["state"] = "normal"
        self.sweeping_distance_entry["state"] = "normal"
        self.step_value_entry["state"] = "normal"
        self.com_btn_connect["state"] = "normal"
        self.drop_home_to_maxima["state"] = "normal"
        self.drop_stopvalue["state"] = "normal"
        self.drop_step_value["state"] = "normal"
        self.notebook.add(self.settingsframe, text="Settings")

    def Home(self):
        if self.home_button["text"] == "Home":
            self.t1 = threading.Thread(target=self.homeThreading, daemon=True)
            self.t1.start()
            self.home_button["text"] = "Stop"
            self.status_updater()
        else:
            self.sendCommand(self.command.stopCommand)

    def End(self):
        if self.end_button["text"] == "End":
            self.t1 = threading.Thread(target=self.endThreading, daemon=True)
            self.t1.start()
            self.end_button["text"] = "Stop"
            self.status_updater()
        else:
            self.sendCommand(self.command.stopCommand)

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
            self.sendCommand(self.command.stopCommand)
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
            self.sendCommand(self.command.stopCommand)
            self.run_threading_status = False

    def num_validator(self, received_data):
        self.data = received_data
        try:
            float(self.data)
            return True
        except:
            return False

    def save_button(self):
        def error_thrower(variable, name):
            if len(variable) > 0:
                if self.num_validator(variable):
                    return True
                else:
                    messagebox.showerror("Numerical Error", f"{name} should be a number")
                    return False
            else:
                messagebox.showerror("Empty String Error", f"{name} field should not be blank")
                return False
        if (error_thrower(self.wavelength_entry.get(), "Wavelength")
                and error_thrower(self.laser_beam_input_power_entry.get(), "Power of the Laser")
                and error_thrower(self.laser_beam_diameter_entry.get(), "Diameter of the Laser beam")
                and error_thrower(self.transmittance_entry.get(), "Transmittance")
                and error_thrower(self.transmittance_thickness_entry.get(), "Thickness at Transmittance measurement")
                and error_thrower(self.beam_rad_at_aperture_entry.get(), "Radius of beam at aperture")
                and error_thrower(self.z_scan_sample_thickness_entry.get(), "Z-Scan sample thickness")
                and error_thrower(self.linear_refractive_index_entry.get(), "Linear refractive index")
                and error_thrower(self.thread_spacing_entry.get(), "Thread Spacing")
                and error_thrower(self.pulse_width_entry.get(), "Pulse Width")
                and error_thrower(self.steps_per_rotation_entry.get(), "Steps per Rotation")
                and error_thrower(self.default_home_to_maxima_entry.get(), "Default Home to maxima")
                and error_thrower(self.default_stopvalue_entry.get(), "Half Sweeping Distance")
                and error_thrower(self.default_step_value_entry.get(), "Step Value")
                and error_thrower(self.focal_length_entry.get(), "Focal length of the lens")
                and error_thrower(self.close_aperture_entry.get(), "Radius of the aperture")):
            if 400 <= float(self.wavelength_entry.get()) <= 1100:
                self.xml[0][0].text = self.wavelength_entry.get()
                self.xml[0][1].text = self.clicked_bandwidth.get()
                self.xml[0][2].text = self.thread_spacing_entry.get()
                self.xml[0][3].text = self.laser_beam_input_power_entry.get()
                self.xml[0][4].text = self.laser_beam_diameter_entry.get()
                self.xml[0][5].text = self.transmittance_entry.get()
                self.xml[0][6].text = self.transmittance_thickness_entry.get()
                self.xml[0][7].text = self.beam_rad_at_aperture_entry.get()
                self.xml[0][8].text = self.z_scan_sample_thickness_entry.get()
                self.xml[0][9].text = self.linear_refractive_index_entry.get()
                self.xml[1][0].text = self.default_home_to_maxima_entry.get()
                self.xml[1][1].text = self.default_clicked_home_to_maxima.get()
                self.xml[1][2].text = self.default_stopvalue_entry.get()
                self.xml[1][3].text = self.default_clicked_sweeping_distance.get()
                self.xml[1][4].text = self.default_step_value_entry.get()
                self.xml[1][5].text = self.default_clicked_step_value.get()
                self.xml[1][6].text = self.focal_length_entry.get()
                self.xml[1][7].text = self.close_aperture_entry.get()
                self.xml[1][8].text = self.pulse_width_entry.get()
                self.xml[1][9].text = self.steps_per_rotation_entry.get()

                tree.write("data_file.xml")
                messagebox.showinfo("Save status", "Data saved successfully")
            else:
                messagebox.showerror("Value Error", "Wavelength value must in between 400 nm to 1100 nm")

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
        self.home_to_maxima_entry["state"] = "normal"
        self.sweeping_distance_entry["state"] = "normal"
        self.step_value_entry["state"] = "normal"
        self.drop_home_to_maxima["state"] = "normal"
        self.drop_stopvalue["state"] = "normal"
        self.drop_step_value["state"] = "normal"

    def disable_buttons(self):
        self.home_button["state"] = "disabled"
        self.end_button["state"] = "disabled"
        self.run_button_without_aperture["state"] = "disabled"
        self.run_button_with_aperture["state"] = "disabled"
        self.home_to_maxima_entry["state"] = "disabled"
        self.sweeping_distance_entry["state"] = "disabled"
        self.step_value_entry["state"] = "disabled"
        self.drop_home_to_maxima["state"] = "disable"
        self.drop_stopvalue["state"] = "disable"
        self.drop_step_value["state"] = "disable"
        # self.save_data_button["state"] = "disabled"
        # self.enter_file_name_entry["state"] = "disabled"
        self.reset_all_button["state"] = "disabled"
        self.save_image_data_button["state"] = "disabled"

    def publish(self):
        self.notebook.grid(row=0, column=0)

        self.mainframe.grid(row=0, column=0)
        self.settingsframe.grid(row=0, column=0)
        self.graph_view_frame.grid(row=0, column=0)
        self.calculation_frame.grid(row=0, column=0)

        self.notebook.add(self.mainframe, text="Controls")
        self.notebook.add(self.settingsframe, text="Settings")
        self.notebook.add(self.graph_view_frame, text="Viewer")
        self.notebook.add(self.calculation_frame, text="Calculations")

        self.connection_frame.grid(row=0, column=0, columnspan=3, rowspan=3, padx=5, pady=5)
        self.connection_frame.grid_propagate(False)

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
        self.home_to_maxima.grid(row=0, column=0)
        self.sweeping_distance.grid(row=1, column=0)
        self.step_value.grid(row=2, column=0)
        self.home_to_maxima_entry.grid(row=0, column=1)
        self.sweeping_distance_entry.grid(row=1, column=1)
        self.step_value_entry.grid(row=2, column=1)
        self.drop_home_to_maxima.grid(row=0, column=2, padx=15)
        self.drop_stopvalue.grid(row=1, column=2, padx=15)
        self.drop_step_value.grid(row=2, column=2, padx=15)
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
        self.laser_beam_input_power.grid(row=1, column=0)
        self.laser_beam_diameter.grid(row=2, column=0)
        self.transmittance.grid(row=3, column=0)
        self.transmittance_thickness.grid(row=4, column=0)
        self.bandwidth.grid(row=5, column=0)
        self.thread_spacing.grid(row=6, column=0)
        self.pulse_width_label.grid(row=7, column=0)
        self.steps_per_rotation_label.grid(row=8, column=0)
        self.focal_length.grid(row=0, column=4)
        self.close_aperture_label.grid(row=1, column=4)
        self.beam_rad_at_aperture_label.grid(row=2, column=4)
        self.z_scan_sample_thickness_label.grid(row=3, column=4)
        self.linear_refractive_index_label.grid(row=4, column=4)
        self.default_home_to_maxima.grid(row=5, column=4)
        self.default_stopvalue.grid(row=6, column=4)
        self.default_step_value.grid(row=7, column=4)

        self.wavelength_entry.grid(row=0, column=1, padx=(0, 7))
        self.laser_beam_input_power_entry.grid(row=1, column=1, padx=(0, 7))
        self.laser_beam_diameter_entry.grid(row=2, column=1, padx=(0, 7))
        self.transmittance_entry.grid(row=3, column=1, padx=(0, 7))
        self.transmittance_thickness_entry.grid(row=4, column=1, padx=(0, 7))
        self.drop_bandwidth.grid(row=5, column=1, padx=(0, 7))
        self.thread_spacing_entry.grid(row=6, column=1, padx=(0, 7))
        self.pulse_width_entry.grid(row=7, column=1, padx=(0, 7))
        self.steps_per_rotation_entry.grid(row=8, column=1, padx=(0, 7))
        self.focal_length_entry.grid(row=0, column=5, padx=(0, 7))
        self.close_aperture_entry.grid(row=1, column=5, padx=(0, 7))
        self.beam_rad_at_aperture_entry.grid(row=2, column=5, padx=(0, 7))
        self.z_scan_sample_thickness_entry.grid(row=3, column=5, padx=(0, 7))
        self.linear_refractive_index_entry.grid(row=4, column=5, padx=(0, 7))
        self.default_home_to_maxima_entry.grid(row=5, column=5, padx=(0, 7))
        self.default_stopvalue_entry.grid(row=6, column=5, padx=(0, 7))
        self.default_step_value_entry.grid(row=7, column=5, padx=(0, 7))

        self.wavelength_unit.grid(row=0, column=2)
        self.laser_beam_input_power_unit.grid(row=1, column=2)
        self.laser_beam_diameter_unit.grid(row=2, column=2)
        self.transmittance_unit.grid(row=3, column=2)
        self.transmittance_thickness_unit.grid(row=4, column=2)
        self.thread_spacing_unit.grid(row=6, column=2)
        self.pulse_width_unit.grid(row=7, column=2)
        self.focal_length_unit.grid(row=0, column=6)
        self.close_aperture_unit.grid(row=1, column=6)
        self.beam_rad_at_aperture_unit.grid(row=2, column=6)
        self.z_scan_sample_thickness_unit.grid(row=3, column=6)
        self.default_drop_home_to_maxima.grid(row=5, column=6)
        self.default_drop_stopvalue.grid(row=6, column=6)
        self.default_drop_step_value.grid(row=7, column=6)

        self.seperator2.place(relx=0.46, rely=0, relwidth=0.003, relheight=1)

        self.save_button.grid(row=0, column=1, padx=5)
        self.apply_button.grid(row=0, column=2, padx=5)

        self.button_frame.grid(row=8, column=4, columnspan=3)

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
