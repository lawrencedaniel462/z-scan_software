from tkinter import *
from tkinter import ttk, messagebox, filedialog
import math


def get_file(self):
    if len(self.file_name_entry.get()) == 0:
        messagebox.showerror("Naming Error", "File name should not be empty")
        return False
    else:
        return True
def remove_last(string, end):
    string = string[:len(string) - end]
    return string


class Calculation:
    def initialize_calculation(self, gui):
        self.input_frame = Frame(gui.calculation_frame,
                                     padx=5,
                                     pady=5,
                                     width=420,
                                     height=480,
                                     bg=gui.styles.bg_color)
        self.input_frame.grid_propagate(False)
        self.output_frame = Frame(gui.calculation_frame,
                                  padx=5,
                                  pady=5,
                                  width=450,
                                  height=490,
                                  bg=gui.styles.bg_color)
        self.output_frame.grid_propagate(False)
        self.button_frame = Frame(gui.calculation_frame,
                                  padx=5,
                                  pady=5,
                                  width=850,
                                  height=32,
                                  bg=gui.styles.bg_color)
        # self.button_frame.grid_propagate(False)

        self.seperator = ttk.Separator(gui.calculation_frame, orient="horizontal")
        self.seperator.place(relx=0.335, rely=0, relwidth=0.002, relheight=0.85)

        self.wavelength = Label(self.input_frame, text="Wavelength of the Laser (λ)", anchor="w", width=35, padx=7, pady=8,
                                bg=gui.styles.bg_color)
        self.laser_beam_input_power = Label(self.input_frame, text="Power of the Laser (E\u209A)", anchor="w", width=35,
                                            padx=7, pady=8, bg=gui.styles.bg_color)
        self.laser_beam_diameter = Label(self.input_frame, text="Diameter of the Laser beam (d)", anchor="w", width=35,
                                         padx=7, pady=8, bg=gui.styles.bg_color)
        self.transmittance = Label(self.input_frame, text="Transmittance (T)", anchor="w", width=35,
                                   padx=7, pady=8, bg=gui.styles.bg_color)
        self.transmittance_thickness = Label(self.input_frame, text="Thickness at transmittance measurement (t)", anchor="w",
                                             width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.focal_length = Label(self.input_frame, text="Focal length of the lens (f)", anchor="w", width=35, padx=7,
                                  pady=8, bg=gui.styles.bg_color)
        self.close_aperture = Label(self.input_frame, text="Radius of the Aperture (r\u2090)",
                                    anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.beam_rad_at_aperture = Label(self.input_frame, text="Radius of the beam at aperture (w\u2090)",
                                          anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.z_scan_sample_thickness = Label(self.input_frame, text="Z-scan sample thickness (L)",
                                             anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.linear_refractive_index = Label(self.input_frame, text="Linear refractive index (n\u2080)",
                                             anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi_close = Label(self.input_frame, text="Peak to valley difference of\nnormalized close aperture graph  (\u0394T\u209A\u208B\u1D65)",
                                             anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi_open = Label(self.input_frame,
                                   text="Peak to valley difference of\nnormalized open aperture graph  (\u0394T\u1D65\u208B\u209A)",
                                   anchor="w", width=35, padx=7, pady=8, bg=gui.styles.bg_color)

        self.wavelength_value = Label(self.input_frame, text="-", anchor="w", width=15, padx=7,
                                pady=8,
                                bg=gui.styles.bg_color)
        self.laser_beam_input_power_value = Label(self.input_frame, text="-", anchor="w",
                                            width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.laser_beam_diameter_value = Label(self.input_frame, text="-", anchor="w",
                                         width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.transmittance_value = Label(self.input_frame, text="-", anchor="w", width=15,
                                   padx=7, pady=8, bg=gui.styles.bg_color)
        self.transmittance_thickness_value = Label(self.input_frame, text="-",
                                             anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.focal_length_value = Label(self.input_frame, text="-", anchor="w", width=15,
                                  padx=7, pady=8, bg=gui.styles.bg_color)
        self.close_aperture_value = Label(self.input_frame, text="-",
                                    anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.beam_rad_at_aperture_value = Label(self.input_frame, text="-",
                                          anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.z_scan_sample_thickness_value = Label(self.input_frame, text="-",
                                             anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.linear_refractive_index_value = Label(self.input_frame, text="-",
                                             anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi_close_value = Label(self.input_frame,
                                   text="-",
                                   anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi_open_value = Label(self.input_frame,
                                  text="-",
                                  anchor="w", width=15, padx=7, pady=8, bg=gui.styles.bg_color)

        self.wavelength_unit = " nm"
        self.laser_beam_input_power_unit =" w"
        self.laser_beam_diameter_unit = " mm"
        self.transmittance_unit = " %"
        self.transmittance_thickness_unit = " mm"
        self.focal_length_unit = " mm"
        self.close_aperture_unit = " mm"
        self.beam_rad_at_aperture_unit = " mm"
        self.z_scan_sample_thickness_unit = " mm"

        self.update_button = Button(self.input_frame,
                                    text="Update",
                                    width=26,
                                    command=lambda: self.update_values(gui),
                                    bg=gui.styles.button_background,
                                    fg=gui.styles.button_foreground,
                                    bd=gui.styles.button_border_size,
                                    highlightbackground=gui.styles.button_highlight_background,
                                    highlightcolor=gui.styles.button_highlight_foreground,
                                    disabledforeground=gui.styles.button_disabled_foreground,
                                    activebackground=gui.styles.button_active_background,
                                    activeforeground=gui.styles.button_active_foreground,
                                    font=gui.styles.bold_font)
        self.update_button["state"] = "disabled"

        self.update_values(gui)

        self.alpha = Label(self.output_frame, text="Linear absorption coefficient (\u03B1)", anchor="w",
                           width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.Omega_not = Label(self.output_frame, text="Beam waist at focal point (2\u03C9\u2080)", anchor="w",
                               width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.I_not = Label(self.output_frame, text="(I\u2080)", anchor="w",
                           width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.ZR = Label(self.output_frame, text="Rayleigh length (Z\u1D63)", anchor="w",
                        width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.S = Label(self.output_frame, text="Linear transmittance (S)", anchor="w",
                       width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi = Label(self.output_frame, text="Axis phase shift at focus (|\u0394\u03A6|)", anchor="w",
                             width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.leff = Label(self.output_frame, text="Effective Length (L\u2091)", anchor="w",
                          width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.k = Label(self.output_frame, text="(k)", anchor="w",
                       width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.n2 = Label(self.output_frame, text="Non linear refractive index (n²)", anchor="w",
                        width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.beta = Label(self.output_frame, text="Non linear absorption coefficient (β)", anchor="w",
                          width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.real = Label(self.output_frame, text="Real part of Susceptibility (Re χ⁽³⁾)", anchor="w",
                          width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.imaginary = Label(self.output_frame, text="Imaginary part of Susceptibility (Im χ⁽³⁾)", anchor="w",
                               width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.susceptibility = Label(self.output_frame, text="Susceptibility (χ⁽³⁾)", anchor="w",
                                    width=35, padx=7, pady=8, bg=gui.styles.bg_color)
        self.file_name = Label(self.output_frame, text="Enter file name", anchor="w",
                                    width=35, padx=7, pady=8, bg=gui.styles.bg_color)

        width = 20
        self.alpha_value = Label(self.output_frame, text="-", anchor="w",
                                 width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.Omega_not_value = Label(self.output_frame, text="-", anchor="w",
                                     width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.I_not_value = Label(self.output_frame, text="-", anchor="w",
                                 width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.ZR_value = Label(self.output_frame, text="-", anchor="w",
                              width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.S_value = Label(self.output_frame, text="-", anchor="w",
                             width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.del_phi_value = Label(self.output_frame, text="-", anchor="w",
                                   width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.leff_value = Label(self.output_frame, text="-", anchor="w",
                                width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.k_value = Label(self.output_frame, text="-", anchor="w",
                             width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.n2_value = Label(self.output_frame, text="-", anchor="w",
                              width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.beta_value = Label(self.output_frame, text="-", anchor="w",
                                width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.real_value = Label(self.output_frame, text="-", anchor="w",
                                width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.imaginary_value = Label(self.output_frame, text="-", anchor="w",
                                     width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.susceptibility_value = Label(self.output_frame, text="-", anchor="w",
                                          width=width, padx=7, pady=8, bg=gui.styles.bg_color)
        self.file_name_entry = Entry(self.output_frame,
                                      width=19,
                                      font=('Helvetica', 10),
                                      bg=gui.styles.entry_background_color,
                                      highlightbackground=gui.styles.entry_border_color,
                                      highlightthickness=gui.styles.entry_border_size,
                                      highlightcolor=gui.styles.active_entry_border_color,
                                      disabledbackground=gui.styles.disabled_background_entry,
                                      disabledforeground=gui.styles.disabled_foreground_entry)

        self.alpha_unit = "m⁻¹"
        self.Omega_not_unit = "m"
        self.I_not_unit = "W/m²"
        self.ZR_unit = "mm"
        self.S_unit = ""
        self.del_phi_unit = ""
        self.leff_unit = "m"
        self.k_unit = "m⁻¹"
        self.n2_unit = "m²/W"
        self.beta_unit = "mW"
        self.real_unit = "cm²/W"
        self.imaginary_unit = "cm/W"
        self.susceptibility_unit = "(esu)"

        self.calculate_button = Button(self.button_frame,
                                       text="Calculate",
                                       width=23,
                                       command=self.calculate,
                                       bg=gui.styles.button_background,
                                       fg=gui.styles.button_foreground,
                                       bd=gui.styles.button_border_size,
                                       highlightbackground=gui.styles.button_highlight_background,
                                       highlightcolor=gui.styles.button_highlight_foreground,
                                       disabledforeground=gui.styles.button_disabled_foreground,
                                       activebackground=gui.styles.button_active_background,
                                       activeforeground=gui.styles.button_active_foreground,
                                       font=gui.styles.bold_font)
        self.export_data_button = Button(self.button_frame,
                                       text="Export data",
                                       width=23,
                                       command=lambda: self.export_excel_data(gui),
                                       bg=gui.styles.button_background,
                                       fg=gui.styles.button_foreground,
                                       bd=gui.styles.button_border_size,
                                       highlightbackground=gui.styles.button_highlight_background,
                                       highlightcolor=gui.styles.button_highlight_foreground,
                                       disabledforeground=gui.styles.button_disabled_foreground,
                                       activebackground=gui.styles.button_active_background,
                                       activeforeground=gui.styles.button_active_foreground,
                                       font=gui.styles.bold_font)
        self.export_graph_button = Button(self.button_frame,
                                         text="Export graph",
                                         width=23,
                                         command=lambda: self.update_values(gui),
                                         bg=gui.styles.button_background,
                                         fg=gui.styles.button_foreground,
                                         bd=gui.styles.button_border_size,
                                         highlightbackground=gui.styles.button_highlight_background,
                                         highlightcolor=gui.styles.button_highlight_foreground,
                                         disabledforeground=gui.styles.button_disabled_foreground,
                                         activebackground=gui.styles.button_active_background,
                                         activeforeground=gui.styles.button_active_foreground,
                                         font=gui.styles.bold_font)
        self.export_custom_graph_button = Button(self.button_frame,
                                          text="Export custom graph",
                                          width=23,
                                          command=lambda: self.update_values(gui),
                                          bg=gui.styles.button_background,
                                          fg=gui.styles.button_foreground,
                                          bd=gui.styles.button_border_size,
                                          highlightbackground=gui.styles.button_highlight_background,
                                          highlightcolor=gui.styles.button_highlight_foreground,
                                          disabledforeground=gui.styles.button_disabled_foreground,
                                          activebackground=gui.styles.button_active_background,
                                          activeforeground=gui.styles.button_active_foreground,
                                          font=gui.styles.bold_font)

        # (⁰ ¹ ² ³ ⁴ ⁵ ⁶ ⁷ ⁸ ⁹ ⁺ ⁻ ⁼ ⁽ ⁾ ₀ ₁ ₂ ₃ ₄ ₅ ₆ ₇ ₈ ₉ ₊ ₋ ₌ ₍ ₎),
        # a full superscript Latin lowercase alphabet except q(ᵃ ᵇ ᶜ ᵈ ᵉ ᶠ ᵍ ʰ ⁱ ʲ ᵏ ˡ ᵐ ⁿ ᵒ ᵖ ʳ ˢ ᵗ ᵘ ᵛ ʷ ˣ ʸ ᶻ ),
        # a limited uppercase Latin alphabet(ᴬ ᴮ ᴰ ᴱ ᴳ ᴴ ᴵ ᴶ ᴷ ᴸ ᴹ ᴺ ᴼ ᴾ ᴿ ᵀ ᵁ ⱽ ᵂ ),
        # a few subscribed lowercase letters(ₐ ₑ ₕ ᵢ ⱼ ₖ ₗ ₘ ₙ ₒ ₚ ᵣ ₛ ₜ ᵤ ᵥ ₓ ),
        # and some Greek letters(ᵅ ᵝ ᵞ ᵟ ᵋ ᶿ ᶥ ᶲ ᵠ ᵡ ᵦ ᵧ ᵨ ᵩ ᵪ ).

        self.input_frame.grid(row=0, column=0)
        self.output_frame.grid(row=0, column=1)
        self.button_frame.grid(row=1, column=0, columnspan=2, pady=(20, 0))

        self.wavelength.grid(row=0, column=0)
        self.laser_beam_input_power.grid(row=1, column=0)
        self.laser_beam_diameter.grid(row=2, column=0)
        self.transmittance.grid(row=3, column=0)
        self.transmittance_thickness.grid(row=4, column=0)
        self.focal_length.grid(row=5, column=0)
        self.close_aperture.grid(row=6, column=0)
        self.beam_rad_at_aperture.grid(row=7, column=0)
        self.z_scan_sample_thickness.grid(row=8, column=0)
        self.linear_refractive_index.grid(row=9, column=0)
        self.del_phi_close.grid(row=11, column=0)
        self.del_phi_open.grid(row=12, column=0)

        self.wavelength_value.grid(row=0, column=1)
        self.laser_beam_input_power_value.grid(row=1, column=1)
        self.laser_beam_diameter_value.grid(row=2, column=1)
        self.transmittance_value.grid(row=3, column=1)
        self.transmittance_thickness_value.grid(row=4, column=1)
        self.focal_length_value.grid(row=5, column=1)
        self.close_aperture_value.grid(row=6, column=1)
        self.beam_rad_at_aperture_value.grid(row=7, column=1)
        self.z_scan_sample_thickness_value.grid(row=8, column=1)
        self.linear_refractive_index_value.grid(row=9, column=1)
        self.del_phi_close_value.grid(row=11, column=1)
        self.del_phi_open_value.grid(row=12, column=1)

        self.update_button.grid(row=10, column=0, columnspan=2)

        self.alpha.grid(row=0, column=0)
        self.Omega_not.grid(row=1, column=0)
        self.I_not.grid(row=2, column=0)
        self.ZR.grid(row=3, column=0)
        self.S.grid(row=4, column=0)
        self.del_phi.grid(row=5, column=0)
        self.leff.grid(row=6, column=0)
        self.k.grid(row=7, column=0)
        self.n2.grid(row=8, column=0)
        self.beta.grid(row=9, column=0)
        self.real.grid(row=10, column=0)
        self.imaginary.grid(row=11, column=0)
        self.susceptibility.grid(row=12, column=0)
        self.file_name.grid(row=13, column=0)

        self.alpha_value.grid(row=0, column=1)
        self.Omega_not_value.grid(row=1, column=1)
        self.I_not_value.grid(row=2, column=1)
        self.ZR_value.grid(row=3, column=1)
        self.S_value.grid(row=4, column=1)
        self.del_phi_value.grid(row=5, column=1)
        self.leff_value.grid(row=6, column=1)
        self.k_value.grid(row=7, column=1)
        self.n2_value.grid(row=8, column=1)
        self.beta_value.grid(row=9, column=1)
        self.real_value.grid(row=10, column=1)
        self.imaginary_value.grid(row=11, column=1)
        self.susceptibility_value.grid(row=12, column=1)
        self.file_name_entry.grid(row=13, column=1)

        self.calculate_button.grid(row=0, column=0, padx=5)
        self.export_data_button.grid(row=0, column=1, padx=5)
        self.export_graph_button.grid(row=0, column=2, padx=5)
        self.export_custom_graph_button.grid(row=0, column=3, padx=5)

    def update_values(self,gui):
        self.wavelength_value["text"] = gui.wavelength_entry.get() + self.wavelength_unit
        self.laser_beam_input_power_value["text"] = gui.laser_beam_input_power_entry.get() + self.laser_beam_input_power_unit
        self.laser_beam_diameter_value["text"] = gui.laser_beam_diameter_entry.get() + self.laser_beam_diameter_unit
        self.transmittance_value["text"] = gui.transmittance_entry.get() + self.transmittance_unit
        self.transmittance_thickness_value["text"] = gui.transmittance_thickness_entry.get() + self.transmittance_thickness_unit
        self.focal_length_value["text"] = gui.focal_length_entry.get() + self.focal_length_unit
        self.close_aperture_value["text"] = gui.close_aperture_entry.get() + self.close_aperture_unit
        self.beam_rad_at_aperture_value["text"] = gui.beam_rad_at_aperture_entry.get() + self.beam_rad_at_aperture_unit
        self.z_scan_sample_thickness_value["text"] = gui.z_scan_sample_thickness_entry.get() + self.z_scan_sample_thickness_unit
        self.linear_refractive_index_value["text"] = gui.linear_refractive_index_entry.get()

    def calculate(self):
        transmittance = 0
        transmittance_thickness = 0
        wavelength = 0
        focal_length = 0
        laser_beam_diameter = 0
        laser_power = 0
        beam_rad_at_aperture = 0
        close_aperture = 0
        VtoP = 0
        z_scan_sample_thickness = 0
        NtoV = 0
        n_not = 0
        epsilon = 8.854187817 * 10E-12
        C = 299792458

        alpha = 0
        I_not = 0
        S = 0
        del_phi = 0
        leff = 0
        k = 0
        n2 = 0
        beta = 0
        real = 0
        imaginary = 0

        try:
            transmittance = float(remove_last(self.transmittance_value["text"], 2))
            transmittance_thickness = float(remove_last(self.transmittance_thickness_value["text"], 3))
        except:
            pass
        try:
            alpha = (2.306 * 1000 / transmittance_thickness) * math.log10(100 / transmittance)
            self.alpha_value["text"] = f"{alpha} " + self.alpha_unit
        except:
            pass

        try:
            wavelength = float(remove_last(self.wavelength_value["text"], 3))
            focal_length = float(remove_last(self.focal_length_value["text"], 3))
            laser_beam_diameter = float(remove_last(self.laser_beam_diameter_value["text"], 3))
        except:
            pass
        try:
            omega_not = (4 * wavelength * focal_length * 10E-9)/(math.pi * laser_beam_diameter)
            self.Omega_not_value["text"] = f"{omega_not} " + self.Omega_not_unit
        except:
            pass

        try:
            laser_power = float(remove_last(self.laser_beam_input_power_value["text"], 2))
            beam_rad_at_aperture = float(remove_last(self.beam_rad_at_aperture_value["text"], 3))
        except:
            pass
        try:
            I_not = laser_power / (math.pi * beam_rad_at_aperture**2 * 10E-6)
            self.I_not_value["text"] = f"{I_not} " + self.I_not_unit
        except:
            pass

        try:
            ZR = (math.pi * beam_rad_at_aperture**2 * 10E+6) / wavelength
            self.ZR_value["text"] = f"{ZR} " + self.ZR_unit
        except:
            pass

        try:
            close_aperture = float(remove_last(self.close_aperture_value["text"], 3))
        except:
            pass
        try:
            S = 1 - math.exp((-2 * close_aperture**2) / beam_rad_at_aperture**2)
            self.S_value["text"] = f"{S} " + self.S_unit
        except:
            pass

        try:
            VtoP = abs(float(self.del_phi_close_value["text"]))
        except:
            pass
        try:
            del_phi = VtoP / (0.406 * (1-S)**0.25)
            self.del_phi_value["text"] = f"{del_phi} " + self.del_phi_unit
        except:
            pass

        try:
            z_scan_sample_thickness = float(remove_last(self.z_scan_sample_thickness_value["text"], 3))
        except:
            pass
        try:
            leff = (1 - math.exp(-alpha * z_scan_sample_thickness)) / alpha
            self.leff_value["text"] = f"{leff} " + self.leff_unit
        except:
            pass

        try:
            k = (2 * math.pi) / wavelength
            self.k_value["text"] = f"{k} " + self.k_unit
        except:
            pass

        try:
            n2 = del_phi / (k * I_not * leff)
            self.n2_value["text"] = f"{n2} " + self.n2_unit
        except:
            pass

        try:
            NtoV = abs(float(self.del_phi_open_value["text"]))
        except:
            pass
        try:
            beta = (2 * math.sqrt(2) * NtoV) / (I_not * leff)
            self.beta_value["text"] = f"{beta} " + self.beta_unit
        except:
            pass

        try:
            n_not = float(self.linear_refractive_index_value["text"])
        except:
            pass
        try:
            real = (10E-4 * epsilon * (C**2) * (n_not**2) * n2) / math.pi
            self.real_value["text"] = f"{real} " + self.real_unit
        except:
            pass

        try:
            imaginary = (10E-2 * epsilon * (C**2) * (n_not**2) * wavelength * beta) / (4 * (math.pi**2))
            self.imaginary_value["text"] = f"{imaginary} " + self.imaginary_unit
        except:
            pass

        try:
            sus = math.sqrt((real**2) + (imaginary**2))
            self.susceptibility_value["text"] = f"{sus} " + self.susceptibility_unit
        except:
            pass


    def export_excel_data(self, gui):
        if get_file(self):
            file_path = filedialog.askdirectory(title="Select the folder to save")
            file_path = file_path + f"/{self.file_name_entry.get()}.xlsx"
            workbook, worksheet = gui.excel.initial_details(file_path)
            self.title = workbook.add_format({'border': 1, 'bold': True, 'align': 'center',
                                              'font_color': 'red', 'font_size': 14})
            self.border = workbook.add_format({'border': 1})
            self.bold_and_border = workbook.add_format({'border': 1, 'bold': True})
            self.bold = workbook.add_format({'bold': True})
            self.bold_and_border_and_center = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
            self.superscript_bold_border = workbook.add_format({'font_script': 1, 'bold': True, 'border': 1})
            self.subscript_bold = workbook.add_format({'font_script': 2, 'bold': True})
            self.make_as_text = workbook.add_format({'num_format': '@'})
            self.default = workbook.add_format()

            worksheet.merge_range("A1:G1", gui.title_value, self.title)
            worksheet.write("C2", gui.sweeping_distance_value, self.default)
            worksheet.write("C3", gui.step_distance_value, self.default)
            worksheet.write("C4", gui.bandwidth_value, self.default)
            worksheet.write("C5", self.wavelength_value, self.default)
            worksheet.write("C6", self.laser_beam_input_power_value, self.default)
            worksheet.write("C7", self.laser_beam_diameter_value, self.default)
            worksheet.write("C8", self.focal_length_value, self.default)
            worksheet.write("C9", self.close_aperture_value, self.default)
            worksheet.write("C10", self.beam_rad_at_aperture_value, self.default)
            worksheet.write("C11", self.z_scan_sample_thickness_value, self.default)
            worksheet.write("C12", "-", self.make_as_text)
            worksheet.write("C13", "-", self.make_as_text)
            worksheet.write("C14", "-", self.make_as_text)
            workbook.close()
