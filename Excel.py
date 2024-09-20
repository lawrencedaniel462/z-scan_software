from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import xlsxwriter
import datetime
from styles import Style
style = Style()


class Excel:
    def save_data_excel(self, data):
        self.without_aperture_excel(data)
        # if (len(data.without_aperture_x_data) != 0) and (len(data.with_aperture_x_data) == 0):
        #     self.without_aperture_excel(data)
        # elif (len(data.without_aperture_x_data) == 0) and (len(data.with_aperture_x_data) != 0):
        #     self.with_aperture_excel(data)
        # elif (len(data.without_aperture_x_data) != 0) and (len(data.with_aperture_x_data) != 0):
        #     self.excel_selector(data)

    def get_file(self, data):
        # regex = re.compile('.')
        # if (regex.search(file_name_entry.get())==None):
        #     return True
        if len(data.enter_file_name_entry.get()) == 0:
            messagebox.showerror("Naming Error", "File name should not be empty")
            return False
        else:
            return True

    def initial_details(self, file_path):
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()
        self.title = workbook.add_format({'border': 1, 'bold': True, 'align': 'center',
                                          'font_color': 'red', 'font_size': 14})
        self.border = workbook.add_format({'border': 1})
        print(type(self.border))
        self.bold_and_border = workbook.add_format({'border': 1, 'bold': True})
        self.bold = workbook.add_format({'bold': True})
        self.bold_and_border_and_center = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
        self.superscript_bold_border = workbook.add_format({'font_script': 1, 'bold': True, 'border': 1})
        self.subscript_bold = workbook.add_format({'font_script': 2, 'bold': True})
        self.make_as_text = workbook.add_format({'num_format': '@'})
        self.default = workbook.add_format()
        worksheet.set_column_pixels(0, 7, 120)

        worksheet.merge_range("A2:B2", "Sweeping Distance", self.bold)
        worksheet.merge_range("A3:B3", "Step Distance", self.bold)
        worksheet.merge_range("A4:B4", "Bandwidth", self.bold)
        worksheet.merge_range("A5:B5", "Wavelength (λ)", self.bold)
        worksheet.merge_range("A6:B6", "")
        worksheet.write_rich_string("A6",
                                    self.bold, "Power of the Laser (E",
                                    self.subscript_bold, "P",
                                    self.bold, ")")
        worksheet.merge_range("A7:B7", "Diameter of the Laser beam (d)", self.bold)
        worksheet.merge_range("A8:B8", "Focal Length of the Lens (f)", self.bold)
        worksheet.merge_range("A9:B9", "")
        worksheet.write_rich_string("A9",
                                    self.bold, "Radius of the Aperture (r",
                                    self.subscript_bold, "a",
                                    self.bold, ")")
        worksheet.merge_range("A10:B10", "")
        worksheet.write_rich_string("A10",
                                    self.bold, "Radius of the beam at Aperture (w",
                                    self.subscript_bold, "a",
                                    self.bold, ")")
        worksheet.merge_range("A11:B11", "Z-Scan Sample Thickness (L)", self.bold)
        worksheet.merge_range("A12:B12", "")
        worksheet.write_rich_string("A12",
                                    self.bold, "Linear refractive index (n",
                                    self.subscript_bold, "o",
                                    self.bold, ")")
        worksheet.merge_range("A13:B13", "Transmittance (T)", self.bold)
        worksheet.merge_range("A14:B14", "Thickness (t)", self.bold)
        return workbook, worksheet


    def fill_values(self, workbook, worksheet, data):
        x = datetime.datetime.now()
        title_string = (f"Z-Scan Sample analysis had done at SSN Research Center on {x.strftime("%d")}"
                        f" {x.strftime("%b")} {x.strftime("%Y")} {x.strftime("%X")}")
        worksheet.merge_range("A1:G1", title_string, self.title)

        worksheet.write("C2", data.sweeping_distance_entry.get() + " " + data.clicked_sweeping_distance.get(),
                        self.default)
        worksheet.write("C3", data.step_value_entry.get() + " " + data.clicked_step_value.get(), self.default)
        worksheet.write("C4", data.xml[0][1].text, self.default)
        worksheet.write("C5", f"{data.xml[1][6].text} nm", self.default)
        worksheet.write("C6", f"{data.xml[0][3].text} mW", self.default)
        worksheet.write("C7", f"{data.xml[0][4].text} mm", self.default)
        worksheet.write("C8", f"{data.xml[1][6].text} mm", self.default)
        worksheet.write("C9", f"{data.xml[1][7].text} mm", self.default)
        worksheet.write("C10", f"{data.xml[0][7].text} mm", self.default)
        worksheet.write("C11", f"{data.xml[0][8].text} mm", self.default)
        worksheet.write("C12", "-", self.make_as_text)
        worksheet.write("C13", "-", self.make_as_text)
        worksheet.write("C14", "-", self.make_as_text)

        return workbook, worksheet

    def without_aperture_excel(self, data):
        if self.get_file(data):
            file_path = filedialog.askdirectory(title="Select the folder to save")
            file_path = file_path + f"/{data.enter_file_name_entry.get()} open aperture.xlsx"
            workbook, worksheet = self.initial_details(file_path)
            workbook, worksheet = self.fill_values(workbook, worksheet, data)
            worksheet.merge_range("A17:B17", "Open Aperture", self.bold_and_border_and_center)
            worksheet.write("A18", "Distance in mm", self.bold_and_border_and_center)
            worksheet.write("B18", "Power in W", self.bold_and_border_and_center)
            row = 18
            column = 0

            for item in data.without_aperture_x_data:
                worksheet.write(row, column, item, self.border)
                row += 1

            row = 18
            column = 1

            for item in data.without_aperture_y_data:
                worksheet.write(row, column, item, self.border)
                row += 1

            workbook.close()

    def with_aperture_excel(self, data):
        if self.get_file(data):
            file_path = filedialog.askdirectory(title="Select the folder to save")
            file_path = file_path + f"/{data.enter_file_name_entry.get()} close aperture.xlsx"
            workbook, worksheet = self.initial_details(file_path)
            workbook, worksheet = self.fill_values(workbook, worksheet, data)
            worksheet.merge_range("A17:B17", "Close Aperture", self.bold_and_border_and_center)
            worksheet.write("A18", "Distance in mm", self.bold_and_border_and_center)
            worksheet.write("B18", "Power in W", self.bold_and_border_and_center)
            row = 18
            column = 0

            for item in data.with_aperture_x_data:
                worksheet.write(row, column, item, self.border)
                row += 1

            row = 18
            column = 1

            for item in data.with_aperture_y_data:
                worksheet.write(row, column, item, self.border)
                row += 1

            workbook.close()

    def excel_selector(self, data):
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
                                   command=lambda: self.excel_mode_selector(data),
                                   bg=style.button_background,
                                   fg=style.button_foreground,
                                   bd=style.button_border_size,
                                   highlightbackground=style.button_highlight_background,
                                   highlightcolor=style.button_highlight_foreground,
                                   disabledforeground=style.button_disabled_foreground,
                                   activebackground=style.button_active_background,
                                   activeforeground=style.button_active_foreground,
                                   font=style.bold_font)
        self.export_excel.grid(row=2, column=0, pady=5)

    def excel_mode_selector(self, data):
        if self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 0:
            messagebox.showerror("Selection Error", "Select at least one box")
        elif self.excel_without_aperture_var.get() == 1 and self.excel_with_aperture_var.get() == 0:
            self.without_aperture_excel(data)
        elif self.excel_without_aperture_var.get() == 0 and self.excel_with_aperture_var.get() == 1:
            self.with_aperture_excel(data)
        elif self.excel_without_aperture_var.get() == 1 and self.excel_with_aperture_var.get() == 1:
            if self.get_file(data):
                file_path = filedialog.askdirectory(title="Select the folder to save")
                file_path = file_path + f"/{data.enter_file_name_entry.get()}.xlsx"
                workbook, worksheet = self.initial_details(file_path)
                workbook, worksheet = self.fill_values(workbook, worksheet, data)

                worksheet.merge_range("A17:B17", "Open Aperture", self.bold_and_border_and_center)
                worksheet.write("A18", "Distance in mm", self.bold_and_border_and_center)
                worksheet.write("B18", "Power in W", self.bold_and_border_and_center)

                worksheet.merge_range("C17:D17", "Close Aperture", self.bold_and_border_and_center)
                worksheet.write("C18", "Distance in mm", self.bold_and_border_and_center)
                worksheet.write("D18", "Power in W", self.bold_and_border_and_center)

                row = 18
                column = 0

                for item in data.without_aperture_x_data:
                    worksheet.write(row, column, item, self.border)
                    row += 1

                row = 18
                column = 1

                for item in data.without_aperture_y_data:
                    worksheet.write(row, column, item, self.border)
                    row += 1

                row = 18
                column = 2

                for item in data.with_aperture_x_data:
                    worksheet.write(row, column, item, self.border)
                    row += 1

                row = 18
                column = 3

                for item in data.with_aperture_y_data:
                    worksheet.write(row, column, item, self.border)
                    row += 1

                workbook.close()
            self.excel_window.destroy()


if __name__ == "__main__":
    Excel()
