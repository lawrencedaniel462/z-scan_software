import xlsxwriter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import re
import xml.etree.ElementTree as ET
tree = ET.parse("data_file.xml")
xml = tree.getroot()

sweeping_distance = 50
step_distance = 20

root = Tk()
root.geometry("200x100")
without_aperture_x_data = [0,1,2,3,4,5,6,7,8,9]
# with_aperture_x_data = [0,1,2,3,4,5,6,7,8,9]
with_aperture_x_data = []
without_aperture_y_data = [0,2,4,6,8,10,12,14,16,18]
# with_aperture_y_data = [0,2,4,6,8,10,12,14,16,18]
with_aperture_y_data = []
# y = []


def get_file():
    # regex = re.compile('.')
    # if (regex.search(file_name_entry.get())==None):
    #     return True
    if len(file_name_entry.get()) == 0:
        messagebox.showerror("Naming Error", "File name should not be empty")
        return False
    # elif:
    #     messagebox.showerror("Naming Error", "Special Characters are not allowed")
    #     return False
    else:
        return True


def excel_selector():
    save_data_button["state"] = "disabled"
    excel_window = Toplevel()
    excel_window.title("Plot Selector")
    excel_window.geometry("250x120")
    excel_without_aperture_var = IntVar()
    Checkbutton(excel_window, text="Plot without aperture", variable=excel_without_aperture_var).grid(row=0, column=0, pady=5, sticky=W)
    excel_with_aperture_var = IntVar()
    Checkbutton(excel_window, text="Plot with aperture", variable=excel_with_aperture_var).grid(row=1, column=0, pady=5, sticky=W)
    export_excel = Button(excel_window, text="Prepare Data", command=excel_mode_selector)
    export_excel.grid(row=2, column=0, pady=5)

def excel_mode_selector():
    save_data_button["state"] = "active"
    if excel_without_aperture_var.get() == 0 and self.with_aperture_var.get() == 0:
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

def without_aperture_excel():
    if get_file() == True:
        file_path = filedialog.askdirectory(title="Select the folder to save")
        file_path = file_path + f"/{file_name_entry.get()} open aperture.xlsx"
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

        worksheet.write("C1", f"{xml[0][0].text} nm", border)
        worksheet.write("C2", xml[0][1].text, border)
        worksheet.write("C3", f"{xml[1][6].text} mm", border)
        worksheet.write("C4", f"{xml[1][7].text} mm", border)
        worksheet.write("C5", sweeping_distance, border)
        worksheet.write("C6", step_distance, border)

        worksheet.merge_range("A8:B8", "Open Aperture", bold_and_border_and_center)
        worksheet.write("A9", "Distance in mm", bold_and_border)
        worksheet.write("B9", "Power in W", bold_and_border)
        row = 9
        column = 0

        for item in without_aperture_x_data:
            worksheet.write(row, column, item, border)
            row += 1

        row = 9
        column = 1

        for item in without_aperture_y_data:
            worksheet.write(row, column, item, border)
            row += 1

        workbook.close()
        print(file_path)

def with_aperture_excel():
    plt.plot(self.with_aperture_x_data, self.with_aperture_y_data, color="tab:red")
    plt.xlabel("Distance in mm")
    plt.ylabel("Power in W")
    plt.show()

def save_data_excel():
    if (len(without_aperture_x_data) != 0) and (len(with_aperture_x_data) == 0):
        without_aperture_excel()
    elif (len(without_aperture_x_data) == 0) and (len(with_aperture_x_data) != 0):
        with_aperture_excel()
    elif (len(without_aperture_x_data) != 0) and (len(with_aperture_x_data) != 0):
        excel_selector()



file_name_entry = Entry(root)
save_data_button = Button(root, text="Save", command=save_data_excel)

file_name_entry.grid(row=0, column=0)
save_data_button.grid(row=1, column=0)

root.mainloop()