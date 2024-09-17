import matplotlib.pyplot as plt
from tkinter import *
from tkinter import messagebox
from styles import Style
style = Style()


class Graph:
    def save_as_image(self, with_aperture_x_data, without_aperture_x_data,
                      with_aperture_y_data, without_aperture_y_data):
        self.with_aperture_x_data = with_aperture_x_data
        self.without_aperture_x_data = without_aperture_x_data
        self.with_aperture_y_data = with_aperture_y_data
        self.without_aperture_y_data = without_aperture_y_data
        if (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) == 0):
            self.without_aperture_plotter()
        elif (len(self.without_aperture_x_data) == 0) and (len(self.with_aperture_x_data) != 0):
            self.with_aperture_plotter()
        elif (len(self.without_aperture_x_data) != 0) and (len(self.with_aperture_x_data) != 0):
            self.plot_selector()

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
                                            bg=style.button_background,
                                            fg=style.button_foreground,
                                            bd=style.button_border_size,
                                            highlightbackground=style.button_highlight_background,
                                            highlightcolor=style.button_highlight_foreground,
                                            disabledforeground=style.button_disabled_foreground,
                                            activebackground=style.button_active_background,
                                            activeforeground=style.button_active_foreground,
                                            font=style.bold_font)
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


if __name__ == "__main__":
    Graph()
