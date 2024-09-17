from tkinter import *

class Style:
    def __init__(self):
        self.bg_color = "#FFE5B4"

class RootGUI:
    def __init__(self, style):
        self.root = Tk()
        self.style = style
        self.root.title("Z Scan")
        self.root.geometry("1200x600")
        self.root.configure(background=style.bg_color)


style_instance = Style()
RootMaster = RootGUI(style_instance)