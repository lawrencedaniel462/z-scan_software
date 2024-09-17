from tkinter import *


class RootGUI:
    def __init__(self, style):
        self.root = Tk()
        self.style = style
        self.root.title("Z Scan")
        self.root.geometry("1200x600")
        self.root.configure(background=style.bg_color)


if __name__ == "__main__":
    RootGUI()
