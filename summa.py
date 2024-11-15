# import datetime
#
# x = datetime.datetime.now()
#
# print(x.year)
# print(f"Z-Scan Sample analysis had done at SSN Research Center on {x.strftime("%d")}"
#                         f" {x.strftime("%b")} {x.strftime("%Y")} {x.strftime("%X")}")

from tkinter import ttk, messagebox, filedialog
from tkinter import *
root = Tk()
# file_path = filedialog.askdirectory(title="Select the folder to save")
#asksaveasfilename
# if file is None:
#     return

file_path = filedialog.SaveFileDialog(root, title="Select the folder to save")
root.mainloop()