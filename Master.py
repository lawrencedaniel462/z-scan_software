from GUI_Master import CommGUI
from Root_gui import RootGUI
from pyvisa import ResourceManager
from styles import Style


style_master = Style()
RootMaster = RootGUI(style_master)
MyUSB = ResourceManager('@py')


CommMaster = CommGUI(RootMaster.root, MyUSB, style_master)

RootMaster.root.mainloop()
