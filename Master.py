from GUI_Master import RootGUI, CommGUI
from pyvisa import ResourceManager


RootMaster = RootGUI()
MyUSB = ResourceManager('@py')


CommMaster = CommGUI(RootMaster.root, MyUSB)

RootMaster.root.mainloop()
