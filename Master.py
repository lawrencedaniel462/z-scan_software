from GUI_Master import CommGUI
from Root_gui import RootGUI
from pyvisa import ResourceManager
from styles import Style
from commands import Commands
from excel import Excel
from graph import Graph
from graph_view_processor import graph_processor
from calculations import Calculation


style_master = Style()
RootMaster = RootGUI(style_master)
MyUSB = ResourceManager('@py')
CommandMaster = Commands()
ExcelMaster = Excel()
GraphMaster = Graph()
ViewMaster = graph_processor()
CalculationMaster = Calculation()


CommMaster = CommGUI(RootMaster.root, MyUSB, style_master, CommandMaster, ExcelMaster, GraphMaster, ViewMaster, CalculationMaster)

RootMaster.root.mainloop()
