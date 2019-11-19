import clr
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import System
from System import Array

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

def getSystems (_brd):
	allsys = _brd.MEPModel.ElectricalSystems
	lowsys = _brd.MEPModel.AssignedElectricalSystems
	if lowsys:
		lowsysId = [i.Id for i in lowsys]
		mainboard = [i for i in allsys if i.Id not in lowsysId][0]
		lowsys = [i for i in allsys if i.Id in lowsysId]
		return mainboard, lowsys
	else:
		return None

brdName = IN[0]
reload = IN[1]

#get mainBrd by name
testParam = BuiltInParameter.RBS_ELEC_PANEL_NAME
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringEquals()
frule = FilterStringRule(pvp, fnrvStr, brdName, False)
filter = ElementParameterFilter(frule)
mainBrd = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_ElectricalEquipment).\
	WhereElementIsNotElementType().\
	WherePasses(filter).\
	FirstElement()

#get connectedBrds
testParam = BuiltInParameter.RBS_ELEC_PANEL_SUPPLY_FROM_PARAM
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringEquals()
frule = FilterStringRule(pvp, fnrvStr, brdName, False)
filter = ElementParameterFilter(frule)
connectedBrds = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_ElectricalEquipment).\
	WhereElementIsNotElementType().\
	WherePasses(filter).\
	ToElements()
connectedBrds = sorted(connectedBrds, key=lambda brd:brd.Name)

lowbrds = list()
lowbrds.append(mainBrd)
map(lambda x: lowbrds.append(x), connectedBrds)

lowbrds = [i for i in lowbrds if 
			i.LookupParameter(
			"MC Panel Code").AsString() == brdName]

brdSystems = map(lambda x: getSystems(x), lowbrds)

#======================
OUT = brdSystems
#======================
