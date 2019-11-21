import clr
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import System
from System import Array

import itertools
import math

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

def getSystems (_brd):
	allsys = _brd.MEPModel.ElectricalSystems
	lowsys = _brd.MEPModel.AssignedElectricalSystems
	if lowsys:
		outlist = list()
		lowsysId = [i.Id for i in lowsys]
		mainboardsys = [i for i in allsys if i.Id not in lowsysId][0]
		lowsys = [i for i in allsys if i.Id in lowsysId]
		outlist.append(mainboardsys)
		map(lambda x: outlist.append(x), lowsys)
		return outlist
	else:
		return list(allsys)

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
map(lambda x: lowbrds.append(x), connectedBrds)

lowbrds = [i for i in lowbrds if 
			i.LookupParameter(
			"MC Panel Code").AsString() == brdName]

#get systems
lowSystems = map(lambda x: getSystems(x), lowbrds)
lowSystemsId = list(itertools.chain.from_iterable(lowSystems))
lowSystemsId = [i.Id for i in lowSystemsId]

mainSystems = [i for i in getSystems(mainBrd)
				if i.Id not in lowSystemsId]

allSystems = list()
allSystems.append(mainSystems)
map(lambda x: allSystems.append(x), lowSystems)

pages = math.ceil((len(allSystems) + len(
		list(itertools.chain.from_iterable(
		allSystems))))/8.0)

#get TitleBlocks
testParam = BuiltInParameter.SYMBOL_NAME_PARAM
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringEquals()
frule = FilterStringRule(pvp, fnrvStr, "WSP_Plankopf_Shema_Titelblatt", False)
filter = ElementParameterFilter(frule)
titleblatt = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_TitleBlocks).\
	WhereElementIsElementType().\
	WherePasses(filter).\
	FirstElement()

testParam = BuiltInParameter.SYMBOL_NAME_PARAM
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringEquals()
frule = FilterStringRule(pvp, fnrvStr, "WSP_Plankopf_Shema", False)
filter = ElementParameterFilter(frule)
shemaPlankopf = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_TitleBlocks).\
	WhereElementIsElementType().\
	WherePasses(filter).\
	FirstElement()

TransactionManager.Instance.EnsureInTransaction(doc)

newSheet = ViewSheet.Create(doc, titleblatt.Id)

TransactionManager.Instance.TransactionTaskDone()

#======================
OUT = newSheet
#======================
