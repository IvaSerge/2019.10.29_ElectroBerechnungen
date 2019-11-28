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

def setPageParam(_lst):
	global doc
	global brdName
	page = _lst[0]
	pName = _lst[1]
	pNumber = _lst[2]
	page.get_Parameter(BuiltInParameter.SHEET_NAME).Set(pName)
	page.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(pNumber)
	page.LookupParameter("MC Panel Code").Set(brdName)

def getByCatAndStrParam (_bic, _bip, _val, _isType):
	global doc
	if _isType:
		fnrvStr = FilterStringEquals()
		pvp = ParameterValueProvider(ElementId(int(_bip)))
		frule = FilterStringRule(pvp, fnrvStr, _val, False)
		filter = ElementParameterFilter(frule)
		elem = FilteredElementCollector(doc).\
			OfCategory(_bic).\
			WhereElementIsElementType().\
			WherePasses(filter).\
			ToElements()
	else:
		fnrvStr = FilterStringEquals()
		pvp = ParameterValueProvider(ElementId(int(_bip)))
		frule = FilterStringRule(pvp, fnrvStr, _val, False)
		filter = ElementParameterFilter(frule)
		elem = FilteredElementCollector(doc).\
			OfCategory(_bic).\
			WhereElementIsNotElementType().\
			WherePasses(filter).\
			ToElements()
	
	return elem

def getTypeByCatFamType (_bic, _fam, _type):
	global doc
	fnrvStr = FilterStringEquals()
	
	pvpType = ParameterValueProvider(ElementId(int(BuiltInParameter.SYMBOL_NAME_PARAM)))
	pvpFam = ParameterValueProvider(ElementId(int(BuiltInParameter.ALL_MODEL_FAMILY_NAME)))
	
	fruleF = FilterStringRule(pvpFam, fnrvStr, _fam, False)
	filterF = ElementParameterFilter(fruleF)
	
	fruleT = FilterStringRule(pvpType, fnrvStr, _type, False)
	filterT = ElementParameterFilter(fruleT)
	
	filter = LogicalAndFilter(filterT, filterF)
	
	elem = FilteredElementCollector(doc).\
	OfCategory(_bic).\
	WhereElementIsElementType().\
	WherePasses(filter).\
	FirstElement()
	
	return elem

class dia():
	"""Diagramm class"""
	currentPage = 0
	currentPos = 0
	subBoardObj = None
	
	#coordinates of points on scheet
	coordList = list()
	coordList.append(XYZ(0.0738188976375485, 0.66929133858268, 0))
	coordList.append(XYZ(0.113188976377707, 0.66929133858268, 0))
	coordList.append(XYZ(0.204396325459072, 0.66929133858268, 0))
	coordList.append(XYZ(0.295603674540437, 0.66929133858268, 0))
	coordList.append(XYZ(0.386811023621802, 0.669291338582679, 0))
	coordList.append(XYZ(0.478018372703167, 0.669291338582679, 0))
	coordList.append(XYZ(0.569225721784533, 0.669291338582679, 0))
	coordList.append(XYZ(0.660433070865899, 0.669291338582678, 0))
	coordList.append(XYZ(0.751640419947264, 0.669291338582678, 0))
	coordList.append(XYZ(0.84284776902863, 0.669291338582678, 0))

	
	def __init__(self, _rvtSys, _brdIndex, _sysIndex):
		self.brdIndex = _brdIndex
		self.sysIndex = _sysIndex
		self.rvtSys = _rvtSys
		self.location = None
		self.cbType = _rvtSys.LookupParameter("MC CB Type").AsString()
		self.nPoles = _rvtSys.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES).AsInteger()
		self.pageN = None
		
		self.schType = self.__getType__()
		self.__getLocation__()

	def __getType__ (self):
		global mainIsDisc
		global doc
		brdi = self.brdIndex
		sysi = self.sysIndex
		
		#for "Einspeisung" if no disconnector in board
		if all([brdi == 0, sysi == 0, mainIsDisc == 0]):
			schFamily = "E_SCH_Einspeisung-3P"
			schType = "Schutzschalter"
		
		#for "Einspeisung" if disconnector in board
		elif all([brdi == 0, sysi == 0, mainIsDisc > 0]):
				schFamily = "E_SCH_Einspeisung-3P"
				schType = "Ausschalter"
		
		#mainBrd systems QF 1phase
		elif all([brdi == 0, sysi > 0, self.cbType == "QF", self.nPoles == 1]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-1P"
			schType = "Schutzschalter"
		
		#mainBrd systems QF 3phase
		elif all([brdi == 0, sysi > 0, self.cbType == "QF", self.nPoles == 3]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-3P"
			schType = "Schutzschalter"
		
		#2lvl main systems QF
		elif all([brdi > 0, sysi == 0, self.cbType == "QF"]):
			schFamily = "E_SCH_Einspeisung-3P_2lvl"
			schType = "Schutzschalter"
			dia.subBoardType = "QF"
		
		#2lvl main systems QF-FI
		elif all([brdi > 0, sysi == 0, self.cbType == "QF-FI"]):
			schFamily = "E_SCH_Einspeisung-3P_2lvl"
			schType = "QF-FI_Schalter"
			dia.subBoardType = "QF-FI"
		
		#2lvl systems QF 1phase QF in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF", 
				self.nPoles == 1, self.cbType == "QF"]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-1P"
			schType = "Schutzschalter_Zusätzliche"
		
		#2lvl systems QF 1phase QF-FI in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF-FI", 
						self.nPoles == 1, self.cbType == "QF"]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-1P"
			schType = "Schutzschalter_Zusätzliche_N"
		
		#2lvl systems QF 3phase QF in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF", 
				self.nPoles == 3, self.cbType == "QF"]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-3P"
			schType = "Schutzschalter_Zusätzliche"
		
		#2lvl systems QF 3phase QF in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF-FI", 
				self.nPoles == 3, self.cbType == "QF"]):
			schFamily = "E_SCH_SICHERUNGSSCHALTER-3P"
			schType = "Schutzschalter_Zusätzliche_N"
		
		#2lvl systems QF-FI phase QF in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF",
						self.nPoles == 1, self.cbType == "QF-FI"]):
			schFamily = "E_SCH_QF-FI-SCHALTER-1P"
			schType = "QF-FI_Zusätzliche"
		
		#2lvl systems QF-FI 1phase QF-FI in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF-FI",
						self.nPoles == 1, self.cbType == "QF-FI"]):
			schFamily = "E_SCH_QF-FI-SCHALTER-1P"
			schType = "QF-FI_Zusätzliche_N"

		#2lvl systems QF-FI 3phase QF in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF",
						self.nPoles == 3, self.cbType == "QF-FI"]):
			schFamily = "E_SCH_QF-FI-SCHALTER-3P"
			schType = "QF-FI_Zusätzliche"
		
		#2lvl systems QF-FI 3phase QF-FI in subboard
		elif all([brdi > 0, sysi > 0, dia.subBoardType == "QF-FI",
						self.nPoles == 3, self.cbType == "QF-FI"]):
			schFamily = "E_SCH_QF-FI-SCHALTER-3P"
			schType = "QF-FI_Zusätzliche_N"
		
		else:
			schFamily = ""
			schType = ""
		
		tp = getTypeByCatFamType(
				BuiltInCategory.OST_GenericAnnotation,
				schFamily,
				schType)
		
		return tp

	def __getLocation__ (self):
		modulSize = self.schType.LookupParameter("E_PositionsHeld").AsInteger()
		
		#Zerro modul 
		if modulSize == 0:
			dia.currentPage += 1
			dia.currentPos = 0
			self.location = dia.coordList[dia.currentPos]
			dia.currentPos += 3
		
		#next modules
		nextPos = dia.currentPos + modulSize
		if nextPos <= 8 and modulSize > 0: #enought
			self.location = dia.coordList[dia.currentPos]
			dia.currentPos = nextPos

		if nextPos > 8 and modulSize > 0:
			dia.currentPage += 1
			dia.currentPos = 1
			self.location = dia.coordList[dia.currentPos]
			dia.currentPos = modulSize
		
		#set page
		self.pageN = dia.currentPage

	def placeDiagramm (self):
		global doc
		global sheetLst
		self.diaInst = doc.Create.NewFamilyInstance(
					self.location, 
					self.schType,
					sheetLst[self.pageN])

brdName = IN[0]
reload = IN[1]

#get mainBrd by name
mainBrd = getByCatAndStrParam(
		BuiltInCategory.OST_ElectricalEquipment,
		BuiltInParameter.RBS_ELEC_PANEL_NAME,
		brdName, False)[0]
mainIsDisc = mainBrd.LookupParameter("E_IsDisconnector").AsInteger()

#get connectedBrds
connectedBrds = getByCatAndStrParam(
		BuiltInCategory.OST_ElectricalEquipment,
		BuiltInParameter.RBS_ELEC_PANEL_SUPPLY_FROM_PARAM,
		brdName, False)

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

diaList = list()
#========Initialaise dia class
for i, sysLst in enumerate(allSystems):
	for j, sys in enumerate (sysLst):
		diaList.append(dia(sys, i, j))

pages = max([x.pageN for x in diaList])

pageNumLst = [brdName + "_" + str(n).zfill(3) for n in range(pages+1)]
pageNameLst = [brdName] * (pages+1)

#get TitleBlocks
titleblatt = getByCatAndStrParam(
		BuiltInCategory.OST_TitleBlocks,
		BuiltInParameter.SYMBOL_NAME_PARAM,
		"WSP_Plankopf_Shema_Titelblatt", True)[0]

#get schemaPlankopf
shemaPlankopf = getByCatAndStrParam(
		BuiltInCategory.OST_TitleBlocks,
		BuiltInParameter.SYMBOL_NAME_PARAM,
		"WSP_Plankopf_Shema", True)[0]

#========Find sheets
existingSheets = [i for i in FilteredElementCollector(doc).
			OfCategory(BuiltInCategory.OST_Sheets).
			WhereElementIsNotElementType().
			ToElements()
			if i.LookupParameter("MC Panel Code"
			).AsString() == brdName]

#=========Start transaction
TransactionManager.Instance.EnsureInTransaction(doc)


# #map(lambda x:doc.Delete(x.Id), existingSheets)

# #========Create sheets
sheetLst = existingSheets

# sheetLst = list()
# sheetLst.append(ViewSheet.Create(doc, titleblatt.Id))

# map(lambda x:sheetLst.append(ViewSheet.Create(doc, shemaPlankopf.Id)),
			# range(pages))
# map(lambda x:setPageParam(x), zip(sheetLst, pageNameLst, pageNumLst))

map(lambda x: x.placeDiagramm(), diaList)

#=========End transaction
TransactionManager.Instance.TransactionTaskDone()

OUT = map(lambda x: [x.location, x.rvtSys, x.schType], diaList)
#OUT = sheetLst