import clr
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import System
from System import Array
from System.Collections.Generic import *

import itertools
import math

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

#region

def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	param = []
	
	for i in builtInParams:
		if i.ToString() == paramName:
			param.append(i)
			break
		else:
			continue
	return param[0]

def SetupParVal(elem, name, pValue):
	global doc
	try:
		elem.LookupParameter(name).Set(pValue)
	except:
		bip = GetBuiltInParam(name)
		elem.get_Parameter(bip).Set(pValue)
	return elem

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

def addFooter(_diaList):
	global sheetLst
	outlist = list()
	schFamily = "E_SCH_Footer"
	schType = ""
	pages = max([x.pageN for x in _diaList]) + 1
	
	for i in range(1, pages):
		onPage = [x for x in _diaList if x.pageN == i]
		brd = onPage[0].brdIndex
		brdSys = [x for x in _diaList
				if x.brdIndex == brd
				and x.sysIndex == 0][0]
		cbType = brdSys.cbType
		
		#=====getType=====
		#mainBrd systems QF 1phase
		if brd == 0 and brdSys.sysIndex == 0:
			schType = "Primärreifen"
		
		#2lvl systems QF
		if brd > 0 and brdSys.cbType == "QF":
			schType = "Zusätzliche"
		
		#2lvl systems FU+FI
		if brd > 0 and "FI" in brdSys.cbType:
			schType = "Zusätzliche_N"
		
		#=====getLocation=====
		lastIndex = max([dia.coordList.index((x.location))
							for x in _diaList
							if x.pageN == i])
		footIndex = lastIndex +1
		locPnt = dia.coordList[footIndex]
		
		tp = getTypeByCatFamType(
			BuiltInCategory.OST_GenericAnnotation,
			schFamily,
			schType)
		
		#=====create=====
		diaInst = doc.Create.NewFamilyInstance(
					locPnt, 
					tp,
					sheetLst[i])
		
		outlist.append(diaInst)
	return outlist

def addFiller(_diaList):
	global sheetLst
	outlist = list()
	schFamily = "E_SCH_Filler"
	schType = "Filler"
	tp = getTypeByCatFamType(
			BuiltInCategory.OST_GenericAnnotation,
			schFamily,
			schType)
	
	pages = max([x.pageN for x in _diaList]) + 1
	for i in range(1, pages):
		onPage = [x for x in _diaList if x.pageN == i]
		fillIndex = max([dia.coordList.index((x.location))
							for x in _diaList
							if x.pageN == i]) + 1
		while fillIndex <= 9:
			locPnt = dia.coordList[fillIndex]
			diaInst = doc.Create.NewFamilyInstance(
					locPnt, 
					tp,
					sheetLst[i])
			fillIndex += 1
			outlist.append(diaInst)
	return outlist

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
#endregion

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
		self.diaInst = None
		self.paramLst = list()
		
		try:
			self.schType = self.__getType__()
		except:
			raise ValueError("No 2D diagram found for {0.brdIndex}, {0.sysIndex}".format(self))
		
		self.__getLocation__()
		# self.__getParameters__()

	def __getType__ (self):
		global mainBrd
		global doc
		brdi = self.brdIndex
		sysi = self.sysIndex
		
		#for "Einspeisung" schema
		#diagramm is writen in board parameter
		if all([brdi == 0, sysi == 0]):
			schFamily = mainBrd.LookupParameter("E_Sch_Family").AsString()
			schType = mainBrd.LookupParameter("E_Sch_FamilyType").AsString()

		#for "Electrical" schema
		else:
			#diagramm is writen in electrical system parameter
			schFamily = self.rvtSys.LookupParameter("E_Sch_Family").AsString()
			schType = self.rvtSys.LookupParameter("E_Sch_FamilyType").AsString()
		
		tp = getTypeByCatFamType(
				BuiltInCategory.OST_GenericAnnotation,
				schFamily,
				schType)
		return tp

	def __getLocation__ (self):
		brdi = self.brdIndex
		sysi = self.sysIndex
		modulSize = self.schType.LookupParameter("E_PositionsHeld").AsInteger()
		
		#Start modul 
	 	if 	sysi == 0:
	 		dia.currentPage += 1
	 		dia.currentPos = 1
	 		self.location = dia.coordList[dia.currentPos]
		
	 	#next modules
	 	nextPos = dia.currentPos + modulSize
	 	if nextPos <= 9:  #enought space for next element
	 		self.location = dia.coordList[dia.currentPos]
			dia.currentPos = nextPos

	 	if nextPos > 9:
	 		dia.currentPage += 1
	 		dia.currentPos = 1
	 		self.location = dia.coordList[dia.currentPos]
	 		dia.currentPos = 1 + modulSize
		
	 	#set page
	 	self.pageN = dia.currentPage

	def placeDiagramm (self):
		global doc
		global sheetLst
		self.diaInst = doc.Create.NewFamilyInstance(
					self.location, 
					self.schType,
					sheetLst[self.pageN])

	# def __getParameters__ (self):
	# 	outlist = list()
	# 	#read info from board
	# 	if self.sysIndex == 0:
	# 		brd = [x for x in self.rvtSys.Elements][0]
	# 		frmSize = brd.LookupParameter("MC Frame Size").AsDouble()
	# 		isD = brd.LookupParameter("E_IsDisconnector").AsInteger()
	# 		if isD == 0:
	# 			cbType = "QF"
	# 		else:
	# 			cbType = "QS"

	# 	#read info from system
	# 	if self.sysIndex > 0:
	# 		frmSize = self.rvtSys.LookupParameter("MC Frame Size").AsDouble()
	# 		cbType = self.rvtSys.LookupParameter("MC CB Type").AsString()
		
	# 	cName = self.rvtSys.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NAME).AsString()
	# 	cab = self.rvtSys.LookupParameter("E_CableType").AsString()
		
	# 	outlist.append(["MC Frame Size", frmSize])
	# 	outlist.append(["E_CableType", cab])
	# 	outlist.append(["MC CB Type", cbType])
	# 	outlist.append(["RBS_ELEC_CIRCUIT_NAME", cName])
	# 	map(lambda x: self.paramLst.append(x), outlist)

	# def setParameters (self):
	# 	for i in self.paramLst:
	# 		elem = self.diaInst
	# 		pName = i[0]
	# 		pValue = i[1]
	# 		SetupParVal (elem, pName, pValue)

brdName = IN[0]
createNewScheets = IN[1]
reload = IN[2]

#get mainBrd by name
mainBrd = getByCatAndStrParam(
		BuiltInCategory.OST_ElectricalEquipment,
		BuiltInParameter.RBS_ELEC_PANEL_NAME,
		brdName, False)[0]

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

if lowSystems:
	lowSystemsId = list(itertools.chain.from_iterable(lowSystems))
	lowSystemsId = [i.Id for i in lowSystemsId]
	mainSystems = [i for i in getSystems(mainBrd)
 					if i.Id not in lowSystemsId]
else:
	try:
		mainSystems = [i for i in getSystems(mainBrd)]
	except:
		raise ValueError("Board have no electrical system")

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

# #========Find sheets
existingSheets = [i for i in FilteredElementCollector(doc).
			OfCategory(BuiltInCategory.OST_Sheets).
			WhereElementIsNotElementType().
			ToElements()
			if i.LookupParameter("MC Panel Code"
			).AsString() == brdName]

#=========Start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

# #========Create sheets========
sheetLst = list()
if createNewScheets == False:
	sheetLst = existingSheets
	elemsOnSheet = list()
	#remove all instances on sheet
	for sheet in sheetLst:
		elems = FilteredElementCollector(doc
				).OwnedByView(sheet.Id
				).OfCategory(BuiltInCategory.OST_GenericAnnotation
				).WhereElementIsNotElementType().ToElementIds()
		map(lambda x: elemsOnSheet.append(x), elems)
	typed_list = List[ElementId](elemsOnSheet)
	doc.Delete(typed_list)

if createNewScheets == True:
	map(lambda x:doc.Delete(x.Id), existingSheets)
	sheetLst.append(ViewSheet.Create(doc, titleblatt.Id))
	map(lambda x:sheetLst.append(ViewSheet.Create(
				doc, shemaPlankopf.Id)), range(pages))
	map(lambda x:setPageParam(x), zip(sheetLst, pageNameLst, pageNumLst))

map(lambda x: x.placeDiagramm(), diaList)
# footers = addFooter(diaList)
# fillers = addFiller(diaList)
# map(lambda x: x.setParameters(), diaList)

# #=========End transaction
# TransactionManager.Instance.TransactionTaskDone()


#OUT = map(lambda x: [dia.coordList.index(x.location), x.pageN], diaList)
#OUT = map(lambda x: x.paramLst, diaList)
OUT = pageNameLst
#OUT = mainBrd.LookupParameter("E_Sch_Family").AsString()