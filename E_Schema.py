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

	
	currentPage = 1
	currentPos = 0
	
	def __init__(self, rvtSys, _brdIndex, _sysIndex):
		self.brdIndex = _brdIndex
		self.sysIndex = _sysIndex
		self.location = None
		self.schType = self.__getType__()
		

	def __getType__ (self):
		global mainIsDisc
		global doc
		brdi = self.brdIndex
		sysi = self.sysIndex
		
		#for "Einspeisung" if no disconnector in board
		if all([brdi == 0, sysi == 0, mainIsDisc == 0]):
			schFamily = "E_SCH_Einspeisung-3P"
			schType = "Ausschalter"
		
		#for "Einspeisung" if disconnector in board
		elif all([brdi == 0, sysi == 0, mainIsDisc > 0]):
				schFamily = "E_SCH_Einspeisung-3P"
				schType = "Schutzschalter"
		
		else:
			schFamily = ""
			schType = ""
		
		tp = getTypeByCatFamType(
				BuiltInCategory.OST_GenericAnnotation,
				schFamily,
				schType)
		
		return tp
		
		# try:
			# return getByCatAndStrParam(
					# BuiltInCategory.OST_GenericAnnotation,
					# BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
					# schFamilyType, True)[0]
		# except:
			# #raise NameError ('No 2d diagramm found')
			# pass

	# def __getLocation__ (self):
		# #if == 0 it is "Einspeisung" 2 pos+indent
		# coordList = list()
		# coordList.append()
		# self.location = None

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

#diaList = 

# pages = int(math.ceil((len(allSystems) + len(
		# list(itertools.chain.from_iterable(
		# allSystems))))/8.0))+1

# pageNumLst = [brdName + "_" + str(n).zfill(3) for n in range(pages)]
# pageNameLst = [brdName] * pages

# #get TitleBlocks
# titleblatt = getByCatAndStrParam(
		# BuiltInCategory.OST_TitleBlocks,
		# BuiltInParameter.SYMBOL_NAME_PARAM,
		# "WSP_Plankopf_Shema_Titelblatt", True)[0]

# #get schemaPlankopf
# shemaPlankopf = getByCatAndStrParam(
		# BuiltInCategory.OST_TitleBlocks,
		# BuiltInParameter.SYMBOL_NAME_PARAM,
		# "WSP_Plankopf_Shema", True)[0]

# #========Find sheets
# existingSheets = [i for i in FilteredElementCollector(doc).
			# OfCategory(BuiltInCategory.OST_Sheets).
			# WhereElementIsNotElementType().
			# ToElements()
			# if i.LookupParameter("MC Panel Code"
			# ).AsString() == brdName]

# #=========Start transaction
# TransactionManager.Instance.EnsureInTransaction(doc)


# map(lambda x:doc.Delete(x.Id), existingSheets)

# #========Create sheets
# sheetLst = list()
# sheetLst.append(ViewSheet.Create(doc, titleblatt.Id))

# map(lambda x:sheetLst.append(ViewSheet.Create(doc, shemaPlankopf.Id)),
			# range(pages-1))
# map(lambda x:setPageParam(x), zip(sheetLst, pageNameLst, pageNumLst))


# #=========End transaction
# TransactionManager.Instance.TransactionTaskDone()

OUT = map(lambda x: x.schType, diaList)

