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

#region "global functions"

def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	param = []
	for i in builtInParams:
		if i.ToString() == paramName:
			param.append(i)
			return i

def getParVal(elem, name):
	value = None
	#Параметр пользовательский
	param = elem.LookupParameter(name)
	#параметр не найден. Надо проверить, есть ли такой же встроенный параметр
	if param == None:
		param = elem.get_Parameter(GetBuiltInParam(name))

	#Если параметр найден, считываем значение
	try:
		storeType = param.StorageType
		#value = storeType
		if storeType == StorageType.String:
			value = param.AsString()
		elif storeType == StorageType.Integer:
			value  = param.AsDouble()
		elif storeType == StorageType.Double:
			value = param.AsDouble()
		elif storeType == StorageType.ElementId:
			value = param.AsValueString()
	except:
		pass
	return value

def setParVal(elem, name, pValue):
	global doc
	#Параметр пользовательский
	param = elem.LookupParameter(name)
	#параметр не найден. Надо проверить, есть ли такой же встроенный параметр
	if param == None:
		param = elem.get_Parameter(GetBuiltInParam(name))
	if param:
		param.Set(pValue)
	return elem

def getSystems (_brd):
	allsys = _brd.MEPModel.ElectricalSystems
	lowsys = _brd.MEPModel.AssignedElectricalSystems
	if lowsys:
		outlist = list()
		lowsysId = [i.Id for i in lowsys]
		mainboardsysLst = [i for i in allsys if i.Id not in lowsysId]
		if len(mainboardsysLst) == 0:
			raise ValueError("Board not connectet")
		else:
			mainboardsys = mainboardsysLst[0]
		
		lowsys = [i for i in allsys if i.Id in lowsysId]
		lowsys.sort(key = lambda x: float(getParVal
(x, "RBS_ELEC_CIRCUIT_NUMBER")))
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

def create_dia_by_board_Name(_brd_name, _brdLevel=0):
	"""
	 Creates list of diagramm objects for _brd_name

	 The function is recursive.
	 If to the board other subboards is connected, then
	 function would be called for subboard.
	 :in:
	 	_brd_name - str Board name
	 	_brdLevel - current level of the board. 
		 			_brdLevel = 0 - main board
					_brdLevel = n, n > 0 - level of current subboard.
					_brdLevel can be increased only with step += 1

	 :return:
		every new diagramm object would be appended to
		the flat list.
	"""
	global doc
	outlist = list()

	#get board by name
	brd_instance = getByCatAndStrParam(
		BuiltInCategory.OST_ElectricalEquipment,
		BuiltInParameter.RBS_ELEC_PANEL_NAME,
		_brd_name, False)[0]
	
	brd_bic = BuiltInCategory.OST_ElectricalEquipment
	brd_cat = Autodesk.Revit.DB.Category.GetCategory(
						doc, ElementId(brd_bic)).Id
	
	brd_circuits = [i for i in getSystems(brd_instance)]
	
	return brd_instance.Category.Id == brd_cat
	# endregion 	

class dia():
	"""Diagramm class"""

#region "class varaibles"

	currentPage = 0
	currentPos = 0
	
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
	coordList.append(XYZ(0.0738188976375485, 0.66929133858268, 0))

	parToSet = list()
	parToSet.append("RBS_ELEC_CIRCUIT_NAME")
	parToSet.append("RBS_ELEC_CIRCUIT_NUMBER")
	parToSet.append("RBS_ELEC_CIRCUIT_WIRE_TYPE_PARAM")
	parToSet.append("CBT:CIR_Kabel")
	parToSet.append("CBT:CIR_Nennstrom")
	parToSet.append("CBT:CIR_Schutztyp")
	parToSet.append("CBT:CIR_Elektrischen Schlag")
	parToSet.append("E_Stromkreisprefix")
	#endgerion

	def __init__ (self, _rvtSys, _brdIndex, _sysIndex, _brdLevel):
		self.brdLevel = _sysIndex
		self.brdIndex = _brdIndex
		self.sysIndex = _sysIndex
		self.rvtSys = _rvtSys
		self.location = None
		self.pageN = None
		self.diaInst = None
		self.paramLst = list()
		
		try:
			self.schType = self.__getType()
		except:
			raise ValueError("No 2D diagram found for {0.brdIndex}, {0.sysIndex}".format(self))
		
		self.__getLocation()

	def __getType (self):
		global mainBrd
		global doc
		global diaList
		brdi = self.brdIndex
		sysi = self.sysIndex
		schFamily = None
		schType = None

		#for "Einspeisung" schema
		#diagramm is writen in board parameter
		if all([brdi == 0, sysi == 0]):
			schFamily = mainBrd.LookupParameter("E_Sch_Family").AsString()
			schType = mainBrd.LookupParameter("E_Sch_FamilyType").AsString()

		#for header
		elif not(self.rvtSys) and sysi == 10:
			schFamily = "E_SCH_Filler"
			schType = "Filler_Start"
		
		#for footer
		elif not(self.rvtSys) and sysi == 11:
			#find board schema type
			boardDia = [x.schType for x in diaList
						if x.brdIndex == brdi][0]
			schFamily = boardDia.LookupParameter("E_Sch_Family").AsString()
			schType = boardDia.LookupParameter("E_Sch_FamilyType").AsString()

		#for filler
		elif not(self.rvtSys) and sysi < 10:
			#find board schema type
			schFamily = "E_SCH_Filler"
			schType = "Filler_1modul"	

		#for "Electrical" schema
		elif self.rvtSys:
			#diagramm is writen in electrical system parameter
			schFamily = self.rvtSys.LookupParameter("E_Sch_Family").AsString()
			schType = self.rvtSys.LookupParameter("E_Sch_FamilyType").AsString()
		
		else: pass
		tp = getTypeByCatFamType(
				BuiltInCategory.OST_GenericAnnotation,
				schFamily,
				schType)
		return tp

	def __getLocation (self):
		global dialist
		brdi = self.brdIndex
		sysi = self.sysIndex
		if self.schType == None:
			raise ValueError("No 2D diagram was found for {},{}".format(brdi, sysi))
		
		try:
			modulSize = self.schType.LookupParameter("E_PositionsHeld").AsInteger()
		except:
			raise ValueError("No Type Parameter \"E_PositionsHeld\" in Family {0.schType.Family.Name}".format(self))
		nextPos = dia.currentPos + modulSize
		
		#Start modul 
		if 	sysi == 0 and brdi == 0:
			#if it is not the first board - create page break
			if brdi == 0:
				dia.currentPos = 1
				dia.currentPage += 1
				self.location = dia.coordList[dia.currentPos]
			self.pageN = dia.currentPage
			dia.currentPos += modulSize

		#Header
		elif sysi == 10 and not(self.rvtSys):
			self.location = dia.coordList[0]
			#brdIndex == is equal page number
			self.pageN = self.brdIndex		
		
		#Footer
		elif sysi == 11 and not(self.rvtSys):
			lastDia = [x.schType for x in diaList
						if x.brdIndex == brdi][-1]
			previousModulSize = lastDia.LookupParameter("E_PositionsHeld").AsInteger()
			lastIndex = [dia.coordList.index((x.location)) for x in diaList
						if x.brdIndex == brdi][-1]
			footIndex = lastIndex + previousModulSize
			
			#enought space for Footer
			if footIndex <= 10:
				self.location = dia.coordList[footIndex]
				self.pageN = max([x.pageN for x in diaList
								if x.brdIndex == brdi])
				dia.currentPos += modulSize
			
			else:
			#not enought space for next Footer
			#no need to created footer
				self.location = None
				self.pageN = None

		#Filler
		elif not(self.rvtSys) and sysi < 10:
			self.location = dia.coordList[sysi]
			#brdIndex == is equal page number
			self.pageN = self.brdIndex	

		#next modules
		#enought space for next element
		elif nextPos <= 9:
			self.location = dia.coordList[dia.currentPos]
			dia.currentPos = nextPos
			self.pageN = dia.currentPage
		
		#next modules
		#not enought space for next element
		elif nextPos > 9:
			dia.currentPage += 1
			dia.currentPos = 1
			self.location = dia.coordList[dia.currentPos]
			dia.currentPos = 1 + modulSize
			self.pageN = dia.currentPage
		else: pass

	def placeDiagramm (self):
		global doc
		global sheetLst
		self.diaInst = doc.Create.NewFamilyInstance(
					self.location, 
					self.schType,
					sheetLst[self.pageN])

	def getParameters (self):
		#для вводного щита
		#для электрической системы
		self.paramLst = [[x, getParVal(self.rvtSys, x)]
						for x in dia.parToSet]

	def setParameters (self):
		for i,j in self.paramLst:
			elem = self.diaInst
			if j == None:
				j = " "
			setParVal (elem, i, j)

brdName = IN[0]
createNewScheets = IN[1]
reload = IN[2]

outlist = list()
footers = list()
outlist = list()

diaList = create_dia_by_board_Name(brdName)

# #create tree of electircal systems and initiate dia class
# for i, sys in enumerate(mainSystems):
# 	brdCategory = mainBrd.Category.Id
# 	#Is this system contains subsystems "Sammelschiene"?
# 	lowbrd = None
# 	elems = [elem for elem in sys.Elements]
# 	elem = elems[0]
# 	#is it electrical board?
# 	if elem.Category.Id == brdCategory:
# 		brdCode = elem.LookupParameter("MC Panel Code").AsString()
# 		#is this board marked as subboard?
# 		if brdCode == brdName:
# 			lowbrd = elem

# 	if i == 0:
# 	#System 0,0
# 		diaList.append(dia(sys, 0, 0))
# 		outlist.append(str(0)+","+str(i))

# 	#LowBoards systems
# 	elif lowbrd:
# 		lowSystems = getSystems(lowbrd)
# 		for j, lowSys in enumerate(lowSystems):
# 			outlist.append(str(i)+","+str(j))
# 			diaList.append(dia(lowSys, i, j))
# 			#If it is the last system - create footer.
# 			#if j == len(lowSystems) - 1:
# 				#newFooter = dia(None, i, 11)
# 				#check if we need footer
# 				#if newFooter.location:
# 				#footers.append(newFooter)

# 	#Systems without boards
# 	else:
# 		outlist.append(str(0)+","+str(i))
# 		diaList.append(dia(sys, 0, i))
# 		#If it is the last system - create footer.
# 		#if i == len(mainSystems) - 1:
# 				#newFooter = dia(None, 0, 11)
# 				#footers.append(newFooter)

# #========Prepare model parameters========
# map(lambda x: x.getParameters(), diaList)


#region "pages properties"

# pages = max([x.pageN for x in diaList])
# pageNumLst = [brdName + "_" + str(n).zfill(3) for n in range(pages+1)]
# pageNameLst = [brdName] * (pages+1)

# #========Initialaise dia class for Footers Headers and Fillers
# headers = [dia(None, x, 10) for x in range(1, pages + 1)]

# #fillers for pages
# fillers = list()
# for page in range(1, pages+1):
# 	lastPageIndex = [dia.coordList.index((x.location)) for x in diaList
# 						if x.pageN == page][-1]
# 	fillersOnPage = [dia(None, page, x) 
# 					for x in range(lastPageIndex + 1, 10)]
# 	map(lambda x: fillers.append(x), fillersOnPage)

# #get TitleBlocks
# titleblatt = getByCatAndStrParam(
# 		BuiltInCategory.OST_TitleBlocks,
# 		BuiltInParameter.SYMBOL_NAME_PARAM,
# 		"WSP_Plankopf_Shema_Titelblatt", True)[0]

# #get schemaPlankopf
# shemaPlankopf = getByCatAndStrParam(
# 		BuiltInCategory.OST_TitleBlocks,
# 		BuiltInParameter.SYMBOL_NAME_PARAM,
# 		"WSP_Plankopf_Shema", True)[0]

# #========Find sheets
# existingSheets = [i for i in FilteredElementCollector(doc).
# 			OfCategory(BuiltInCategory.OST_Sheets).
# 			WhereElementIsNotElementType().
# 			ToElements()
# 			if i.LookupParameter("MC Panel Code"
# 			).AsString() == brdName]

#endregion 

#=========Start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

# #========Create sheets========
# sheetLst = list()
# if createNewScheets == False:
# 	sheetLst = existingSheets
# 	elemsOnSheet = list()
# 	#remove all instances on sheet
# 	for sheet in sheetLst:
# 		elems = FilteredElementCollector(doc
# 				).OwnedByView(sheet.Id
# 				).OfCategory(BuiltInCategory.OST_GenericAnnotation
# 				).WhereElementIsNotElementType().ToElementIds()
# 		map(lambda x: elemsOnSheet.append(x), elems)
# 	typed_list = List[ElementId](elemsOnSheet)
# 	doc.Delete(typed_list)

# if createNewScheets == True:
# 	map(lambda x:doc.Delete(x.Id), existingSheets)
# 	sheetLst.append(ViewSheet.Create(doc, titleblatt.Id))
# 	map(lambda x:sheetLst.append(ViewSheet.Create(
# 				doc, shemaPlankopf.Id)), range(pages))
# 	map(lambda x:setPageParam(x), zip(sheetLst, pageNameLst, pageNumLst))

# #========Place diagramms========
# map(lambda x: x.placeDiagramm(), diaList)
# # map(lambda x: x.placeDiagramm(), footers)
# map(lambda x: x.placeDiagramm(), headers)
# map(lambda x: x.placeDiagramm(), fillers)

# #========Set Parameters========
# map(lambda x: x.setParameters(), diaList)

#=========End transaction
TransactionManager.Instance.TransactionTaskDone()

# OUT = map(lambda x: ["{},{}".format(x.brdIndex, x.sysIndex), x.rvtSys, x.schType, dia.coordList.index((x.location)), x.pageN], diaList)
OUT = diaList
