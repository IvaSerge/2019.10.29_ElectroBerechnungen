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

import math
from math import sqrt

import re
from re import *

def GetParVal(elem, name):
	#Параметр пользовательский
	try:
		param = elem.LookupParameter(name)
		storeType = param.StorageType
		if storeType == StorageType.String:
			value = elem.LookupParameter(name).AsString()
		elif storeType == StorageType.Integer:
			value  = elem.LookupParameter(name).AsInteger()
		elif storeType == StorageType.Double:
			value = elem.LookupParameter(name).AsDouble()

	#Параметр встроенный
	except:
		bip = GetBuiltInParam(name)
		storeType = elem.get_Parameter(bip).StorageType
		if storeType == StorageType.String:
			value = elem.get_Parameter(bip).AsString()
		elif storeType == StorageType.Integer:
			value  = elem.get_Parameter(bip).AsDouble()
		elif storeType == StorageType.Double:
			value = elem.get_Parameter(bip).AsDouble()
	return value

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

def UpdateWireParam(_info):
	wire = _info[0]
	param = _info[1]
	mepSys = wire.MEPSystem
	sysVal = GetParVal(mepSys, param)
	SetupParVal(wire, param, sysVal)
	return sysVal

class ElSys:
	"""Описывает электрическую систему и её параметры"""
	def __init__(self, rvtSys):
		self.Sys = rvtSys
		self.SysNumber = rvtSys.Name
		self.board = rvtSys.BaseEquipment
		self.cos = rvtSys.PowerFactor
		self.Poles = rvtSys.PolesNumber
		self.ToCalculate = rvtSys.LookupParameter("E_IsLocked")\
							.AsValueString()
		self.P_apparent = None
		self.S_apparent = None
		self.I_apparent = None
		self.CBType = None
		self.CBCurrent = None
		self.isChanged = True
		self.IsLocked = GetParVal(rvtSys, "E_IsLocked")
		
		loadClass = self.Sys.LoadClassifications
		loadClass = loadClass + ";"
		pattern = re.compile(r"(.*?)(?:;\s?)")
		res = sorted(pattern.findall(loadClass))
		self.loadClassStr = "; ".join(res)
		self.brds = self.GetBoardsInSys()

		self.Test = None

	def SetApparentValues(self):
		global doc
		global testBoard
		calcSystem = self.Sys
		phaseN = self.Poles
		#необходимо запомнить, к какому щиту подключена группа
		calcBoard = self.board
		#группа переключается на вспомогательный щит
		calcSystem.SelectPanel(testBoard)
		doc.Regenerate()
		
		#Далее стандартный расчёт электрических параметров
		rvtS = testBoard.get_Parameter(\
						BuiltInParameter.\
						RBS_ELEC_PANEL_TOTALESTLOAD_PARAM).AsDouble()
		rvtS = UnitUtils.ConvertFromInternalUnits(\
						rvtS, DisplayUnitType.DUT_VOLT_AMPERES)
		self.S_apparent = rvtS
		cos_Phi = self.cos
		self.P_apparent = rvtS*cos_Phi
		Papp = self.P_apparent

		if phaseN == 1:
			Iapp = Papp/(cos_Phi*220)
		else:
			Iapp = Papp/(cos_Phi*380*sqrt(3))
		self.I_apparent = round(Iapp, 2)

		#Запись расчётных значений в параметры группы
		calcSystem.LookupParameter("E_EstimatedPower").\
								Set(self.P_apparent)
		calcSystem.LookupParameter("E_EstimatedCurrent").\
								Set(self.I_apparent)
		#Возвращение группы на прежнее место в старом щите
		calcSystem.SelectPanel(calcBoard)
		doc.Regenerate()

	def SetCBType(self):
		"""
		Выбор аппарата защиты на основании класса нагрузки
		освещение 1ф - QF (автомат)
		ОВК 1ф - QF (автомат)
		системы управления 1ф - К (контактор)
		всё остальное - QD (дифавтомат)
		
		Если имя системы равно имени щита, то это
		питающая система щита - QS
		"""
		if self.IsLocked: return None
		global doc
		calcSystem = self.Sys
		global boards
		sysBoard = self.board
		phaseN = self.Poles
		loadClassStr = self.loadClassStr

		caseQF = [
				"Beleuchtung",
				"Beleuchtung; Sonstige"
				]
		caseQD = [
				"Power",
				"TEST"
				]

		if loadClassStr in caseQF:
			calcSystem.LookupParameter("MC CB Type").Set("QF")
			calcSystem.LookupParameter("MC Has RCD").Set(False)

		elif loadClassStr in caseQD:
			calcSystem.LookupParameter("MC CB Type").Set("QD")
			calcSystem.LookupParameter("MC Has RCD").Set(True)

		else:
			calcSystem.LookupParameter("MC CB Type").Set("N\A")
			calcSystem.LookupParameter("MC Has RCD").Set(False)

		#Для щитов по умолчанию QF
		#Find all boards in system
		elems = calcSystem.Elements
		bic = BuiltInCategory.OST_ElectricalEquipment
		brdCat = Category.GetCategory(doc, bic).Id
		if elems:
			boardsInSys = [i.Category.Id for i in elems
				if i.Category.Id == brdCat]
		if boardsInSys:
			calcSystem.LookupParameter("MC CB Type").Set("QF")
			calcSystem.LookupParameter("MC Has RCD").Set(False)
		doc.Regenerate()
		self.Test = "Beleuchtung; Sonstige" == str(loadClassStr)
		
	def GetBoardsInSys(self):
		"""
		Получение щитов, подключённых к сети
		"""
		global boards
		#all elements in system
		elems = self.Sys.Elements
		
		bic = BuiltInCategory.OST_ElectricalEquipment
		brdCat = Category.GetCategory(doc, bic).Id
		if elems:
			boardsInSys = [i.Category.Id for i in elems if i.Category.Id == brdCat]
			boardsId = [brd.Board.Id for brd in boards]
			outlist = [elem for elem in elems if elem.Id in boardsId]
			return outlist
		else:
			return None

	def SetMinimalQF(self):
		"""
		Минимальное значение уставки:
		для освещения и управления - 10А
		для щитов - 20А
		для всего остального - 16А
		"""
		global doc
		calcSystem = self.Sys
		calcBoard = self.board
		loadClass = self.loadClassStr
		
		if self.IsLocked:
			parameter = "RBS_ELEC_CIRCUIT_RATING_PARAM"
			self.CBCurrent = GetParVal(calcSystem, parameter)
			return None
		
		caseTen = [
				"Beleuchtung",
				"Beleuchtung; Sonstige"
				]
		phaseN = self.Poles
		
		#Если это щит, то минимальная уставка 20
		if phaseN == 3 and self.brds: 
			current = 16
		elif phaseN == 3:
			current = 16
		elif loadClass in caseTen:
			current = 10
		else:
			current = 16
		self.CBCurrent = current
		
		parameter = BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM
		calcSystem.get_Parameter(parameter).Set(current)
		calcSystem.LookupParameter("MC Frame Size").Set(current)
		doc.Regenerate()
		self.Test = loadClass

	def SetQFByI(self):
		"""
		Корректировка уставок по расчётному току.
		"""
		
		if self.IsLocked: return None
		global doc
		global standartQF
		global boards
		calcSystem = self.Sys
		I_apparent = self.I_apparent
		I_system = self.CBCurrent
		current = 0

		if I_apparent*1.2 > I_system:
			current = min([i for i in standartQF if i>I_apparent*1.2])
			self.CBCurrent = current
			SetupParVal(calcSystem, "RBS_ELEC_CIRCUIT_RATING_PARAM", current)
			SetupParVal(calcSystem, "MC Frame Size", current)
		
		doc.Regenerate()

	def QFBySelectivity(self):
		"""
		Проверка на селективность.
		В алгоритме селективность проверяется многократно.
		До тех пор, пока во всех єлектрических системах
		не будет изменений.
		"""
		global doc
		global boards
		global standartQF
		parameter = "BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM"

		#Принимается, что в системе нет изменений
		self.isChanged = False
		#Найти все щиты, в группе (шлейф)
		calcSystem = self.Sys
		rvtBoards = self.brds
		rvtBrdId = [i.Id for i in rvtBoards]
		upperBoard = self.board
		objUpperBrd = [brd for brd in boards if upperBoard.Id == brd.Board.Id][0]
		downBrds = [brd for brd in boards if brd.Board.Id in rvtBrdId]
		#если щитов нет - проверка на селективность не нужна
		if not downBrds: return None
		isD = all([brd.IsDisconnector for brd in downBrds])
		
		#Во всех щитах найти аппараты на вводе.
		brdQF = [int(brd.QF) for brd in downBrds]
		#Найти максимальный аппарат на вводе.
		brdMax = max(brdQF)
		#Всем щитам поменять аппарат на максимальный.
		for brd in downBrds:
			brd.QF = brdMax
			param = "MC Frame Size"
			SetupParVal(brd.Board, param, brdMax)

		#системе назначить аппарат на ступень больше максимального
		if not(self.IsLocked) and not isD:
			sysNewQF = [i for i in standartQF if i > brdMax][0]
			self.CBCurrent = int(sysNewQF)
			param = "RBS_ELEC_CIRCUIT_RATING_PARAM"
			SetupParVal(calcSystem, param, int(sysNewQF))
			param = "MC Frame Size"
			SetupParVal(calcSystem, param, int(sysNewQF))
			doc.Regenerate()
			boardNewQF = [i for i in standartQF if i > sysNewQF][0]
			if objUpperBrd.QF < boardNewQF:
				objUpperBrd.QF = boardNewQF
				objUpperBrd.Board.LookupParameter("MC Frame Size").Set(boardNewQF)
				#установить отметку, что в системе произошли изменения.
				self.isChanged = True
			else: return None
		
		#системе назначить аппарат равный максимальному
		elif not(self.IsLocked) and isD:
			sysNewQF = brdMax
			self.CBCurrent = int(sysNewQF)
			param = "RBS_ELEC_CIRCUIT_RATING_PARAM"
			SetupParVal(calcSystem, param, int(sysNewQF))
			param = "MC Frame Size"
			SetupParVal(calcSystem, param, int(sysNewQF))
			boardNewQF = [i for i in standartQF if i > sysNewQF][0]
			doc.Regenerate()
			if objUpperBrd.QF < boardNewQF:
				objUpperBrd.QF = boardNewQF
				objUpperBrd.Board.LookupParameter("MC Frame Size").Set(boardNewQF)
				#установить отметку, что в системе произошли изменения.
				self.isChanged = True
			else: return None
			
		#системе назначить аппарат равный максимальному
		elif self.IsLocked:
			if isD:
				brdMax = self.CBCurrent
				for brd in downBrds:
						brd.QF = brdMax
						param = "MC Frame Size"
						SetupParVal(brd.Board, param, brdMax)
				boardNewQF = [i for i in standartQF if i > self.CBCurrent][0]
				if objUpperBrd.QF < boardNewQF:
					objUpperBrd.QF = boardNewQF
					objUpperBrd.Board.LookupParameter("MC Frame Size").Set(boardNewQF)
					#установить отметку, что в системе произошли изменения.
					self.isChanged = True
				else: return None
			else:
				brdMax = max([i for i in standartQF if i < self.CBCurrent])
				for brd in downBrds:
					brd.QF = brdMax
					param = "MC Frame Size"
					SetupParVal(brd.Board, param, brdMax)
				boardNewQF = [i for i in standartQF if i > self.CBCurrent][0]
				if objUpperBrd.QF < boardNewQF:
					objUpperBrd.QF = boardNewQF
					objUpperBrd.Board.LookupParameter("MC Frame Size").Set(boardNewQF)
					#установить отметку, что в системе произошли изменения.
					self.isChanged = True
				else: return None

	def SetCableType(self):
		"""
		Устанавливает сечение кабеля системы на основании 
		уставки автоматического выключателя
		"""
		global doc
		global cableTable
		calcSystem = self.Sys
		cbCurrent = self.CBCurrent
		cableSize = cableTable.get(cbCurrent)

		phaseN = self.Poles
		if phaseN == 3:
			WireType = "NYM 5х"
		elif phaseN == 1:
			WireType = "NYM 3х"
		else:
			""

		cableText = WireType+str(cableSize)
		SetupParVal(calcSystem, "E_CableType", cableText)
		self.Test = WireType

class Board:
	"""Щит электрический"""
	def __init__(self, rvtBoard):
		global doc
		self.Board = rvtBoard
		self.brdId = rvtBoard.Id
		self.Name = rvtBoard.Name
		self.System = self.GetElSys()
		self.IsDisconnector = GetParVal(self.Board, "E_IsDisconnector")
		
		#Минимальное значение автомата на вводе - 20А
		self.QF = 20
		SetupParVal(self.Board, "MC Frame Size", 20)
		doc.Regenerate()
		self.Test = None

	def GetElSys(self):
		"""Получение сети, к которой принадлежит щит"""
		board = self.Board
		brdId = self.brdId
		elSystems = board.MEPModel.ElectricalSystems
		elSystems = [i for i in elSystems if i.BaseEquipment]
		system = [i for i in elSystems if i.BaseEquipment.Id != brdId]
		if system:
			return  system[0]
		else:
			return None

	def CheckQFinBoard (self):
		global doc
		global standartQF
		elSys = self.System
		isD = self.IsDisconnector
		
		#находим систему в списке систем
		try:
			sysUpper = [i for i in systems if i.Sys.Id == elSys.Id][0]
		except:
			return None
		board = self.Board
		boardId = self.brdId
		boardQF = self.QF
		isD = self.IsDisconnector
		isLocked = sysUpper.IsLocked
		sysFrameSize = sysUpper.CBCurrent
		
		#if upper system is locked
		if isLocked and not isD:
			boardNewQF = max([i for i in standartQF if i < sysFrameSize])
			self.QF = boardNewQF
			SetupParVal(self.Board, "MC Frame Size", boardNewQF)
			return None
		elif isLocked and isD:
			boardNewQF = sysFrameSize
			self.QF = boardNewQF
			SetupParVal(self.Board, "MC Frame Size", boardNewQF)
			return None
		else:
			pass
		
		#Рассчетная уставка щита
		boardCalcQF = sysFrameSize
		
		#Estimated current by Selectivity
		#Find all lower systems and get max QF value in
		rvtLowSystems = board.MEPModel.AssignedElectricalSystems
		if not(rvtLowSystems): return None
		lowSystemsId = [i.Id for i in rvtLowSystems]
		lowSystems = [i for i in systems if i.Sys.Id in lowSystemsId]
		lowSysQF = max([i.CBCurrent for i in lowSystems])
		
		#if it is QF
		if not(isD) and lowSysQF >= boardCalcQF:
			boardNewQF = [i for i in standartQF if i > lowSysQF][0]
		elif isD:
			boardNewQF = [i for i in standartQF if i > lowSysQF][0]
		else:
			boardNewQF = boardCalcQF
		
		self.QF = boardNewQF
		SetupParVal(self.Board, "MC Frame Size", boardNewQF)

reload = IN[0]
outlist = list()
standartQF = [10, 16, 20, 25, 32, 
				40, 50, 63, 80, 100,
				125, 160, 200, 250, 315,
				400, 500, 630, 800, 1000,
				1500,2000,2500]
standartСable = ["1.5", "2.5", "2.5", "4", "4", 
				"6", "10", "16","25", "35",
				"50", "70", "95", "120", "150",
				"185", "(2x120)", "(3x95)","(4x95)","(5x95)",
				"(6x95)","(7x95)","(8x95)"]

cableTable = dict(zip(standartQF, standartСable))

#Все электрические сети
#исключить системы, которые не относятся к электрике
#код эл.систем - 6
testParam = BuiltInParameter.RBS_ELEC_CIRCUIT_TYPE
pvp = ParameterValueProvider(ElementId(int(testParam)))
sysRule = FilterIntegerRule(pvp, FilterNumericEquals(), 6)

testParam = BuiltInParameter.RBS_ELEC_CIRCUIT_TYPE
pvp = ParameterValueProvider(ElementId(int(testParam)))
sysRule = FilterIntegerRule(pvp, FilterNumericEquals(), 6)
filter = ElementParameterFilter(sysRule)

rvtAllSystems = FilteredElementCollector(doc).\
			OfCategory(BuiltInCategory.OST_ElectricalCircuit).\
			WhereElementIsNotElementType().WherePasses(filter).\
			ToElements()

#Дальнейшая обработка систем через питон
#Исключаются не подключенные системы
rvtSystems = [sys for sys in rvtAllSystems if sys.BaseEquipment]

#Все электрические щиты
#Нужны щиты LIC KRA NOT PVA
testParam = BuiltInParameter.ELEM_FAMILY_PARAM
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringContains()
ruleList = list()
ruleList.append(FilterStringRule(pvp, fnrvStr, "LIC", False))
ruleList.append(FilterStringRule(pvp, fnrvStr, "KRA", False))
ruleList.append(FilterStringRule(pvp, fnrvStr, "NOT", False))
ruleList.append(FilterStringRule(pvp, fnrvStr, "PVA", False))
elemFilterList = [ElementParameterFilter(x) for x in ruleList]
filter = LogicalOrFilter(elemFilterList )

rvtAllBoards = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_ElectricalEquipment).\
	WhereElementIsNotElementType().\
	WherePasses(filter).\
	ToElements()

#Дальнейшая обработка щитов через питон
rvtBoardsAll = [
	brd for brd in rvtAllBoards]
#Исключаются не подключенные щиты
rvtBoards = list()
for brd in rvtBoardsAll:
	name = brd.Name
	#Если у щита нет электрических систем выдаст ошибку
	try:
		elSystems = brd.MEPModel.ElectricalSystems
		system = [i for i in elSystems if i.PanelName != name]
		if system:
			rvtBoards.append(brd)
	except: pass

#Создание пустого щита для расчётов. Далее щит будет удален
#Тип щита "хард-кодед" для упрощения задачи
testBrdTyp = "KRA_OHT_0800x0400x2000"
testParam = BuiltInParameter.SYMBOL_NAME_PARAM
pvp = ParameterValueProvider(ElementId(int(testParam)))
fnrvStr = FilterStringEquals()
filter = ElementParameterFilter(FilterStringRule(
									pvp, fnrvStr, testBrdTyp, False))

testBoardType = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_ElectricalEquipment).\
	WhereElementIsElementType().\
	WherePasses(filter).\
	ToElements()[0]

TransactionManager.Instance.EnsureInTransaction(doc)

testBoard = doc.Create.NewFamilyInstance(
		XYZ(0,0,0), testBoardType, Structure.StructuralType.NonStructural) 

boards = [Board(i) for i in rvtBoards]
systems = [ElSys(i) for i in rvtSystems]

#==============CALCULATIONS===============
map(lambda x:x.SetApparentValues(), systems)
map(lambda x:x.SetCBType(), systems)
map(lambda x:x.SetMinimalQF(), systems)
map(lambda x:x.SetQFByI(), systems)

# #Корректировка уставок в щитах
map(lambda x:x.CheckQFinBoard(), boards)

# #Корректировка уставок в системе
changesInSys = True
while changesInSys:
	map(lambda x:x.QFBySelectivity(), systems)
	doc.Regenerate()
	changesInSys = any([i.isChanged for i in systems])
	outlist.append([[i, i.isChanged] for i in systems])

map(lambda x:x.SetCableType(), systems)

#Блок обновления информации в проводе
#Получение всех проводов в проекте\
wireList = FilteredElementCollector(doc).\
	OfCategory(BuiltInCategory.OST_Wire).\
	WhereElementIsNotElementType().\
	ToElements()
wireList = [i for i in wireList if i.MEPSystem]

param = "MC Frame Size"
paramList = [[i, param] for i in wireList]
map(UpdateWireParam, paramList)

param = "E_CableType"
paramList = [[i, param] for i in wireList]
map(UpdateWireParam, paramList)

doc.Delete(testBoard.Id)

TransactionManager.Instance.TransactionTaskDone()

#OUT = outlist
#OUT = [i.Test for i in systems]
#OUT = [i.Test for i in boards]
