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


def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	param = []
	for i in builtInParams:
		if i.ToString() == paramName:
			param.append(i)
			return i


def getParVal(elem, name):
	"""Get parameter value of the element.

	elem - element,
	name - parameter name, or BuiltInParametre str.
	"""
	value = None
	# Параметр пользовательский
	param = elem.LookupParameter(name)
	# параметр не найден. Надо проверить, есть ли такой же встроенный параметр
	if not(param):
		param = elem.get_Parameter(GetBuiltInParam(name))

	# Если параметр найден, считываем значение
	try:
		storeType = param.StorageType
		# value = storeType
		if storeType == StorageType.String:
			value = param.AsString()
		elif storeType == StorageType.Integer:
			value = param.AsDouble()
		elif storeType == StorageType.Double:
			value = param.AsDouble()
		elif storeType == StorageType.ElementId:
			value = param.AsValueString()
	except:
		pass
	return value


def setParVal(elem, name, pValue):
	global doc
	# Параметр пользовательский
	param = elem.LookupParameter(name)
	# параметр не найден. Надо проверить, есть ли такой же встроенный параметр
	if not(param):
		param = elem.get_Parameter(GetBuiltInParam(name))
	if param:
		param.Set(pValue)
	return elem


def getSystems(_brd):
	"""Get all systems of electrical board.

		args:
		_brd - electrical board FamilyInstance

		return:
		list(1, 2) where:
		1 - feeder
		2 - list of branch systems
	"""
	brd_name = getParVal(_brd, "RBS_ELEC_PANEL_NAME")
	try:
		board_all_systems = [i for i in _brd.MEPModel.ElectricalSystems]
	except TypeError:
		raise TypeError("Board \"%s\" have no systems" % brd_name)
	try:
		board_branch_systems = [i for i in _brd.MEPModel.AssignedElectricalSystems]
	except TypeError:
		raise ValueError("Board \"%s\" have no branch systems" % brd_name)
	if len(board_branch_systems) == len(board_all_systems):
		raise ValueError("Board \"%s\" have no feeder" % brd_name)

	board_branch_systems.sort(
		key=lambda x:
		float(getParVal(x, "RBS_ELEC_CIRCUIT_NUMBER")))
	branch_systems_id = [i.Id for i in board_branch_systems]
	board_feeder = [
		i for i in board_all_systems
		if i.Id not in branch_systems_id][0]
	return board_feeder, board_branch_systems


def setPageParam(_lst):
	global doc
	global brd_name
	page = _lst[0]
	pName = _lst[1]
	pNumber = _lst[2]
	page.get_Parameter(BuiltInParameter.SHEET_NAME).Set(pName)
	page.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(pNumber)
	page.LookupParameter("MC Panel Code").Set(brd_name)


def getByCatAndStrParam(_bic, _bip, _val, _isType):
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


def getTypeByCatFamType(_bic, _fam, _type):
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


def create_dia(_brd_name, _brd_sys_index=0):
	"""
	Creates list of diagramm objects for _brd_name

	The function is recursive.
	If to the board other subboards is connected, then`
	function would be called for subboard.

	:attrubutes:
		_brd_name - str Board name
		_brd_sys_list - current tree of the systems
			if list is empty - it is the main board

	:return:
		every new diagramm object would be appended to
		the flat list.
	"""
	global doc
	brd_sys_list = list()
	# start from phase L1
	sch_phase = 1

	try:
		brd_instance = getByCatAndStrParam(
			BuiltInCategory.OST_ElectricalEquipment,
			BuiltInParameter.RBS_ELEC_PANEL_NAME,
			_brd_name, False)[0]
	except ValueError:
		raise ValueError(
			"Board \"%s\" not found "
			% _brd_name)

	brd_circuits_all = getSystems(brd_instance)
	brd_feeder = brd_circuits_all[0]
	brd_circuits = brd_circuits_all[1]

	# check if it is feeder of main board
	if _brd_sys_index == 0:
		sys_upper_index = 1
		diagramm = dia(brd_feeder, sys_upper_index, "Feeder")
		brd_sys_list.append(diagramm)
	else:
		sys_upper_index = _brd_sys_index + 1

	for circuit in brd_circuits:
		diagramm = dia(circuit, sys_upper_index, "Branch")
		brd_sys_list.append(diagramm)

		# set dia_phase
		# if it is 1phase dia - set phase
		if circuit.PolesNumber == 1:
			diagramm.dia_phase = sch_phase
			sch_phase = sch_phase + 1
			# calculating next phase
			if sch_phase > 3:
				sch_phase = 1

		# check if circuit contains subboard
		elems = [elem for elem in circuit.Elements]

		# only 1 board in circuit allowd
		# it is ehought to check first element
		elem = elems[0]
		elem_cat = elem.Category.Id
		elem_panel_code = getParVal(elem, "MC Panel Code")
		brd_bic = BuiltInCategory.OST_ElectricalEquipment
		brd_cat = Autodesk.Revit.DB.Category.GetCategory(
			doc, ElementId(brd_bic)).Id
		check_cat = elem_cat == brd_cat
		check_param = elem_panel_code == _brd_name

		if check_cat and check_param:
			subboard_name = getParVal(elem, "RBS_ELEC_PANEL_NAME")
			subboard_dia_list = create_dia(
				subboard_name,
				sys_upper_index)
			for subboard_dia in subboard_dia_list:
				brd_sys_list.append(subboard_dia)
	return brd_sys_list


class dia:
	"""
		Electrical diagramm class

	:class properties:
		par_to_set - list()
			parameters, that would be read from electrical system
			and would be set to 2D diagramm instance

	:instance properties:
		dia_level - busbur level.
			Main board - is lvl_1

		self.dia_sys - Revit system

		self.dia_description - text diagramm description
			"Feeder" - diagramm for feeder
			"Branch" - diagram for branch circuit
			"Filler" - dummy lines
			"Heater " - first 2D diagramm on the page
			"Footer" - last 2D diagramm on the page

		brd, self.brd_name - board that system connected

		self.dia_index - readable adress of the system in system tree

		self.dia_family_type - type of 2D diagramm to be placed on the page

	:methods:
		get_type() - get Type of 2D diagramm.
			read text parameters in Revit systems
			find FamilyType in Revit

		get_parameters() - get parameters value from systems
	"""

	par_to_set = list()
	par_to_set.append("RBS_ELEC_CIRCUIT_NAME")
	# par_to_set.append("RBS_ELEC_CIRCUIT_NUMBER")
	par_to_set.append("RBS_ELEC_CIRCUIT_WIRE_TYPE_PARAM")
	par_to_set.append("E_CableSize")
	par_to_set.append("E_Sch_Nennstrom")
	par_to_set.append("E_Sch_Schutztyp")
	par_to_set.append("E_Sch_Schutztyp_kurz")
	par_to_set.append("E_Sch_Elektrischen Schlag ")
	# par_to_set.append("E_Stromkreisprefix")

	def __init__(self, _rvtSys, _brd_lvl, _description):
		self.dia_level = _brd_lvl
		self.dia_sys = _rvtSys
		self.circuit_number = None
		self.dia_description = _description
		self.brd = None
		self.brd_name = None
		self.param_list = list()
		self.dia_family_type = None
		self.dia_module_size = int()
		self.dia_inst = None
		dia.dia_phase = int()

		if _description == "Branch":
			self.circuit_number = str(getParVal(
				_rvtSys,
				"RBS_ELEC_CIRCUIT_NUMBER"))
			self.brd = self.dia_sys.BaseEquipment
			self.brd_name = self.brd.Name

		elif _description == "Feeder":
			self.circuit_number = None
			global MAIN_BRD_INST
			self.brd = MAIN_BRD_INST
			self.brd_name = self.brd.Name

		else:
			pass

		if self.circuit_number:
			self.dia_index = str(self.dia_level) + "." + self.circuit_number
		else:
			self.dia_index = str(self.dia_level)

	def get_type(self):
		dia_description = self.dia_description
		sch_family = None
		sch_type = None

		# for "Einspeisung" schema
		# diagramm is writen in board parameter
		if dia_description == "Feeder":
			sch_family = self.brd.LookupParameter("E_Sch_Family").AsString()
			sch_type = self.brd.LookupParameter("E_Sch_FamilyType").AsString()

		# for "Electrical" schema
		elif dia_description == "Branch":
			# diagramm is writen in electrical system parameter
			sch_family = self.dia_sys.LookupParameter("E_Sch_Family").AsString()
			sch_type = self.dia_sys.LookupParameter("E_Sch_FamilyType").AsString()

		# for "Filler"
		elif dia_description == "Filler":
			# diagramm is hardcoded
			sch_family = "E_SCH_Filler"
			sch_type = "Filler_1modul"

		# for "Header"
		elif dia_description == "Header":
			# diagramm is hardcoded
			sch_family = "E_SCH_Filler"
			sch_type = "Filler_Start"

		# for "Footer"
		elif dia_description == "Footer":
			# diagramm is hardcoded
			sch_family = "E_SCH_Filler"
			sch_type = "Filler_End"

		else:
			pass

		if not(sch_family) or not(sch_type):
			raise ValueError(
				"Parameters for 2D diagram are empty for {},{}".
				format(self.brd_name, self.circuit_number))

		self.dia_family_type = getTypeByCatFamType(
			BuiltInCategory.OST_GenericAnnotation,
			sch_family,
			sch_type)
		if not(self.dia_family_type):
			raise ValueError(
				"2D diagram FamilyType not found for {},{}".
				format(self.brd_name, self.circuit_number))

		self.dia_module_size = self.dia_family_type.\
			LookupParameter("E_PositionsHeld").AsInteger()
		return self.dia_family_type

	def get_parameters(self):
		"""Get parameters for Main board and brunch systems"""
		description = self.dia_description

		# for system if description is "Branch"

		self.param_list = [
			[x, getParVal(self.dia_sys, x)]
			for x in dia.par_to_set]

		if description == "Branch":
			# Update circuit number.
			# Circuit number = busbur number + system number
			# busbur number is panel Prefix
			# circuit number - 2 symbols.
			sys_prefix = self.brd.get_Parameter(
				BuiltInParameter.RBS_ELEC_CIRCUIT_PREFIX).AsString()
			sys_num = '{:02}'.format(int(self.circuit_number))
			self.param_list.append(["RBS_ELEC_CIRCUIT_NUMBER", sys_num])
			self.param_list.append(["E_Stromkreisprefix", sys_prefix])
			self.param_list.append(["MC Phases", str(self.dia_phase)])

		elif description == "Feeder":
			self.param_list.append(["RBS_ELEC_CIRCUIT_NUMBER", "Einspeisung"])

		else:
			pass

	def placeDiagramm(self):
		global doc
		global sheetLst
		self.diaInst = doc.Create.NewFamilyInstance(
			self.schType,
			sheetLst[self.pageN])

	def set_parameters(self):
		elem = self.dia_inst

		if self.dia_level and elem.LookupParameter("E_StromschieneEbene"):
			elem.LookupParameter("E_StromschieneEbene").Set(self.dia_level)

		# check if it is any parameter to be set
		if not(self.param_list):
			return None

		for i, j in self.param_list:
			if not(j):
				j = " "
			setParVal(elem, i, j)
		return elem


class page:
	"""Page class contains info and methods for creating pages"""

	# coordinates of points on scheet
	coord_list = list()
	coord_list.append(XYZ(0.0738188976375485, 0.66929133858268, 0))
	coord_list.append(XYZ(0.113188976377707, 0.66929133858268, 0))
	coord_list.append(XYZ(0.204396325459072, 0.66929133858268, 0))
	coord_list.append(XYZ(0.295603674540437, 0.66929133858268, 0))
	coord_list.append(XYZ(0.386811023621802, 0.66929133858268, 0))
	coord_list.append(XYZ(0.478018372703167, 0.66929133858268, 0))
	coord_list.append(XYZ(0.569225721784533, 0.66929133858268, 0))
	coord_list.append(XYZ(0.660433070865899, 0.66929133858268, 0))
	coord_list.append(XYZ(0.751640419947264, 0.66929133858268, 0))
	coord_list.append(XYZ(0.84284776902863, 0.66929133858268, 0))
	coord_list.append(XYZ(0.0738188976375485, 0.66929133858268, 0))

	total_pages = None
	existing_sheets = None

	title_first_page = getByCatAndStrParam(
		BuiltInCategory.OST_TitleBlocks,
		BuiltInParameter.SYMBOL_NAME_PARAM,
		"WSP_Plankopf_Shema_Titelblatt", True)[0]

	title_page = getByCatAndStrParam(
		BuiltInCategory.OST_TitleBlocks,
		BuiltInParameter.SYMBOL_NAME_PARAM,
		"WSP_Plankopf_Shema", True)[0]

	@classmethod
	def get_existing_sheets(cls, _brd_name):
		"""Get existing lists for diagramm

		If the page parameter "MC Panel Code" contains board name,
		then the page was made for correct board.
		"""
		global doc
		existing_sheets = [
			i for i in FilteredElementCollector(doc).
			OfCategory(BuiltInCategory.OST_Sheets).
			WhereElementIsNotElementType().
			ToElements()
			if i.LookupParameter("MC Panel Code").
			AsString() == _brd_name]
		existing_sheets.sort(key=lambda x: x.SheetNumber)

		cls.existing_sheets = existing_sheets
		return existing_sheets

	@classmethod
	def get_total_pages(cls, _dia_list):
		"""Calculate total ammount of pages"""
		modules_total = sum([x.dia_module_size for x in _dia_list])
		cls.total_pages = int(math.ceil(modules_total / 8.0)) + 1

	@classmethod
	def divide_pro_page(cls, _dia_list):
		"""Get list of diagramms -> list of diagramms pro page"""
		outlist = [[]]
		modules_on_page = 0
		current_page = 0
		for d in _dia_list:
			modules_on_page += d.dia_module_size
			if modules_on_page <= 8:
				outlist[current_page].append(d)
			else:
				outlist.append([])
				modules_on_page = d.dia_module_size
				current_page += 1
				outlist[current_page].append(d)
		return outlist

	@classmethod
	def check_page_ammount(cls, _create_new):
		"""Check if existing sheets is enough to place all dias"""

		existing_sheets = cls.existing_sheets
		total_pages = cls.total_pages
		if existing_sheets:
			ex_sh_ammount = len(existing_sheets)
		else:
			ex_sh_ammount = 0
		if _create_new:
			# it does not metter
			return True
		elif not(_create_new) and ex_sh_ammount >= total_pages:
			# it is enought
			return True
		elif not(_create_new) and ex_sh_ammount < total_pages:
			raise TypeError("Not enought existing pages")

	@classmethod
	def delete_not_used(cls, sheets_in_use):
		for sheet in cls.existing_sheets:
			if sheet not in sheets_in_use:
				doc.Delete(sheet.Id)

	def __init__(self, _page_number):
		self.page_number = _page_number
		self.dia_on_page = None
		self.sheet_inst = None

	def fill_page(self):
		"""Create list of diagramms to be put on the page

		Create list of diagramms, footers and fillers for current page.
		If self.page_number == 0 - it is the main page. No diagramms there
		else - it is a page with diagramms
		Header need to be placed for every page with diagramms
		Maximum module size for A4 paper is 9.
		If number of modules is less then 9 - fillers are created.

		return:
			list of elements on the current page
		"""

		global dia_list
		page_n = self.page_number
		dia_on_page = list()

		# =========Place branch diagramms
		if page_n > 0:
			branch_dia = page.divide_pro_page(dia_list)[page_n - 1]
			map(lambda x: dia_on_page.append(x), branch_dia)

		# =========Place header
		# if it is the page 1 - no need to set header
		if page_n > 0:
			header = dia(None, None, "Header")
			header.get_type()
			# header location is allways module 0.
			dia_on_page.insert(0, header)

			# header level = level of 1 diagramm on the page
			frst_dia_lvl = dia_on_page[1].dia_level
			header.dia_level = frst_dia_lvl

			# if page number > 1 - set previous page number
			if page_n > 1:
				header.param_list.append(
					["E_Page_previous", str(int(page_n - 1))])

		# =========Place footer
		if page_n > 0:
			# footer location - next to the last diagramm.
			footer = dia(None, None, "Footer")
			footer.get_type()

			if page_n != page.total_pages - 1:
				# if it is the not the last page
				# set next page
				footer.param_list.append(
					["E_Page_next", str(int(page_n + 1))])

				# dia level == the dia of the next page
				next_page_first = page.divide_pro_page(dia_list)[page_n][0]
				footer.dia_level = next_page_first.dia_level

			else:
				# if it is the last page - no need to set up next page number
				# dia level == level of the last elem on page
				last_on_page = dia_on_page[-1]
				footer.dia_level = last_on_page.dia_level
			dia_on_page.append(footer)

		# =========Place filler
		if dia_on_page:
			# check how much free modules available
			dia_used_modules = sum([
				x.dia_module_size for x in dia_on_page])
			dia_free_modules = 10 - dia_used_modules
			for _ in range(dia_free_modules):
				filler = dia(None, None, "Filler")
				filler.get_type()
				dia_on_page.append(filler)

		self.dia_on_page = dia_on_page
		return dia_on_page

	def get_sheet(self, _create_new):
		global doc
		global MAIN_BRD_NAME
		page_number = self.page_number
		total_pages = page.total_pages
		sheet_num = MAIN_BRD_NAME + "_" + str(page_number).zfill(3)
		sheet_name = MAIN_BRD_NAME

		# get title block
		if page_number == 0:
			title_block = page.title_first_page
		else:
			title_block = page.title_page

		# find current sheet
		existing_sheets = page.existing_sheets
		if existing_sheets and page_number <= len(existing_sheets) - 1:
			current_sheet = existing_sheets[page_number]
		else:
			current_sheet = None

		if _create_new and current_sheet:
			# delete old page and create new
			doc.Delete(current_sheet.Id)
			current_sheet = ViewSheet.Create(doc, title_block.Id)
			existing_sheets[page_number] = current_sheet

		elif _create_new and not(current_sheet):
			# create new pages only if no existing pages found
			current_sheet = ViewSheet.Create(doc, title_block.Id)
			existing_sheets.append(current_sheet)

		else:
			# existing sheet found.
			# remove all instances on existing sheet.
			elemsOnSheet = list()
			elems = FilteredElementCollector(doc)\
				.OwnedByView(current_sheet.Id)\
				.OfCategory(BuiltInCategory.OST_GenericAnnotation)\
				.WhereElementIsNotElementType().ToElementIds()
			map(lambda x: elemsOnSheet.append(x), elems)
			typed_list = List[ElementId](elemsOnSheet)
			doc.Delete(typed_list)
			self.sheet_inst = current_sheet
			return current_sheet

		# set new sheet parameters
		param_list = list()
		param_list.append(["MC Number of Pages", str(total_pages - 1)])
		param_list.append(["MC Page Number", str(page_number)])
		param_list.append(["SHEET_NAME", sheet_name])
		param_list.append(["SHEET_NUMBER", sheet_num])
		param_list.append(["MC Panel Code", MAIN_BRD_NAME])
		map(lambda x: setParVal(current_sheet, x[0], x[1]), param_list)
		self.sheet_inst = current_sheet
		return current_sheet

	def create_2D(self):
		global doc
		dia_on_page = self.dia_on_page
		if not(dia_on_page):
			return None
		position_index = 0
		outlist = list()
		for dia in dia_on_page:
			insert_point = page.coord_list[position_index]
			dia.dia_inst = doc.Create.NewFamilyInstance(
				insert_point,
				dia.dia_family_type,
				self.sheet_inst)
			outlist.append([insert_point, dia.dia_family_type, self.sheet_inst])
			if position_index == 0:
				# it is start position
				position_index += 1
			else:
				position_index += dia.dia_module_size
		return outlist

	def set_par_for_all_dia_on_page(self):
		global doc
		doc.Regenerate()
		dia_on_page = self.dia_on_page
		if not(dia_on_page):
			return None
		return [x.set_parameters() for x in dia_on_page]


MAIN_BRD_NAME = IN[0]
MAIN_BRD_INST = getByCatAndStrParam(
	BuiltInCategory.OST_ElectricalEquipment,
	BuiltInParameter.RBS_ELEC_PANEL_NAME,
	MAIN_BRD_NAME,
	False)[0]
create_new_sheets = IN[1]
reload = IN[2]

outlist = list()
footers = list()
outlist = list()

# ========Initialaise dia class
dia_list = create_dia(MAIN_BRD_NAME)
map(lambda x: x.get_type(), dia_list)
map(lambda x: x.get_parameters(), dia_list)

# ========Initialaise page class
total_pages = page.get_total_pages(dia_list)
existing_sheets = page.get_existing_sheets(MAIN_BRD_NAME)

# check is it enough existing pages
page.check_page_ammount(create_new_sheets)
page_list = [page(i) for i in range(page.total_pages)]
map(lambda x: x.fill_page(), page_list)

# =========Start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

sheet_list = map(lambda x: x.get_sheet(create_new_sheets), page_list)
# delete sheets, that not in use any more
page.delete_not_used(sheet_list)
list_2D = map(lambda x: x.create_2D(), page_list)
map(lambda x: x.set_par_for_all_dia_on_page(), page_list)

TransactionManager.Instance.TransactionTaskDone()
# =========End transaction


OUT = [x.dia_on_page for x in page_list]
# OUT = [x.param_list for x in dia_list]
