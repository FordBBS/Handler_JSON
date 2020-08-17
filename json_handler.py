####################################################################################################
#                                                                                                  #
# Python, JSON handler, BBS					   						                               #
# GitHub: https://github.com/FordBBS/Handler_JSON 											       #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/08/08, BBS:	- First release
# 					- Implemented Data Manipulation, Importing, Exporting function groups
# 2020/08/09, BBS: 	- Implemented following functions completely
# 						'IBase_update_dict_by_path', 'IBase_merge_dict_single_path', 
# 						'IBase_get_keylist_of_dict_single_path'
# 2020/08/16, BBS:	- Cplt Issue #2, #4
# 					- Improve 'hs_create_dict_from_serial_list' for array result
# 					- Improve 'IBase_validate_dict_key_exist'
# 					- Implemented 'hs_get_formatted_listpair'
# 					- Implemented 'IBase_transform_dict_to_list'
# 					- Implemented 'IBase_manipulate_dicts_target_key'
# 					- Implemented 'IBase_operation_union_dicts'
# 					- Remove 'IBase_update_dict_by_path', 'IBase_merge_dict_single_path'
#
#***************************************************************************************************

#*** Function Group List ***************************************************************************
# - Helper
# - Importing
# - Exporting
# - Data Manipulation



#*** Library Import ********************************************************************************
#--- Operating system ------------------------------------------------------------------------------
import  os
import  json

#--- BBS Modules -----------------------------------------------------------------------------------
import  sys
sys.path.append(r"C:\Backup\03 SelfMade_Tools\Python\BBS_Modules")

# OS handler
import 	os_handler 			as hs_os

# Dataset handler
import  dataset_handler 	as hs_dataset
from    dataset_handler  	import IBase_get_reduced_list			as bbs_reduce
from 	dataset_handler 	import IBase_list_remove_duplicate  	as bbs_remove_dupl



#*** Function Group: Helper ************************************************************************
def getconst_join_character():
	RetVal = "."
	return RetVal

def hs_create_dict_from_serial_list(listInput):
	#*** Input-Validation **************************************************************************
	if not isinstance(listInput, list): return {}

	#*** Initialization ****************************************************************************
	dictRes = {}
	mainKey = listInput[0]
	thisVal = listInput[-1]

	#*** Operations ********************************************************************************
	#--- Case: more than 2 elements ----------------------------------------------------------------
	if len(listInput) > 2:
		listPrepVal = []

		if not isinstance(thisVal, list): thisVal = [thisVal]
		for eachVal in thisVal:
			tmpDict = {}
			tmpDict[listInput[-2]] = eachVal

			for idx in range(len(listInput) - 3, 0, -1):
				curKey  = listInput[idx]
				curDict = tmpDict.copy()
				tmpDict = {}
				tmpDict[curKey] = curDict
			listPrepVal.append(tmpDict)

		if len(listPrepVal) == 1: listPrepVal = listPrepVal[0]
		listRes = [mainKey, listPrepVal]

	#--- Case: only 2 elements ---------------------------------------------------------------------
	else: listRes = [mainKey, thisVal]

	#--- Release -----------------------------------------------------------------------------------
	return listRes

def hs_get_formatted_listpair(listPair, flg_allstring, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Return a formatted Parameter-Value list for further work in this module

	[list] listPair, 	  Target Parameter-Value list to be formatted
	[bool] flg_allstring, True: Convert all values into String, False: no conversion
	[int]  flg_case,  	  0: Do nothing, 1: Convert all string to upper case, 2: ... to lower case

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(listPair, list) and len(listPair) >= 2): return [[], []]

	#*** Initialization ****************************************************************************
	if isinstance(flg_allstring, int):
		if flg_allstring != 0: flg_allstring = True
		else: flg_allstring = False
	
	if not isinstance(flg_allstring, bool): flg_allstring = False
	if not (isinstance(flg_case, int) and flg_case in range(0, 3)): flg_case = 0

	listParam = listPair[0]
	listValue = listPair[1]
	listPrep = [[], []]

	#*** Operations ********************************************************************************
	for eachKey in listParam:
		if not eachKey in listPrep[0]:
			idx = 0
			listPrep[0].append(eachKey)
			listPrep[1].append([])

			while idx >= 0:
				try:
					idx 	= listParam.index(eachKey, idx)
					eachVal = listValue[idx]

					if flg_allstring: eachVal = str(eachVal)
					if flg_case == 1: eachVal = eachVal.upper()
					elif flg_case == 2: eachVal = eachVal.lower()

					listPrep[1][len(listPrep[1]) - 1].append(eachVal)
					idx += 1
				except: idx = -1
	return listPrep



#*** Function Group: Importing *********************************************************************
def IUser_read_json(pathFile):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Return a dictionary of JSON content available in 'pathFile' JSON's file path

	[str] pathFile,	A path of target JSON file to be read

	'''

	#*** Input Validation **************************************************************************
	if not os.path.isfile(pathFile): return 101

	#*** Initialization ****************************************************************************
	# Nothing to be initialized

	#*** Operations ********************************************************************************
	with open(pathFile) as curJson:
		thisContent = json.load(curJson)
		curJson.close()
	return thisContent 



#*** Function Group: Exporting *********************************************************************
def IUser_write_json(objJson, pathDest, nameFile, flg_tryremove, flg_timestamp):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Create a json file of 'objJson' content at 'pathDest' directory with 'nameFile' as a filename
		If there is already an existence file, renaming mechanism will be done automatically based
		on 'flg_tryremove' and 'flg_timestamp' parameters

	[str]  objJson,  JSON content
	[str]  pathDest, Target destination where this json file shall be located at
					 Desktop's directory is used as a default
	[str]  nameFile, Target filename, "json_" + timestamp is used as a default
	[bool] flg_tryremove, True: Try to delete existing file first, False: Not try
	[bool] flg_timestamp, True: Use timestamp first after found same 'nameFile', False: Apply counter
							immediately

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(objJson, str): return False

	#*** Initialization ****************************************************************************
	chr_path     = hs_os.getconst_chr_path()[1]
	strTimestamp = hs_os.IBase_get_timestamp("yyyymmdd_hhmmss")
	pathDefault  = hs_os.IBase_get_desktop_path()
	nameDefault  = "json_" + strTimestamp

	#*** Operations ********************************************************************************
	#--- Preparation: Destination ------------------------------------------------------------------
	if not (isinstance(pathDest, str) and len(pathDest) > 0): pathDest = pathDefault
	if not os.path.exists(pathDest): pathDest = os.path.split(pathDest)[0] + chr_path \
	 											+ IUser_create_folder(pathDest, "", False, True)

	#--- Preparation: Filename ---------------------------------------------------------------------
	if not (isinstance(nameFile, str) and len(nameFile) > 0): nameFile = nameDefault

	#--- Create JSON file --------------------------------------------------------------------------
	RetVal = hs_os.IUser_create_file_fromstr(nameFile, pathDest, objJson, ".json", \
												flg_tryremove, flg_tryremove)

	#--- Release -----------------------------------------------------------------------------------
	if isinstance(RetVal, str): return True
	else: return False



#*** Function Group: Data Manipulation *************************************************************
def IBase_get_keylist_of_dict_single_path(mainDict):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Return a list of key available on 'mainDict'. Only the first found path will be done

		Result Format
		[listFoundKey, value of deepest dictionary' key]

	[dict] mainDict, Main dictionary

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(mainDict, dict): return []

	#*** Initialization ****************************************************************************
	listRes = []
	dictRes = mainDict.copy()

	#*** Operations ********************************************************************************
	while True:
		for key in dictRes:
			listRes.append(key)
			dictRes = dictRes[key]
			if not isinstance(dictRes, dict): return [listRes, dictRes]
			break

def IBase_validate_dict_key_exist(mainDict, tarDict):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Return True if keypath in 'tarDict' exists in 'mainDict', False vice versa
		For example,
		Case 1)
			mainDict = {"role": "Bag", "table": {"delaytime": "no_delay"}}
			tarDict  = {"table": {"range": "Modal"}}
			Result   = False
		
		Case 2)
			mainDict = {"role": "Bag", "table": {"delaytime": "no_delay"}}
			tarDict  = {"table": {"delaytime": "delay_1"}}
			Result   = True

		Case 3)
			mainDict = {"role": "Bag", "table": {"delaytime": "no_delay"}}
			tarDict  = {"role": "Diluted;Bag"}
			Result   = True

	[dict] mainDict, Main dictionary
	[dict] tarDict,  Dictionary to be validated

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(mainDict, dict) and isinstance(tarDict, dict)): return 101

	#*** Initialization ****************************************************************************
	chr_join = getconst_join_character()
	blnRes   = False

	#*** Operations ********************************************************************************
	#--- Get info of 'tarDict' ---------------------------------------------------------------------
	keypath = chr_join.join(IBase_get_keylist_of_dict_single_path(tarDict)[0])

	#--- Get Parameter-Value list of 'mainDict' ----------------------------------------------------
	listParam = IBase_transform_dict_to_list(mainDict)[0]

	#--- Validation --------------------------------------------------------------------------------
	blnRes = keypath in listParam

	#--- Release -----------------------------------------------------------------------------------
	return blnRes

def IBase_manipulate_dicts_target_key(dictA, dictB, strKey, flg_mode):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Manipulate target key on dictionary A and B and return the result for 'strKey' key path
		Result is manipulated depends on 'flg_mode' operation mode

	[dict] dictA, 	 Dictionary A
	[dict] dictB, 	 Dictionary B
	[str]  strKey,   A string of key path (e.g. "ecs", "ecs.devicename", "gmd.table.range")
	[int]  flg_mode, 0: Intersection, 1: Union

	'''

	#*** Input Validation **************************************************************************
	if not (isinstance(dictA, dict) and isinstance(dictB, dict) \
			and (isinstance(strKey, str) and len(strKey) > 0)): return {}

	#*** Initialization ****************************************************************************
	if not (isinstance(flg_mode, int) and flg_mode in range(0, 2)): flg_mode = 1

	chr_join = getconst_join_character()
	flg_sub  = False
	resVal   = "default"

	#*** Operations ********************************************************************************
	#--- Get listPair of both dictionaries ---------------------------------------------------------
	listPairA = IBase_transform_dict_to_list(dictA)
	listPairB = IBase_transform_dict_to_list(dictB)

	#--- Ensure 'strKey' exists on both dictionaries -----------------------------------------------
	flg_val_A = strKey in listPairA[0]
	flg_val_B = strKey in listPairB[0]

	if flg_val_A: idx_a = listPairA[0].index(strKey)
	if flg_val_B: idx_b = listPairB[0].index(strKey)

	# Both dicts have no 'strKey' then nothing left for any mode
	if not (flg_val_A or flg_val_B): resVal = ""

	# Both dicts must have 'strKey' on Intersection mode
	elif not (flg_val_A and flg_val_B) and flg_mode == 0: resVal = ""

	# Only one dict has 'strKey' on Union mode, return immediately that dict
	elif flg_val_A and not flg_val_B: resVal = listPairA[1][idx_a]
	elif flg_val_B and not flg_val_A: resVal = listPairB[1][idx_b]

	#--- Manipulation ------------------------------------------------------------------------------
	if resVal == "default":
		valDictA = listPairA[1][idx_a]
		valDictB = listPairB[1][idx_b]
		resVal   = []

		if flg_mode == 0:
			if valDictA == valDictB: resVal = valDictA
			elif isinstance(valDictA, dict) != isinstance(valDictB, dict): resVal = ""
			elif isinstance(valDictA, list) or isinstance(valDictB, list):
				if not isinstance(valDictA, list): valDictA = [valDictA]
				if not isinstance(valDictB, list): valDictB = [valDictB]

				if len(valDictA) < len(valDictB):
					listBase  = valDictA
					listCheck = valDictB
				else:
					listBase  = valDictB
					listCheck = valDictA

				for eachInfo in listBase:
					if eachInfo in listCheck: resVal.append(eachInfo)
			else: resVal = ""

		elif flg_mode == 1:
			if isinstance(valDictA, list): resVal.extend(valDictA)
			else: resVal.append(valDictA)

			if isinstance(valDictB, list): resVal.extend(valDictB)
			else: resVal.append(valDictB)
			
	if isinstance(resVal, list) and len(resVal) == 1: resVal = resVal[0]
	tmpList = strKey.split(chr_join)
	tmpList.append(resVal)
	resVal  = hs_create_dict_from_serial_list(tmpList)[1]

	#--- Release -----------------------------------------------------------------------------------
	return resVal

def IBase_operation_union_dicts(listDict):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Combine every dictionaries provided in 'listDict' with 'Union' operation

	[list] listDict, A list of dictionaries to be combined with union operation

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listDict, list):
		if isinstance(listDict, dict): return listDict
		else: return {}

	#*** Initialization ****************************************************************************
	chr_join = getconst_join_character()
	dictRes  = {}

	#*** Operations ********************************************************************************
	for eachDict in listDict:
		if len(dictRes.keys()) == 0: dictRes = eachDict.copy()
		else:
			listParam = IBase_transform_dict_to_list(dictRes)[0]
			listParam.extend(IBase_transform_dict_to_list(eachDict)[0])
			listParam = bbs_remove_dupl(listParam, True)
			listRes   = [[], []]

			for eachKey in listParam:
				listRes[0].append(eachKey.split(chr_join)[0])
				listRes[1].append(IBase_manipulate_dicts_target_key(dictRes, eachDict, eachKey, 1))

			listMainKey = listRes[0].copy()
			listMainKey = bbs_remove_dupl(listMainKey, True)
			listPrepVal = []

			for thisMainKey in listMainKey:
				listTmpVal = []
				
				for idx, eachVal in enumerate(listRes[1]):
					if listRes[0][idx] == thisMainKey: listTmpVal.append(eachVal)

				if len(listTmpVal) == 1: listPrepVal.append(listTmpVal[0])
				elif len(listTmpVal) > 1:
					combVal = []
					for eachTmpVal in listTmpVal:
						if not isinstance(eachTmpVal, list): eachTmpVal = [eachTmpVal]
						if len(combVal) == 0: combVal.extend(eachTmpVal)
						else:
							for cnt in range(0, len(eachTmpVal)):
								arr_idx = 0
								curKeys = IBase_get_keylist_of_dict_single_path(eachTmpVal[cnt])[0]
								tarDict = {}
								tarDict[curKeys[0]] = eachTmpVal[cnt][curKeys[0]]

								while IBase_validate_dict_key_exist(combVal[arr_idx], tarDict):
									arr_idx += 1
								combVal[arr_idx] = IBase_operation_union_dicts([combVal[arr_idx], tarDict])


					if len(combVal) == 1: combVal = combVal[0]
					listPrepVal.append(combVal)

			dictRes = {}
			for eachKey, eachValue in zip(listMainKey, listPrepVal): dictRes[eachKey] = eachValue

	#--- Release -----------------------------------------------------------------------------------
	return dictRes

def IBase_transform_list_to_dict(listData):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Transform formatted 'listData' List into Python dictionary object

	[list] listData, Target formatted Parameter-Value list to be transformed to dictionary

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listData, list): return {}	

	#*** Initialization ****************************************************************************
	chr_join  = getconst_join_character()
	listParam = listData[0]
	listValue = listData[1]
	dictRes   = {}

	#*** Operations ********************************************************************************
	#--- Post-Validation ---------------------------------------------------------------------------
	if not (isinstance(listParam, list) and isinstance(listValue, list)): return {}
	if len(listParam) != len(listValue): return {}

	#--- Transform conditioned data to dictionary --------------------------------------------------
	for keypath, arrVal in zip(listParam, listValue):
		listCurDicts = [dictRes]
		listKeys 	 = keypath.split(chr_join)

		for eachVal in arrVal:
			tmpDict = {}
			tmpKeys = listKeys.copy()
			tmpKeys.append(eachVal)
			tmpInfo = hs_create_dict_from_serial_list(tmpKeys)
			tmpDict[tmpInfo[0]] = tmpInfo[1]
			listCurDicts.append(tmpDict)
		dictRes = IBase_operation_union_dicts(listCurDicts)


	#--- Release -----------------------------------------------------------------------------------
	return dictRes

def IBase_transform_dict_to_list(dictData):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Transform Python dictionary object to Python Parameter-Value list

	[list] dictData, Target Python dictionary object

	'''

	#*** Input Validation **************************************************************************
	# Nothing to be pre-validated

	#*** Initialization ****************************************************************************
	chr_join = getconst_join_character()
	listVal  = [] 		# [dictInfoLvl0, dictInfoLvl1, ....], dictInfoLvlX = [listParam, listValue]

	#*** Operations ********************************************************************************
	#--- Post-Validation ---------------------------------------------------------------------------
	if not isinstance(dictData, dict): return [[], []]

	#--- Init Transformation parameters ------------------------------------------------------------
	cnt_dict = 0
	dictPrev = dictData.copy()
	listVal.append([[], []])

	for key in dictData.keys():
		thisValue = dictData[key]
		listVal[0][0].append(key)
		listVal[0][1].append(thisValue)
		if isinstance(thisValue, dict): cnt_dict += 1
		if isinstance(thisValue, list):
			for eachValue in thisValue:
				if isinstance(eachValue, dict): cnt_dict += 1
	if cnt_dict == 0: return listVal[0]

	#--- Transform target dictionary ---------------------------------------------------------------
	flg_work = True

	while flg_work:
		#--- Init current level --------------------------------------------------------------------
		cnt_dict = 0
		col_curr = len(listVal)
		listVal.append([[], []])

		#--- Flatten current level Parameter-Value ------------------------------------------------- 
		for idx, key in enumerate(listVal[col_curr - 1][0]):
			thisValue = listVal[col_curr - 1][1][idx]

			# Case: Not neither 'dict' nor 'list' type, can be appended immediately
			if not (isinstance(thisValue, dict) or isinstance(thisValue, list)):
				listVal[col_curr][0].append(key)
				listVal[col_curr][1].append(thisValue)
			
			# Case: 'dict' or 'list', treat in list way and use recursive method
			else:
				if isinstance(thisValue, dict): thisValue = [thisValue]
				for eachValue in thisValue:
					if isinstance(eachValue, dict): RetVal = IBase_transform_dict_to_list(eachValue)
					else: RetVal = [[key], [eachValue]]

					# Dict-left check
					for eachFlattenVal in RetVal[1]:
						if isinstance(eachFlattenVal, dict): cnt_dict += 1

					# Append sub-current 'RetVal' to Result list
					RetVal[0] = [chr_join.join([key, x]) for x in RetVal[0]]
					listVal[col_curr][0].extend(RetVal[0])
					listVal[col_curr][1].extend(RetVal[1])

		#--- Release current level -----------------------------------------------------------------
		if cnt_dict == 0: flg_work = False

	#--- Release -----------------------------------------------------------------------------------
	return hs_get_formatted_listpair(listVal[len(listVal) - 1], False, 0)

def IBase_transform_dict_to_json(dictData):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Transform Python dictionary object to JSON object

	[list] dictData, Prepared dictionary object

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(dictData, dict): dictData = {}

	#*** Initialization ****************************************************************************
	# Nothing to be initialized

	#*** Operations ********************************************************************************
	return json.dumps(dictData, indent= 4)

def IUser_transform_dataset_to_json(listData, flg_allstring, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Transform Python List/Dataset object to JSON object

	[list] listData, 	  Target flatten list to be transformed to JSON
	[bool] flg_allstring, True: Convert all values into String, False: no conversion
	[int]  flg_case,  	  0: Do nothing, 1: Convert all string to upper case, 2: ... to lower case

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listData, list): return json.dumps({})

	#*** Initialization ****************************************************************************
	# Nothing to be initialized

	#*** Operations ********************************************************************************
	#--- Format 'listData' to expected format ------------------------------------------------------
	listData = hs_get_formatted_listpair(listData, flg_allstring, flg_case)

	#--- Transform 'listData' into Python dictionary object ----------------------------------------
	dictData = IBase_transform_list_to_dict(listData)

	#--- Release -----------------------------------------------------------------------------------
	return IBase_transform_dict_to_json(dictData)


