####################################################################################################
#                                                                                                  #
# Python, JSON handler, BBS					   						                               #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/08/08, BBS:	- First release
# 					- Implemented Data Manipulation, Importing, Exporting function groups
# 2020/08/09, BBS: 	- Implemented following functions completely
# 						'IBase_update_dict_by_path', 'IBase_merge_dict_single_path', 
# 						'IBase_get_keylist_of_dict_single_path'
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

# Dataset handler
import  dataset_handler 	as hs_dataset



#*** Function Group: Helper ************************************************************************
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
		tmpDict = {}
		tmpDict[listInput[-2]] = thisVal

		for idx in range(len(listInput) - 3, 0, -1):
			curKey  = listInput[idx]
			curDict = tmpDict.copy()
			tmpDict = {}
			tmpDict[curKey] = curDict
		listRes = [mainKey, tmpDict]

	#--- Case: only 2 elements ---------------------------------------------------------------------
	else: listRes = [mainKey, thisVal]

	#--- Release -----------------------------------------------------------------------------------
	return listRes



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

def IBase_update_dict_by_path(mainDict, listKey, tarValue, flg_mode):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Update 'mainDict' on requested path, breakdown element in 'listKey', by 'tarValue'

	[dict] mainDict, Main dictionary
	[list] listKey,  List of keys to be used as a path guideline
	[any]  tarValue, Target value to be updated
	[int]  flg_mode, 0: Change, 1: Append, 2: Remove

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(mainDict, dict): return mainDict
	if not (isinstance(listKey, list) or isinstance(listKey, str)) or len(listKey) == 0:
		return mainDict

	#*** Initialization ****************************************************************************
	if not (isinstance(flg_mode, int) and flg_mode in range(0, 3)): flg_mode = 0

	listKey   = hs_dataset.hs_prep_StrList(listKey)
	dictRes   = mainDict.copy()
	tarDict   = {}
	listFound = []

	#*** Operations ********************************************************************************
	#--- Get a list of available key ---------------------------------------------------------------
	listFound = IBase_validate_dict_key_exist(dictRes, listKey, 1)

	#--- Create target dictionary to be merged from the rest that weren't found --------------------
	for idx_pos in range(len(listKey) - 1, -1, -1):
		thisKey = listKey[idx_pos]

		if not thisKey in listFound:
			if tarDict == {}: tarDict[thisKey] = tarValue
			else: tarDict[thisKey] = tarDict
	if tarDict == {}: tarDict = tarValue

	#--- WorkMode ----------------------------------------------------------------------------------
	tmpDictValue = [listFound[0]]
	tmpDict 	 = dictRes.copy()

	# Fetch dictionary value on each level to 'tmpDictValue'
	for idx, key in enumerate(listFound):
		tmpDict = tmpDict[key]
		tmpDictValue.append(tmpDict)

	# Manipulate 'tmpDictValue', it will be only 2 indices in the end which are Key and Value
	if flg_mode == 0: tmpDictValue[1] = tarDict
	elif flg_mode == 1:
		idx_start   = len(tmpDictValue) - 1
		flg_set_val = False

		if len(listFound) < len(listKey):
			curValue = tmpDictValue[len(tmpDictValue) - 1]
			curValue.update(tarDict)
			tmpDictValue[len(tmpDictValue) - 1] = curValue
			flg_set_val  = True
			idx_start	 = idx_start - 1

		else: tarValue = hs_dataset.hs_prep_AnyList(tarValue)
		
		for cnt in range(idx_start, 0, -1):
			if not flg_set_val:
				curValue = tmpDictValue[cnt]

				if not isinstance(curValue, list): curValue = [curValue]
				else: curValue = curValue.copy()
			
				curValue.extend(tarValue)
				flg_set_val = True
			
			else:
				curValue = tmpDictValue[cnt]
				curValue[listFound[cnt]] = tmpDictValue[cnt + 1]
			
			tmpDictValue[cnt] = curValue

	# Manipulate final dictionary based on 'flg_mode'
	if flg_mode < 2: dictRes[tmpDictValue[0]] = tmpDictValue[1]
	elif len(tmpDictValue) > 0: del dictRes[tmpDictValue[0]]
	
	#--- Release -----------------------------------------------------------------------------------
	return dictRes

def IBase_validate_dict_key_exist(mainDict, listKey, flg_mode):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Checks provided keys in 'listKey' (which are in sequence as a path) whether this path does
		exist in 'mainDict'. 'flg_mode' determines the type of result

		Result Format: flg_mode = 0
		True:  'listKey' path does exist in 'mainDict'
		False: 'listKey' path doesn't exist in 'mainDict'

		Result Format: flg_mode = 1
		Return deepest key found in 'mainDict', "" if nothing

	[dict] mainDict, Main dictionary
	[list] listKey,  List of keys to be used as a path guideline
	[int]  flg_mode, 0: True or False, 1: Return deepest key found in 'mainDict', "" if nothing

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(mainDict, dict): return False

	#*** Initialization ****************************************************************************
	if not (isinstance(flg_mode, int) and flg_mode in range(0, 2)): flg_mode = 0
	listKey   = hs_dataset.hs_prep_StrList(listKey)
	listFound = []
	tmpDict   = mainDict.copy()

	#*** Operations ********************************************************************************
	#--- Key Existence checking --------------------------------------------------------------------
	for key in listKey:
		if key in tmpDict:
			tmpDict = tmpDict[key]
			listFound.append(key)

			if not isinstance(tmpDict, dict): break
		else: break

	#--- Result translation ------------------------------------------------------------------------
	if flg_mode == 0:
		if listFound[len(listFound) - 1] == listKey[len(listKey) - 1]: return True
		else: return False
	else: return listFound

def IBase_merge_dict_single_path(mainDict, listDict):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Merge each dictionary in 'listDict' to 'mainDict'
		Each dictionary in 'listDict' must be a single path dictionary. Otherwise, only the first
		path will be merged.

	[dict] mainDict, Main dictionary
	[list] listDict, List of single path dictionary to be merged to Main dictionary

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listDict, list): return {}

	#*** Initialization ****************************************************************************
	dictRes = mainDict.copy()

	#*** Operations ********************************************************************************
	for cnt, thisDict in enumerate(listDict):
		if cnt == 0: dictRes.update(thisDict)
		else:
			listKey = IBase_get_keylist_of_dict_single_path(thisDict)
			dictRes = IBase_update_dict_by_path(dictRes, listKey[0], listKey[1], 1)

	#--- Release -----------------------------------------------------------------------------------
	return dictRes

def IBase_transform_list_to_dict(listData, flg_allstring, flg_case):
	#*** Documentation *****************************************************************************
	'''Documentation 

		Transform 'listData' List into Python dictionary object

	[list] listData, 	  Target flatten list to be transformed to dictionary
	[bool] flg_allstring, True: Convert all values into String, False: no conversion
	[int]  flg_case,  	  0: Do nothing, 1: Convert all string to upper case, 2: ... to lower case

	'''

	#*** Input Validation **************************************************************************
	if not isinstance(listData, list): return {}	

	#*** Initialization ****************************************************************************
	if isinstance(flg_allstring, int):
		if flg_allstring != 0: flg_allstring = True
		else: flg_allstring = False
	
	if not isinstance(flg_allstring, bool): flg_allstring = False
	if not (isinstance(flg_case, int) and flg_case in range(0, 3)): flg_case = 0

	listKey    = listData[0]
	listVal    = listData[1]
	listPrep   = [[], []]
	dictResult = {}

	#*** Operations ********************************************************************************
	#--- Post-Validation ---------------------------------------------------------------------------
	if not (isinstance(listKey, list) and isinstance(listVal, list)): return {}
	if len(listKey) != len(listVal): return {}

	#--- Flatten 'listData' ------------------------------------------------------------------------
	for thisKey, thisVal in zip(listKey, listVal):
		if flg_allstring: thisVal = str(thisVal)
		if flg_case == 1 and isinstance(thisVal, str): thisVal = thisVal.upper()
		if flg_case == 2 and isinstance(thisVal, str): thisVal = thisVal.lower()

		if not thisKey in listPrep[0]:
			listPrep[0].append(thisKey)
			listPrep[1].append([thisVal])
	
		else:
			idx_key = listPrep[0].index(thisKey)
			listPrep[1][idx_key].append(thisVal)

	#--- Transform conditioned data to dictionary --------------------------------------------------
	for idx in range(0, len(listPrep[0])):
		thisKey = listPrep[0][idx].split(".")
		thisVal = listPrep[1][idx]
		
	#--- Release -----------------------------------------------------------------------------------
	return dictResult

def IBase_transform_dict_to_list(dictData):
	print("TODO")

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
	return json.dumps(dictData)

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
	#--- Transform 'listData' into Python dictionary object ----------------------------------------
	RetVal = IBase_transform_list_to_dict(listData, flg_allstring, flg_case)

	#--- Transform Python dictionary object to JSON ------------------------------------------------
	return IBase_transform_dict_to_json(RetVal)



# Debuging Area
listKey = ["nameFlowstream", "nameFlowstream", "ecs.devicename", "ecs.activate", \
 			"gmd.devicename", "gmd.activate", "gmd.table.delaytime", "gmd.table.range", \
 			"gmd.devicename", "gmd.activate", "gmd.table.delaytime", "gmd.table.range"]

listVal = ["B: Gas SULEV/Bag", "TestDouble", "CVS", "Yes", \
 			"SULEV", "Yes", "no_delay", "Modal", \
 			"CONTBAG", "Yes", "Modal", "Modal;Bag"]

listParam = [listKey, listVal]

#RetVal = IUser_transform_dataset_to_json(listParam, False, 0)
#print(RetVal)

dict1 = {"Engineer": {"Name": ["Ford_BBS", "Daniel"]}}
dict2 = {"Engineer": {"Name": "Patrick"}}
dict3 = {"Engineer": {"Skills": "iGEM"}}
dict4 = {"Engineer": {"Skills": "PUMA"}}
print(IBase_merge_dict_single_path({}, [dict1, dict2, dict3, dict4]))