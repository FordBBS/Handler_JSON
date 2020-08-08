####################################################################################################
#                                                                                                  #
# Python, JSON handler, BBS					   						                               #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/08/08, BBS:	- First release
# 					- Implemented Data Conditioning, Importing, Exporting function groups
#
#***************************************************************************************************

#*** Function Group List ***************************************************************************
# - Helper
# - Importing
# - Exporting
# - Data Conditioning



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
		else: dictRes[mainKey] = tmpDict

	#--- Case: only 2 elements ---------------------------------------------------------------------
	else: dictRes[mainKey] = thisVal

	#--- Release -----------------------------------------------------------------------------------
	return dictRes



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




#*** Function Group: Data Conditioning *************************************************************
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
	dictResult = {}
	listPrep   = [[], []]

	#*** Operations ********************************************************************************
	#--- Post-Validation ---------------------------------------------------------------------------
	if not (isinstance(listKey, list) and isinstance(listVal, list)): return {}
	if len(listKey) != len(listVal): return {}

	#--- Data Conditioning -------------------------------------------------------------------------
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
		branch  = len(thisVal)

		# for pos, eachVal in enumerate(thisVal):
		# 	curDict = dictResult
			
		# 	for eachKey in thisKey:
		# 		if isinstance(curDict, list): curDict = curDict[pos]
		# 		if not eachKey in curDict:
		# 			if eachKey == thisKey[0] and branch > 1:

		# 			else:
		# 		else: curDict = curDict[eachKey]



	#--- Release -----------------------------------------------------------------------------------
	return dictResult

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

RetVal = IUser_transform_dataset_to_json(listParam, False, 0)
print(RetVal)