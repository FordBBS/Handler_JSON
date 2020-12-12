'*** History ***************************************************************************************
' 2020/08/27, BBS:	- First Release
' 					- Imported all mandatory materials
' 2020/09/19, BBS:	- Updated 'hs_arr_append'
' 2020/09/21, BBS: 	- Updated 'IUser_translate_json_strContent'
' 2020/10/07, BBS:	- Updated 'IUser_translate_json_strContent'
' 					- Implemented 'IUser_read_json_from_file'
' 2020/12/11, BBS:	- Updated 'IBase_create_resParamValue', 'IUser_translate_json_strContent'
'					- Integrated 'hs_arr_val_exist_ex', 'hs_arr_slice'
'
'***************************************************************************************************

'*** Imported Materials ****************************************************************************
'--- Documentation ---------------------------------------------------------------------------------
' (Version 2020/08/23) IUser_get_value_of_param
' (Version 2020/08/23) IUser_clean_ParamPath
' (Version 2020/12/11) IUser_translate_json_strContent
' (Version 2020/08/27) IUser_translate_resParamValue
' (Version 2020/12/11) IBase_create_resParamValue
' (Version 2020/08/26) IBase_getinfo_resParamValue
' (Version 2020/08/23) IBase_get_value_from_strLine
' (Version 2020/08/27) hs_read_text_file
' (Version 2020/09/19) hs_arr_append
' (Version 2020/08/25) hs_arr_stack
' (Version 2020/12/11) hs_arr_slice
' (Version 2020/12/06) hs_arr_val_exist_ex
' (Version 2020/12/12) hs_parser_remove_redundant_params
'
'---------------------------------------------------------------------------------------------------

Function IUser_get_value_of_param(ByVal arrParamValue, ByVal strTargetParam, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return the value of target Parameter path
	' 	Return "NotExist" if target Parameter path doesn't exist
	'	
	'	Argument(s)
	'	<Array> 	arrParamValue,  Parameter-Value Array
	'	<String> 	strTargetParam, Target parameter path
	' 	<Long> 		flg_case, 		0: Character's case doesn't matter, 1: Vice versa
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_get_value_of_param = "NotExist"

	'*** Pre-Validation ****************************************************************************
	If InStr(LCase(TypeName(arrParamValue)), "variant") = 0 Then Exit Function
	If UBound(arrParamValue) < 1 Then Exit Function
	If len(CStr(strTargetParam)) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, arrParam, arrValue, thisParam

	strTargetParam = CStr(strTargetParam)
	arrParam 	   = arrParamValue(0)
	arrValue 	   = arrParamValue(1)

	If LCase(TypeName(flg_case)) <> "integer" Then flg_case = 1
	If flg_case < 0 or flg_case > 1 Then flg_case = 1
	If flg_case = 0 Then strTargetParam = LCase(strTargetParam)

	'*** Operations ********************************************************************************
	For cnt1 = 0 to UBound(arrParam)
		thisParam = arrParam(cnt1)
		
		If flg_case = 0 Then thisParam = LCase(thisParam)
		If strTargetParam = thisParam Then
			IUser_get_value_of_param = arrValue(cnt1)
			Exit For
		End If
	Next

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IUser_clean_ParamPath(ByVal arrParamPath, ByVal arrReplaceTag)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	' 					- Improved handler of 'arrReplaceTag', If only one replace guide is needed,
	' 					it can be provided as Array("xxxxxx", "yyyyy")
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Replace target Parameter tag in 'Parameter Path' with specific string
	' 	e.g. "config_params.ecs" -> "ecs"
	' 		 "gmd.table.range" 	 -> "gmd.analyzerrange"
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_clean_ParamPath = arrParamPath

	'*** Pre-Validation ****************************************************************************
	If InStr(LCase(TypeName(arrReplaceTag)), "variant") = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt2, thisParam, thisReplaceGuide, flg_not_arr

	flg_not_arr = 0

	'*** Operations ********************************************************************************
	'--- Prepare 'arrReplaceTag' -------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrReplaceTag)
		If InStr(LCase(TypeName(arrReplaceTag(cnt1))), "variant") = 0 Then
			flg_not_arr = flg_not_arr + 1
		End If
	Next

	If UBound(arrReplaceTag) = 1 and flg_not_arr > 0 Then
		arrReplaceTag = Array(arrReplaceTag)
	End If

	'--- Cleaning ----------------------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrParamPath)
		thisParam = "." & arrParamPath(cnt1) & "."

		For cnt2 = 0 to UBound(arrReplaceTag)
			thisReplaceGuide = arrReplaceTag(cnt2)

			If InStr(LCase(TypeName(thisReplaceGuide)), "variant") > 0 Then
				If InStr(thisParam, "." & thisReplaceGuide(0) & ".") > 0 Then
					thisParam = Replace(thisParam, thisReplaceGuide(0), thisReplaceGuide(1))
				End If
			End If
		Next

		thisParam = Replace(thisParam, "..", ".")
		
		If Left(thisParam, 1) = "." Then thisParam = Mid(thisParam, 2, len(thisParam))
		If Right(thisParam, 1) = "." Then thisParam = Left(thisParam, len(thisParam) - 1)

		arrParamPath(cnt1) = thisParam
	Next

	IUser_clean_ParamPath = arrParamPath

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IUser_translate_json_strContent(ByVal strContent)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	' 2020/08/27, BBS:	- Bug fixed
	' 					1) Array tag is not removed when it's only one tag left
	'					2) Parameter path ends with "." for any value that has array value line
	' 2020/09/21, BBS:	- Bug fixed, Array branching is not closed correctly
	' 2020/10/07, BBS:	- Bug fixed, Empty parameter is not translated correctly
	' 2020/12/11, BBS:	- Implemented conditioner before release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return Parameters-Values array of provided JSON content in 'strContent' String type
	'	e.g. (("ecs.devicename", "ecs.activate"), ("CVS_SL", "Yes"))
	' 		 (("gmd.devicename"), ("%tag%0%tag%SULEV%;%%tag%1%tag%CONT_BAG"))
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_translate_json_strContent = Array(Array(), Array())

	'*** Pre-Validation ****************************************************************************
	If TypeName(strContent) <> "String" Then Exit Function
	If len(strContent) < 3 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt_row, flg_append, tagBranch, tagLabel, tagLatest, thisParam, thisValue, existValue
	Dim curRoot, curPath, strTagArray, strTagValue, strTagRemove, strParamEx, strBrnTag, strBrnIdx
	Dim flg_clr_rt, arrContent, arrThisInfo, arrRoot, arrBrnIdx, arrParam(), arrValue()
	Redim Preserve arrParam(0), arrValue(0)

	arrContent  = Split(strContent, vbCrLf)
	cnt_row     = 0
	tagBranch   = "%tag%"
	curRoot 	= ""
	strTagArray = ""				' Storage: Parameter that has branch
	strTagValue = "" 				' Storage: Parameter that has value on its line
	strParamEx  = "" 				' Storage: Appended Parameter path
	strBrnTag   = "" 				' Storage: Branch, Parameter owner of each position index
	strBrnIdx   = "" 				' Storage: Branch, Position index
	
	'*** Operations ********************************************************************************
	'--- Parsing -----------------------------------------------------------------------------------
	while cnt_row < UBound(arrContent)
		arrThisInfo = IBase_get_value_from_strLine(arrContent(cnt_row), ":")
		thisParam   = arrThisInfo(0)
		thisValue	= CStr(arrThisInfo(1))
		flg_append	= False
		flg_clr_rt  = False

		If thisParam <> "" Then 	' Case: Parameter does exist
			' Case: Empty Parameter, No value, No SubGroup
			If InStr(arrContent(cnt_row), "{}") > 0 or InStr(arrContent(cnt_row), "[]") > 0 Then
				flg_append = True

			' Case: Value doesn't exist but SubGroup's or Array's symbol
			ElseIf InStr(arrContent(cnt_row), "{") > 0 or InStr(arrContent(cnt_row), "[") > 0 Then
				If curRoot <> "" Then
					curRoot = Join(Array(curRoot, thisParam), ".")
				Else
					curRoot = thisParam
				End If
				
				If strTagValue <> "" Then
					strTagValue = Join(Array(strTagValue, "%" & thisParam & "%"), ";")
				Else
					strTagValue = "%" & thisParam & "%"
				End If
				
				If InStr(arrContent(cnt_row), "[") > 0 Then
					If strTagArray <> "" Then
						strTagArray = Join(Array(strTagArray, "%" & thisParam & "%"), ";")
					Else
						strTagArray = "%" & thisParam & "%"
					End If

					If strBrnTag = "" Then
						strBrnTag = "%" & thisParam & "%"
						strBrnIdx = "0"
					Else
						strBrnTag = Join(Array(strBrnTag, "%" & thisParam & "%"), ";")
						strBrnIdx = Join(Array(strBrnIdx, "0"), ";")
					End If
				End If
			
			' Case: Value does exist
			Else
				flg_append = True
			End If
		ElseIf thisParam = "" Then 	' Case: Parameter doesn't exist
			' Case: End of current sub-tags group '{}', clear Root, and stored tags string
			If InStr(arrContent(cnt_row), "}") > 0 and len(curRoot) > 0 Then
				flg_clr_rt = True

			' Case: End of latest branch, clear all memo info of latest branch
			ElseIf InStr(arrContent(cnt_row), "]") > 0 Then
				arrRoot 	 = Split(curRoot, ".")
				arrBrn 		 = Split(strTagArray, ";")
				strTagRemove = Replace(arrBrn(UBound(arrBrn)), "%", "")

				If arrRoot(UBound(arrRoot)) = strTagRemove Then
					flg_clr_rt = True
				End If
	
				If InStrRev(strTagArray, ";") > 0 Then
					strTagArray = Mid(strTagArray, 1, InStrRev(strTagArray, ";") - 1)
				Else
					strTagArray = ""
				End If

				If InStrRev(strBrnTag, ";") > 0 Then
					strBrnTag = Mid(strBrnTag, 1, InStrRev(strBrnTag, ";") - 1)
					strBrnIdx = Mid(strBrnIdx, 1, InStrRev(strBrnIdx, ";") - 1)
				Else
					strBrnTag = ""
					strBrnIdx = ""
				End If

			' Case: Value line (e.g. Parameter that has array value will break its value into lines)
			ElseIf InStr(arrContent(cnt_row), "{") = 0 and InStr(arrContent(cnt_row), "[") = 0 Then
				flg_append  = True
			End If
		End If

		' Root Removal
		If flg_clr_rt Then
			strTagRemove = Mid(curRoot, InStrRev(curRoot, ".") + 1, len(curRoot))

			If InStr(strTagArray, "%" & strTagRemove & "%") = 0 and _
			 	InStr(strTagValue, "%" & strTagRemove & "%") > 0 Then

			 	If InStr(curRoot, ".") > 0 Then
			 		curRoot = Mid(curRoot, 1, InStrRev(curRoot, ".") - 1)
			 	Else
			 		curRoot = ""
			 	End If

			 	If InStr(strTagValue, ";") > 0 Then
			 		strTagValue = Mid(strTagValue, 1, InStrRev(strTagValue, ";") - 1)
			 	Else
			 		strTagValue = ""
			 	End If
			End If
		End If

		' Appending
		If flg_append Then
			If thisParam = "" Then
				curPath = curRoot
			ElseIf curRoot <> "" Then
				curPath = Join(Array(curRoot, thisParam), ".")
			Else
				curPath = thisParam
			End If

			' Case: Current Parameter path already exist
			If InStr(strParamEx, "%" & curPath & "%") > 0 Then
				' Get index of this Parameter path and existing value
				For cnt1 = 0 to UBound(arrParam)
					If arrParam(cnt1) = curPath Then
						existValue = arrValue(cnt1)
						Exit For
					End If
				Next

				'Appending - Branch check then append
				If InStrRev(existValue, "%;%") > 0 Then
					tagLatest = Mid(existValue, InStrRev(existValue, "%;%") + 3, len(existValue))
				Else
					tagLatest = existValue
				End If
				
				tagLatest = Mid(tagLatest, len(tagBranch) + 1, _
					 						InStrRev(tagLatest, tagBranch) - len(tagBranch) - 1)

				If tagLatest = strBrnIdx Then
					arrBrnIdx = Split(strBrnIdx, ";")
					arrBrnIdx(UBound(arrBrnIdx)) = CStr(CInt(arrBrnIdx(UBound(arrBrnIdx))) + 1)
					strBrnIdx = Join(arrBrnIdx, ";")
				End If

				arrValue(cnt1) = existValue & "%;%" & tagBranch & strBrnIdx & tagBranch & thisValue
				
			' Case: Current Parameter path has its first time appending
			Else
				' Store current parameter path
				If strParamEx <> "" Then
					strParamEx = Join(Array(strParamEx, "%" & curPath & "%"), ";")
				Else
					strParamEx = "%" & curPath & "%"
				End If

				' Prepare proper size for result arrays
				If Not (UBound(arrParam) = 0 and len(arrParam(0)) = 0) Then
					Redim Preserve arrParam(UBound(arrParam) + 1), arrValue(UBound(arrValue) + 1)
				End If

				' Create branch for 'thisValue' if it's necessary
				If strBrnIdx <> "" Then thisValue = tagBranch & strBrnIdx & tagBranch & thisValue	

				' Append Parameter path and its Value
				arrParam(UBound(arrParam)) = curPath
				arrValue(UBound(arrValue)) = thisValue
			End If
		End If

		' Release current line
		If cnt_row < 0 Then
			cnt_row = UBound(arrContent)
		Else
			cnt_row = cnt_row + 1
		End If
	wend
	
	'--- Translate 'arrValue' ----------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrValue)
		If InStr(arrValue(cnt1), "%tag%") > 0 Then
			existValue 	   = IUser_translate_resParamValue(arrValue(cnt1), "")
			arrValue(cnt1) = existValue
		End If
	Next

	'--- Result's Conditioning ---------------------------------------------------------------------
	Call hs_parser_remove_redundant_params(arrParam, arrValue)
	
	'--- Release -----------------------------------------------------------------------------------
	IUser_translate_json_strContent = Array(arrParam, arrValue)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IUser_translate_resParamValue(ByVal strValue, ByVal strIdxList)
	'*** History ***********************************************************************************
	' 2020/08/27, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return translated value of 'strValue' based on requested index in 'strIdxList'
	' 	e.g.  strValue 	   = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%1;0%tag%SULEV"
	'		  strIdxList = "0"	-> return ("SULEV", "CONT_BAG")
	' 		  strIdxList = "1" 	-> return ("SULEV")
	'' 		  strIdxList = "0;1"  -> return ("CONT_BAG")
	'	
	'	Argument(s)
	'	<String> strValue,   A string value created by 'IUser_translate_json_strContent'
	'	<String> strIdxList, A string list of desire index/indices
	'
	'***********************************************************************************************

	On Error Resume Next
	IUser_translate_resParamValue = IBase_create_resParamValue(strValue)

	'*** Pre-Validation ****************************************************************************
	strValue   = CStr(strValue)
	strIdxList = CStr(strIdxList)
	If Not (len(strValue) > 0 and len(strIdxList) > 0) Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, tagBrnIdx, tmpValue
	Dim arrRes, arrValue, arrTarIdx, arrSplit

	tagBrnIdx = "%tag%"
	arrSplit  = Split(strValue, "%;%")
	arrTarIdx = Split(strIdxList, ";")
	arrValue  = IBase_create_resParamValue(strValue)

	'*** Operations ********************************************************************************
	'--- Post-Validation ---------------------------------------------------------------------------
	If UBound(arrSplit) < 0 Then Exit Function

	'--- Collect data based on target index list ---------------------------------------------------
	For cnt1 = 0 to UBound(arrTarIdx)
		If UBound(arrValue) < CInt(arrTarIdx(cnt1)) Then
			Exit Function
		Else
			tmpValue = arrValue(CInt(arrTarIdx(cnt1)))
			arrValue = tmpValue
		End If
	Next
	
	'--- Release -----------------------------------------------------------------------------------
	If InStr(LCase(TypeName(arrValue)), "variant") > 0 Then
		IUser_translate_resParamValue = arrValue
	Else
		IUser_translate_resParamValue = Array(arrValue)
	End If

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IBase_create_resParamValue(ByVal strValue)
	'*** History ***********************************************************************************
	' 2020/08/26, BBS:	- First Release
	' 2020/08/27, BBS:	- Bug fixed, when 'strValue' has only one level
	'					- Bug fixed, invalid If-Else condition for creating result for new branch
	' 2020/12/11, BBS:	- Overhaul mechanism
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' 	Return a general form Value of 'strValue'
	' 	e.g.  strValue = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%1;0%tag%CONT_BAG_THC"
	' 		  Return (("SULEV", "CONT_BAG"), ("CONT_BAG_THC"))
	'
	' 		  strValue = "%tag%0;0;0%tag%Modal%;%%tag%0;0;1%tag%Bag%;%%tag%0;1;0%tag%THC"
	' 		  Return ((("Modal", "Bag"), ("THC")))
	'
	'		  strValue = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%2;0%tag%CONT_BAG_THC"
	' 		  Return (("SULEV", "CONT_BAG"), (), ("CONT_BAG_THC"))
	'	
	'	Argument(s)
	'	<String> strValue, A string value created by 'IUser_translate_json_strContent'
	'
	'***********************************************************************************************

	On Error Resume Next
	IBase_create_resParamValue = Array()

	'*** Pre-Validation ****************************************************************************
	strValue = CStr(strValue)
	If Len(strValue) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt2, cnt3, stack_lvl, corr_lvl, thisInfo, thisValue, flg_create
	Dim arrRet, arrValue, arrTagIdx, arrPrep, arrTmp, arrSnapIdx, arrSnapObj, arrSnapTmp
	Dim thisTag, tarSnapObj, tarSnapIdx
	
	arrValue = Split(strValue, "%;%")
	arrRet 	 = Array()

	'*** Operations ********************************************************************************
	For cnt1 = 0 to UBound(arrValue)
		thisInfo   = IBase_getinfo_resParamValue(arrValue(cnt1))
		thisValue  = thisInfo(0) 
		arrTagIdx  = Split(thisInfo(1), ";")
		flg_create = False
		flg_snap   = False
		arrSnapIdx = Array()
		arrSnapObj = Array()

		'--- Analysis of 'thisValue' information ---------------------------------------------------
		' Case: First value, skips snapshot analysis
		If cnt1 = 0 Then
			flg_create 	= True
			stack_lvl	= UBound(arrTagIdx) - 1		'Set needed stack level
			corr_lvl 	= 0							'Set start level of correction
		
		' Case: General, performs snapshot analysis
		Else
			arrTmp = arrRet

			For cnt2 = 0 to UBound(arrTagIdx)
				thisTag = CInt(arrTagIdx(cnt2))
				Call hs_arr_append(arrSnapIdx, thisTag)

				' Case: Result array has no target position yet
				If UBound(arrTmp) < thisTag Then
					stack_lvl  = UBound(arrTagIdx) - cnt2 - 1	'Set needed stack level
					corr_lvl   = cnt2 							'Set start level of correction
					flg_create = True
					Exit For

				' Case: Result array covers target position
				Else
					Call hs_arr_append(arrSnapObj, arrTmp) 	'Snapshot

					' Prepared next iteration
					arrSnapTmp = arrTmp(thisTag)
					arrTmp     = arrSnapTmp
				End If
			Next
		End If

		'--- Create new base array for 'thisValue' -------------------------------------------------
		If flg_create Then
			If UBound(arrTagIdx) > 0 and UBound(arrSnapObj) < 0 Then
				arrPrep = Array()

				For cnt2 = corr_lvl to (CInt(arrTagIdx(UBound(arrTagIdx))) - 1)
					Call hs_arr_append(arrPrep, "")
				Next
				
				Call hs_arr_append(arrPrep, thisValue)
				Call hs_arr_stack(arrPrep, stack_lvl)
			Else
				arrPrep = thisValue
			End If
		End If

		'--- Manipulate return array ---------------------------------------------------------------
		' Method: Snapshot Restoration
		If UBound(arrSnapObj) >= 0 Then
			Call hs_arr_append(arrSnapObj, arrPrep) 		'Append prepared value as last Snapshot

			For cnt2 = UBound(arrSnapObj) to 1 Step -1 		'Snapshot Restoration process
				tarSnapObj = arrSnapObj(cnt2 - 1)
				tarSnapIdx = arrSnapIdx(cnt2 - 1)
				arrTmp 	   = tarSnapObj(tarSnapIdx)

				For cnt3 = UBound(arrTmp) to (CInt(arrSnapIdx(cnt2)) - 2)
					Call hs_arr_append(arrTmp, "")
				Next
				Call hs_arr_append(arrTmp, arrSnapObj(cnt2))

				tarSnapObj(tarSnapIdx) = arrTmp
				arrSnapObj(cnt2 - 1)   = tarSnapObj
			Next

			arrRet = arrSnapObj(0) 			'Set top level of Snapshot as current Result

		' Method: Direct Appending
		Else
			For cnt2 = UBound(arrRet) to (CInt(arrTagIdx(0)) - 2)
				Call hs_arr_append(arrRet, "")
			Next
			Call hs_arr_append(arrRet, arrPrep)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	IBase_create_resParamValue = arrRet

	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IBase_getinfo_resParamValue(ByVal strValue)
	'*** History ***********************************************************************************
	' 2020/08/26, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	'	Return Value and TagIndex information of 'strValue'
	'	e.g. strValue = "%tag%0;1;4%tag%SULEV"
	'		 return ("SULEV", "0;1;4")
	'
	'	Argument(s)
	'	<String> strValue, A string of single value created by 'IUser_translate_json_strContent'
	' 					   If more than one value exist, only first value will be manipulated
	'
	'***********************************************************************************************

	On Error Resume Next
	IBase_getinfo_resParamValue = Array("", "")

	'*** Pre-Validation ****************************************************************************
	strValue   = CStr(strValue)
	If len(strValue) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim arrValue, tagBrnIdx, thisValue, thisTagIdx

	tagBrnIdx  = "%tag%"
	arrValue   = Split(strValue, "%;%")

	'*** Operations ********************************************************************************
	If InStr(arrValue(0), tagBrnIdx) > 0 Then 
		thisValue  = Mid(arrValue(0), len(tagBrnIdx) + 1, len(arrValue(0)))
		thisTagIdx = Mid(thisValue, 1, InStr(thisValue, tagBrnIdx) - 1)
		thisValue  = Mid(thisValue, InStr(thisValue, tagBrnIdx) + len(tagBrnIdx), len(thisValue))
		IBase_getinfo_resParamValue = Array(thisValue, thisTagIdx)
	End If

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IBase_get_value_from_strLine(ByVal strLine, ByVal chr_sep)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	' 2020/08/27, BBS: 	- Bug fixed, Case: Value only (e.g. strLine = ""MASS",")
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return array of Parameter or Value available on 'strLine' with JSON structure
	'	e.g. strLine = "testcell_name": "LD06", 	-> ("testcell_name", "LD06")
	' 		 strLine = "config_params": [ 			-> ("config_params")
	' 		 strLine = }, 							-> ("")
	' 		 strLine = "Test" 						-> ("", "Test")
	'
	'	Argument(s)
	'	<String> strLine, A String of content line to be parsed
	'	<String> chr_sep, A character used to separate between Parameter and Value
	'						e.g. chr_sep = "=" for XML, chr_sep = ":" for JSON
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IBase_get_value_from_strLine = Array("", "")

	'*** Pre-Validation ****************************************************************************
	If LCase(TypeName(strLine)) <> "string" Then Exit Function
	If len(strLine) < 1 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim flg_bln, flg_sum, cnt_idx, cnt_pos, curValue, arrChrNotVal(3), arrValue(1)

	If LCase(TypeName(chr_sep)) <> "string" Then chr_sep = ":"
	If len(chr_sep) = 0 Then chr_sep = ":"

	arrChrNotVal(0) = "{"
	arrChrNotVal(1) = "}"
	arrChrNotVal(2) = "["
	arrChrNotVal(3) = "]"
	arrValue(0) 	= ""
	arrValue(1) 	= ""

	'*** Operations ********************************************************************************
	'--- Clear spaces and tabs on left and right sides ---------------------------------------------
	strLine = Trim(strLine)
	cnt_idx = 1
	flg_bln = True
	
	while flg_bln
		flg_sum = 0

		If Left(strLine, 1) = vbTab or Left(strLine, 1) = " " Then
			strLine = Mid(strLine, cnt_idx + 1, len(strLine))
		Else
			flg_sum = flg_sum + 1
		End If

		If Right(strLine, 1) = vbTab or Right(strLine, 1) = " " or Right(strLine, 1) = "," Then
			strLine = Mid(strLine, 1, len(strLine) - 1)
		Else
			flg_sum = flg_sum + 1
		End If
		
		If len(strLine) = 0 or flg_sum = 2 Then flg_bln = False
	wend

	'--- Check Exist -------------------------------------------------------------------------------
	cnt_idx = Instr(strLine, chr(34))

	If cnt_idx > 0 Then
		cnt_pos = InStr(cnt_idx + 1, strLine, chr(34))

		If cnt_pos > 0 Then
			arrValue(0) = Mid(strLine, cnt_idx + 1, cnt_pos - cnt_idx - 1)
			cnt_idx 	= InStr(cnt_pos, strLine, chr_sep) + 1
			
			If cnt_idx = 1 Then
				arrValue(1) = arrValue(0)
				arrValue(0) = ""

			ElseIf InStr(cnt_idx, strLine, chr(34)) > 0 Then
				cnt_idx = InStr(cnt_idx, strLine, chr(34)) + 1
				cnt_pos = InStr(cnt_idx, strLine, chr(34))
				arrValue(1) = Mid(strLine, cnt_idx, cnt_pos - cnt_idx)
			
			Else
				flg_bln = True

				while flg_bln
					If Mid(strLine, cnt_idx, 1) <> " " and Mid(strLine, cnt_idx, 1) <> vbTab Then
						flg_bln 	= False
						arrValue(1) = Mid(strLine, cnt_idx, len(strLine))
						
						For cnt_idx = 0 to UBound(arrChrNotVal)
							If InStr(arrValue(1), arrChrNotVal(cnt_idx)) > 0 Then
								arrValue(1) = ""
								Exit For
							End If
						Next
					Else
						cnt_idx = cnt_idx + 1
					End If
				wend
			End If
		End If

	'--- Check Not Exist ---------------------------------------------------------------------------
	Else
		For cnt_idx = 0 to UBound(arrChrNotVal)
			If InStr(strLine, arrChrNotVal(cnt_idx)) > 0 Then
				Exit For
			ElseIf cnt_idx = UBound(arrChrNotVal) Then
				arrValue(1) = strLine
			End If
		Next
	End If

	IBase_get_value_from_strLine = arrValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_read_text_file(ByVal strPathFile)
	'*** History ***********************************************************************************
	' 2020/08/27, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return a content exists in the file at 'strPathFile' in String type
	'
	'***********************************************************************************************

	On Error Resume Next
	hs_read_text_file = ""

	'*** Pre-Validation ****************************************************************************
	strPathFile = CStr(strPathFile)
	If len(strPathFile) < 1 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim objFSO, objFile, strContent
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'*** Operations ********************************************************************************
	'--- Validate the existence of target file -----------------------------------------------------
	If Not objFSO.FileExists(strPathFile) Then Exit Function

	'--- Read target file --------------------------------------------------------------------------
	Set objFile = objFSO.OpenTextFile(strPathFile, 1)
	strContent  = objFile.ReadAll()

	'--- Release -----------------------------------------------------------------------------------
	hs_read_text_file = strContent
	Set objFile = Nothing
	Set objFSO  = Nothing

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_arr_append(ByRef arrInput, ByVal tarValue)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	' 2020/08/25, BBS:  - Implemented handler for Non-Array 'arrInput'
	' 2020/09/19, BBS: 	- Improved
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Append 'tarValue' to target array provided as 'arrInput', 'arrInput' can be only a single
	'	column array only
	'
	'	Argument(s)
	'	<Array>  arrInput, Base array to be appended 'tarValue'
	'	<Any> 	 tarValue, Desire value to be appended to 'arrInput'
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Initialization ****************************************************************************
	' Nothing to be initialized

	'*** Operations ********************************************************************************
	'--- Ensure 'arrInput' is Array type before doing appending ------------------------------------
	If Not IsArray(arrInput) Then
		arrInput = Array(arrInput)
	End If

	'--- Appending ---------------------------------------------------------------------------------
	If Not (UBound(arrInput) = 0 and LCase(TypeName(arrInput(0))) = "empty") Then
		Redim Preserve arrInput(UBound(arrInput) + 1)
	End If

	arrInput(UBound(arrInput)) = tarValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_arr_stack(ByRef tarValue, ByVal intLevel)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	' 2020/08/25, BBS:	- Bug fixed, Case: 'intLevel' is less than or equal to 0
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Stack 'tarValue' inside out for 'intLevel' additional level
	'	e.g. tarValue = ("SULEV"), intLevel = 2
	' 		 return ((("SULEV")))
	'
	'		 tarValue = "CONT_BAG", intLevel = 2
	'		 return (("CONT_BAG"))
	'
	'	Argument(s)
	'	<Any>  tarValue, Any type of value to be stack
	'	<Long> intLevel, Amount of level
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Pre-Validation ****************************************************************************
	If Not IsNumeric(intLevel) Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, arrRes(), arrTmp()
	Redim Preserve arrRes(0), arrTmp(0)

	intLevel = CInt(intLevel)

	'*** Operations ********************************************************************************
	If intLevel <= 0 Then Exit Function
	If intLevel > 0 Then arrRes(0) = tarValue

	For cnt1 = 2 to intLevel
		Erase arrTmp
		Redim Preserve arrTmp(0)
		
		arrTmp(0) = arrRes(0)
		arrRes(0) = arrTmp
	Next

	tarValue = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_arr_slice(ByVal arrBase, ByVal idxStart, ByVal idxEnd)
	'*** History ***********************************************************************************
	' 2020/12/11, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Array helper, Get sub array of 'arrBase' based on 'idxStart' and 'idxEnd' where both of them
	' can be provided in negative value which means backward counting
	' Example, arrBase = {0, 1, 2, 3, 4, 5}
	'	(idxStart, idxEnd, Result) -> (0, 2, {0, 1, 2}), (3, 7, {3, 4, 5}), (3, 4, {3, 4})
	'								  (-1, -2, {5, 4}), (-2, -1, {4, 5}), (-2, -4, {4, 3, 2})
	'
	' Possible Return Value
	'	<array> Sub array from 'arrBase'
	'	<Null>  If 'arrBase' isn't Array or 'idxStart' and 'idxEnd' are both invalid
	'  
	'***********************************************************************************************

	On Error Resume Next
	hs_arr_slice = Empty

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrBase) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim n_size, cnt, thisStep, arrRet, arrIdx(1)
	n_size 	  = UBound(arrBase)
	arrIdx(0) = idxStart
	arrIdx(1) = idxEnd

	'*** Operations ********************************************************************************
	'--- Conditioning, Indices ---------------------------------------------------------------------
	For cnt = 0 to UBound(arrIdx)
		If arrIdx(cnt) < 0 Then
			If Abs(arrIdx(cnt)) > (1 + n_size) Then
				arrIdx(cnt) = -1*n_size
			End If

			arrIdx(cnt) = n_size + arrIdx(cnt) + 1
		End If

		If arrIdx(cnt) > n_size Then
			arrIdx(cnt) = n_size
		End If
	Next

	'--- Slicing -----------------------------------------------------------------------------------
	n_size = arrIdx(0) - arrIdx(1)
	Redim arrRet(Abs(n_size))

	If n_size > 0 Then
		thisStep = -1
	Else
		thisStep = 1
	End If

	For cnt = arrIdx(0) to arrIdx(1) Step thisStep
		If IsObject(arrBase(cnt)) Then
			Set arrRet(Abs(cnt - arrIdx(0))) = arrBase(cnt)
		Else
			arrRet(Abs(cnt - arrIdx(0))) = arrBase(cnt)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	hs_arr_slice = arrRet

	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function hs_arr_val_exist_ex(ByVal arrInput, ByVal tarValue, ByVal checkMode, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/12/06, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return an index where 'tarValue' is found in 'arrInput'
	' 	Multple indices can be returned depends on 'checkMode'
	'	
	'	Argument(s)
	'	<array> arrInput,	Array to be searched
	'	<str>	tarValue,	Target value in any format but Array, it will be converted to String anyway
	'	<int>	checkMode,	0: exact match, 1: partial match
	'	<bool>	flg_case,	False: Case doesn't matter, True: Case does matter
	'
	'***********************************************************************************************
	
	On Error Resume Next
	hs_arr_val_exist_ex = -1

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrInput) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim idx, thisStr, flg_append, arrIdx()
	Redim Preserve arrIdx(0)
	
	tarValue = CStr(tarValue)

	'*** Operations ********************************************************************************
	'--- Conditioning, flg_case --------------------------------------------------------------------
	If VarType(flg_case) <> 11 Then
		flg_case = LCase(CStr(flg_case))
		
		If flg_case <> "0" and flg_case <> "no" and flg_case <> "false" Then
			flg_case = True
		Else
			flg_case = False
		End If
	End If

	If Not flg_case Then
		tarValue = LCase(tarValue)
	End If

	'--- Conditioning, checkMode -------------------------------------------------------------------
	checkMode = CStr(checkMode)

	If IsNumeric(checkMode) Then
		checkMode = CInt(checkMode)
	Else
		checkMode = 0
	End If

	If checkMode < 0 Then
		checkMode = 0
	ElseIf checkMode > 1 Then
		checkMode = 1
	End If

	'--- Finding -----------------------------------------------------------------------------------
	For idx = 0 to UBound(arrInput)
		If Not IsArray(arrInput(idx)) Then
			thisStr    = CStr(arrInput(idx))
			flg_append = 0

			If Not flg_case Then
				thisStr = LCase(thisStr)
			End If

			' Case: Exact match
			If checkMode = 0 and thisStr = tarValue Then
				flg_append = 1

			' Case: Partial match
			ElseIf checkMode = 1 and InStr(thisStr, tarValue) > 0 Then
				flg_append = 1
			End If

			' Append this index to the result
			If flg_append > 0 Then
				If Not (UBound(arrIdx) = 0 and LCase(TypeName(arrIdx(0))) = "empty") Then
					Redim Preserve arrIdx(UBound(arrIdx) + 1)
				End If

				arrIdx(UBound(arrIdx)) = idx
			End If
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	If Not (UBound(arrIdx) = 0 and LCase(TypeName(arrIdx(0))) = "empty") Then
		If UBound(arrIdx) = 0 Then
			hs_arr_val_exist_ex = arrIdx(0)
		Else
			hs_arr_val_exist_ex = arrIdx
		End If
	End If

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Sub hs_parser_remove_redundant_params(ByRef arrOrder, ByRef arrValue)
	'*** History ***********************************************************************************
	' 2020/12/12, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' Parser helper, Remove redundant parameters
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Pre-Validation ****************************************************************************
	If Not (IsArray(arrParam) and IsArray(arrValue)) Then
		Exit Sub
	End If

	'*** Initialization ****************************************************************************
	Dim idx1, arrCondParam, arrCondValue, objIdx, flg_val

	'*** Operations ********************************************************************************
	For idx1 = 0 to UBound(arrOrder)
		flg_val = False

		If idx1 < UBound(arrOrder) Then
			objIdx = hs_arr_val_exist_ex(hs_arr_slice(arrOrder, idx1 + 1, -1), arrOrder(idx1), 1, False)
		Else
			objIdx = hs_arr_val_exist_ex(hs_arr_slice(arrOrder, -2, 0), arrOrder(idx1), 1, False)
		End If

		If IsArray(objIdx) Then
			If UBound(objIdx) = 0 and IsNumeric(CStr(objIdx(0))) Then
				If CInt(objIdx(0)) < 0 Then
					flg_val = True
				End If
			End If
		ElseIf IsNumeric(CStr(objIdx)) Then
			If CInt(objIdx) < 0 Then
				flg_val = True
			End If
		End If

		If flg_val Then
			Call hs_arr_append(arrCondParam, arrOrder(idx1))
			Call hs_arr_append(arrCondValue, arrValue(idx1))
		End If
	Next

	arrOrder = arrCondParam
	arrValue = arrCondValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Sub
'***************************************************************************************************



'*** Local Material ********************************************************************************
Function IUser_read_json_from_file(ByVal strPathFile, ByVal arrReplaceTag)
	'*** History ***********************************************************************************
	' 2020/08/27, BBS:	- First Release
	' 2020/10/07, BBS:	- Breaks operation into another function for better interface wise
	' 2020/11/06, BBS:	- Slightly adapted on PUMA environment
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return Parameters-Values array of provided JSON file 'strPathFile' (If it does exist)
	'	e.g. (("ecs.devicename", "ecs.activate"), ("CVS_SL", "Yes"))
	' 		 (("gmd.devicename"), ("%tag%0%tag%SULEV%;%%tag%1%tag%CONT_BAG"))
	'
	'***********************************************************************************************

	On Error Resume Next
	IUser_read_json_from_file = Array(Array(), Array())

	'*** Initialization ****************************************************************************
	Dim strContent, RetVal, arrOrder, arrValue

	'*** Operations ********************************************************************************
	'--- Read target JSON file ---------------------------------------------------------------------
	strContent = hs_read_text_file(strPathFile)
	
	If len(strContent) = 0 Then
		Exit Function
	End If

	'--- Translate read JSON file ------------------------------------------------------------------
	RetVal = IUser_read_json(strContent, arrReplaceTag)
	
	If Not IsArray(RetVal) Then
		Exit Function
	End If
	
	'--- Release -----------------------------------------------------------------------------------
	arrOrder = RetVal(0)
	arrValue = RetVal(1)
	IUser_read_json_from_file = Array(arrOrder, arrValue)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function

Function IUser_read_json(ByVal strJsonContent, ByVal arrReplaceTag)
	'*** History ***********************************************************************************
	' 2020/10/07, BBS:	- First Release
	' 2020/11/06, BBS:	- Slightly adapted on PUMA environment
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return Parameters-Values array of provided JSON string content 'strJsonContent'
	'	e.g. (("ecs.devicename", "ecs.activate"), ("CVS_SL", "Yes"))
	' 		 (("gmd.devicename"), ("%tag%0%tag%SULEV%;%%tag%1%tag%CONT_BAG"))
	'
	'***********************************************************************************************

	On Error Resume Next
	IUser_read_json = Array(Array(), Array())

	'*** Pre-Validation ****************************************************************************
	If len(CStr(strJsonContent)) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim RetVal, arrOrder, arrValue

	'*** Operations ********************************************************************************
	'--- Translate read JSON file ------------------------------------------------------------------
	RetVal = IUser_translate_json_strContent(strJsonContent)
	
	If Not IsArray(RetVal) Then
		Exit Function
	End If
	
	'--- Clean Parameter path ----------------------------------------------------------------------
	arrOrder = IUser_clean_ParamPath(RetVal(0), arrReplaceTag)
	
	'--- Release -----------------------------------------------------------------------------------
	arrValue = RetVal(1)
	IUser_read_json = Array(arrOrder, arrValue)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
'***************************************************************************************************
