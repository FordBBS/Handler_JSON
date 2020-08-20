Function IBase_translate_json_strContent(ByVal strContent)
	'*** History ***********************************************************************************
	' 2020/08/19, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return Parameters-Values array of provided JSON content in 'strContent' String type
	'	e.g. (("ecs.devicename", "ecs.activate"), ("CVS_SL", "Yes"))
	'
	'***********************************************************************************************
	On Error Resume Next
	IBase_translate_json_strContent = Array("", "")

	'*** Pre-Validation ****************************************************************************
	If TypeName(strContent) <> "String" Then Exit Function
	If len(strContent) < 3 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim arrContent, arrThisInfo, strParamApp, arrParam(), arrValue()
	Dim cnt_row, curRoot, curParamPath, thisParam, flg_append
	Dim strTagArray, strTagValue, tarRemTag
	Redim Preserve arrParam(0), arrValue(0)

	strTagArray = ""
	strTagValue = ""
	strParamApp = ""

	'*** Operations ********************************************************************************
	'--- Input Data Preparation --------------------------------------------------------------------
	arrContent = Split(strContent, vbCrLf) 			'Split each line into Array

	'--- Parsing -----------------------------------------------------------------------------------
	cnt_row = 0
	curRoot = ""

	while cnt_row < UBound(arrContent)
		' Init current line
		arrThisInfo = IBase_get_value_from_strLine(arrContent(cnt_row), ":")
		thisParam   = arrThisInfo(0)
		flg_append	= False

		' Case: Parameter does exist
		If thisParam <> "" Then
			If InStr(arrContent(cnt_row), "{") > 0 or InStr(arrContent(cnt_row), "[") > 0 Then
				curRoot  	= curRoot & thisParam & "."
				strTagValue = strTagValue & "%" & thisParam & "%;"
				
				If InStr(arrContent(cnt_row), "[") > 0 Then
					strTagArray = strTagArray & "%" & thisParam & "%;"
				End If
			Else
				flg_append = True
			End If

		' Case: Parameter doesn't exist
		ElseIf thisParam = "" Then
			' Case: End of current sub-tags group, clear Root, and stored tags string
			If InStr(arrContent(cnt_row), "}") > 0 Then
				curRoot   = Left(curRoot, len(curRoot) - 1)
				cnt_idx   = InStrRev(curRoot, ".")
				tarRemTag = Mid(curRoot, cnt_idx + 1, len(curRoot))

				If InStr(strTagArray, "%" & tarRemTag & "%") = 0 and _
				 	InStr(strTagValue, "%" & tarRemTag & "%") > 0 Then
					curRoot = Replace(curRoot, tarRemTag, "")

					If InStr(strTagValue, "%" & tarRemTag & "%") > 0 Then
						strTagValue = Left(strTagValue, len(strTagValue) - 1)
						cnt_idx 	= InStrRev(strTagValue, ";")
						strTagValue = Mid(strTagValue, 1, cnt_idx)
					End If
				Else
					curRoot = curRoot & "."
				End If
			
			' Case: End of latest branch, clear latest 'TagArray' out from 'strTagArray'
			ElseIf InStr(arrContent(cnt_row), "]") > 0 Then
				strTagArray = Left(strTagArray, len(strTagArray) - 1)
				cnt_idx 	= InStrRev(strTagArray, ";")
				strTagArray = Mid(strTagArray, 1, cnt_idx)
			
			' Case: Value line (e.g. Parameter that has array value will break its value into lines)
			ElseIf InStr(arrContent(cnt_row), "{") = 0 and InStr(arrContent(cnt_row), "[") = 0 Then
				flg_append  = True
			End If
		End If

		' Appending
		If flg_append Then
			curParamPath = curRoot & thisParam

			' Case: Current Parameter path already exist
			If InStr(strParamApp, "%" & curParamPath & "%") > 0 Then
				For cnt_idx = 0 to UBound(arrParam)
					If arrParam(cnt_idx) = curParamPath Then Exit For
				Next

			' Case: Current Parameter path has its first time appending
			Else
				strParamApp = strParamApp & "%" & curParamPath & "%;"

				If Not (UBound(arrParam) = 0 and len(arrParam(0)) = 0) Then
					Redim Preserve arrParam(UBound(arrParam) + 1), arrValue(UBound(arrValue) + 1)
				End If

				arrParam(UBound(arrParam)) = curParamPath
				arrValue(UBound(arrValue)) = arrThisInfo(1)
			End If
		End If

		' Release current line
		If cnt_row < 0 Then
			cnt_row = UBound(arrContent)
		Else
			cnt_row = cnt_row + 1
		End If
	wend
	
	IBase_translate_json_strContent = Array(arrParam, arrValue)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function