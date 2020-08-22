Function IBase_get_value_from_strLine(ByVal strLine, ByVal chr_sep)
	'*** History ***********************************************************************************
	' 2020/08/20, BBS:	- First Release
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

			If InStr(cnt_idx, strLine, chr(34)) > 0 Then
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