Function IBase_get_value_from_strLine(ByVal strLine, ByVal chr_sep)
	'*** History ***********************************************************************************
	' 2020/08/20, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return array of Parameter or Value available on 'strLine' with JSON structure
	'	e.g. strJsonLine = "testcell_name": "LD06", 	-> ("testcell_name", "LD06")
	' 		 strJsonLine = "config_params": [ 			-> ("config_params")
	' 		 strJsonLine = }, 							-> ("")
	'
	'	Argument(s)
	'	<String> strLine, A String of content line to be parsed
	'	<String> chr_sep, A character used to separate between Parameter and Value
	'						e.g. chr_sep = "=" for XML, chr_sep = ":" for JSON
	'
	'***********************************************************************************************
	On Error Resume Next
	IBase_get_value_from_jsonline = Array("")

	'*** Pre-Validation ****************************************************************************
	If LCase(TypeName(strJsonLine)) <> "string" Then Exit Function
	If len(strJsonLine) < 1 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim arrValue()
	Redim Preserve arrValue(0)

	'*** Operations ********************************************************************************
	'--- Clear spaces and tabs on left and right sides ---------------------------------------------
	strJsonLine = Trim(strJsonLine)


	'--- Check Exist -------------------------------------------------------------------------------
	If Instr(strJsonLine, chr(34)) > 0 Then


	'--- Check Not Exist ---------------------------------------------------------------------------
	Else


	End If

	IBase_get_value_from_jsonline = arrValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function