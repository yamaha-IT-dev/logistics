<%
'H1**********************************************************************************
'Copyright			:Digital Zoo Pty Ltd - 2004
'Library			:FRM_build_form.asp
'
'Creation Date		:4/11/1998
'Version			:1.0
'Author(s)			:Kym Oberauer - Digital Zoo Pty Ltd
'Library Purpose	:This library consists of functions used to create form
'					:elements from a database or list of options. When integrated with other
'					:functions we can create default values etc.
'
'Procedures			:Sub FRM_BuildTwoVariableOptionList(strValueOptionList, strTextOptionList, strSelectedOptions)
'					:Sub FRM_BuildOptionList(strOptionList,strSelectedOption)
'					:Sub FRM_BuildOptionListWithUnselectables(strValueOptionList, strTextOptionList, strSelectedOptions)
'					:Sub FRM_GetDBList(strFirstFieldList, strSecondFieldList, strProcedureName, strProcedureArg)
'					:Sub FRM_RadioList(strValueList,strTextList,strRadioName, strSelectedOption, strDelim, blnLineBreaks)
'					:Sub FRM_RadioItem(strRadioName, strValue, strSelectedOption)
'					:Function FRM_ArrOptionList(arr, ByVal intLevel, ByVal intParentKey, intSelectedKey, intIndentSize, strIndentStr, strLevelChar)
'					:Function FRM_ArrAddOption(arr, intI, intLevel, intIndentSize, strIndentStr, strLevelChar, intSelectedKey)
'					:Function FRM_GetValueByKey(arr, intKey)
'					:function FRM_GetKeyByValue(arr, strValue)
'					:function FRM_ArrAddOptionRow(arr, strValue, intParent, blnChecked, strText)
'					:sub FRM_ArrRecSetToOptionArray(arr, recSet, strValueFld, strTxtFld, intParent, blnChecked)
'					:sub FRM_OptionList(strOptionList, strSelectedOption, strDelim)
'					:sub FRM_OptionListTwoVar(strValueList,strTextList,strSelectedOption, strDelim)
'					:sub FRM_OptionListTwoVarMulti(strValueList, strTextList, strSelectedList, strDelim)
'
'Requirements		:dbase/DB_database.asp
'Special Info		:NA
'Revision History	:################################################################
'					:Simon Abriani - 26.09.2001
'					:Added new function FRM_BuildTwoVariableOptionList() which will
'					:eventually replace/override FRM_BuildTwoVariableOptionList().  The
'					:new function will handle multi select option boxes
'
'					:Indra Hopleythompson - Aug 2003
'					:To close ADO object safely
'
'					:Indra Hopleythompson - May 2004
'					:To include other form building functions
'H2**********************************************************************************


'START ARRAY BASED SELECT LIST FUNCTIONS
Const ARR_OPT_KEY			= 0	'Unique key of option - array index + 1.  Output via the option value attribute, is unique
								'and generated automatically.
const ARR_OPT_VALUE			= 1	'Value of option - eg unique id of record.  This value is not output to HTML, rather
								'is stored in the array and retrieved using the supplied key value.  This value is not
								'output in the HTML as part of the option value attribute as there may be duplicate value
								'causing the html output to produce inconsistent results.

Const ARR_OPT_TEXT			= 2	'Displayed text of option element
Const ARR_OPT_PARENT		= 3	'Key value of parent option list.
Const ARR_OPT_SELECTED		= 4 'True/false, true then option appears selected
Const ARR_OPT_MAXFLDS		= 4
	
	'********************************************************************
	'Function: FRM_BuildTwoVariableOptionList()
	'Arguments: 
	'	strValueOptionList as string - comma delimited string of
	'values
	'	strTextOptionList as string - comma delimited string of
	'text to be displayed.
	'	strSelectedOptions as string - comma delimited string of
	'values that have been selected.
	'	blnMultiSelect as boolean - true = multiselect option box
	'false = NOT multiselect option box
	'Description: This function creates the HTML for and option list, and
	'writes it to the browser immediately.  This function will
	'also highlight the selected options parsed via 
	'strSelectedOptions
	'Output:An option list with the desired values and selected options
	'Special Instructions: If strSelectedOptions is equal to "FIRST" then
	'the first option regardless will be selected.
	'IHT (Oct 05):	additional check in case option lists contain only
	'				one value
	'********************************************************************
	Sub FRM_BuildTwoVariableOptionList(strValueOptionList, strTextOptionList, strSelectedOptions)
		Dim arrValueOptionList
		Dim arrTextOptionList
		Dim arrSelectedOptions
		Dim strCurrentOption
		Dim strSelected
		Dim strFirst
		Dim intLoop
		Dim intI
		
		if inStr(strValueOptionList, ",") < 1 then
			if lcase(trim(strSelectedOptions)) = lcase(trim(strValueOptionList)) then
				response.write("<option SELECTED value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf)
			else
				response.write("<option value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf)
			end if
		else
			arrValueOptionList = Split(strValueOptionList,",",-1)
			arrTextOptionList = Split(strTextOptionList,",",-1)
			if len(strSelectedOptions) > 0 then
				arrSelectedOptions = Split(strSelectedOptions, ",", -1)
			else
				arrSelectedOptions = array()
			end if
		
			for intLoop = 0 to Ubound(arrTextOptionList)
			
				strCurrentOption = arrValueOptionList(intLoop)
				
				if Len(trim(Cstr(strCurrentOption))) > 0 OR Len(trim(arrTextOptionList(intLoop))) > 0 then ' if a vaild string is in either the value or text array
					strSelected = ""
					For intI = 0 To UBound(arrSelectedOptions)
						if lCase(trim(Cstr(strCurrentOption))) = lCase(Trim(Cstr(arrSelectedOptions(intI)))) then 
							strSelected = " selected"
						end if
					Next
					
					If UCASE(strSelectedOptions) = "FIRST" AND intLoop = 0 Then
						response.write("<option SELECTED value='" & arrValueOptionList(intLoop) & "'>" & arrTextOptionList(intLoop) & "</option>" & vbcrlf)
					Else
						response.write("<option" & strSelected & " value='" & arrValueOptionList(intLoop) & "'>" & arrTextOptionList(intLoop) & "</option>" & vbcrlf)	
					End if
				end if
			
			next
		end if
	End Sub
	
	'********************************************************************
	'Function: FRM_BuildOptionList
	'Description: This function builds an option list from a list of supplied
	'options and a selected option.
	'********************************************************************	
	Sub FRM_BuildOptionList(strOptionList,strSelectedOption)

		Dim arrOptionList, strCurrentOption, strSelected
		Dim intLoop

		arrOptionList = Split(strOptionList,",",-1)

		for intLoop = 0 to Ubound(arrOptionList)
			strCurrentOption = arrOptionList(intLoop)
			strSelected = ""
			if strCurrentOption = strSelectedOption then
				strSelected = "selected"
			end if
			response.write("<option " & strSelected & " value='" & strCurrentOption & "'>" & strCurrentOption & "</option>")
		next
	End Sub
	
	'********************************************************************
	'Function: FRM_BuildOptionListWithUnselectables()
	'Author: Indra Hopleythompson (Nov 2005)
	'Arguments: 
	'	strValueOptionList as string - comma delimited string of values
	'	strTextOptionList as string - comma delimited string of text to be displayed.
	'	strSelectableOptionsList as string - comma delimited string of true/false values
	'	strUnselectedClass as string - css class name to format unselectectable options
	'	strSelectedOptions as string - comma delimited string of values that have been selected.
	'Description: This function creates the HTML for an option list, and
	'	writes it to the browser immediately.  This function will
	'	also 
	'	1. highlight the selected options parsed via strSelectedOptions
	'	2. change the formatting and value of options that are not to be selected
	'Output:An option list with the desired values and selected options
	'Special Instructions: If strSelectedOptions is equal to "FIRST" then
	'	the first option regardless will be selected.
	'********************************************************************
	Sub FRM_BuildOptionListWithUnselectables(strValueOptionList, strTextOptionList, strSelectableOptionsList, strUnselectedClass, strSelectedOptions)
		Dim arrValueOptions
		Dim arrTextOptions
		Dim arrSelectableOptions
		Dim arrSelectedOptions
		Dim strCurrentOption
		Dim strSelected
		Dim strFirst
		Dim intLoop
		Dim intI
		
		if inStr(strValueOptionList, ",") < 1 then
			if lcase(trim(strSelectedOptions)) = lcase(trim(strValueOptionList)) then
				response.write("<option SELECTED value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf)
			else
				response.write("<option value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf)
			end if
		else
			'make arrays from the string lists
			arrValueOptions = Split(strValueOptionList,",",-1)
			arrTextOptions = Split(strTextOptionList,",",-1)
			arrSelectableOptions = Split(strSelectableOptionsList,",",-1)
			
			if len(strSelectedOptions) > 0 then
				arrSelectedOptions = Split(strSelectedOptions, ",", -1)
			else
				arrSelectedOptions = array()
			end if
		
			for intLoop = 0 to Ubound(arrTextOptions)
			
				strCurrentOption = arrValueOptions(intLoop)
				
				if Len(trim(Cstr(strCurrentOption))) > 0 OR Len(trim(arrTextOptions(intLoop))) > 0 then ' if a vaild string is in either the value or text array
					strSelected = ""
					For intI = 0 To UBound(arrSelectedOptions)
						if lCase(trim(Cstr(strCurrentOption))) = lCase(Trim(Cstr(arrSelectedOptions(intI)))) then 
							strSelected = " selected"
						end if
					Next

					'check if the option can be selected, if not make value null, and add formatting
					if arrSelectableOptions(intLoop) then
						response.write("<option value='" & arrValueOptions(intLoop) & "' ")
					else
						response.write("<option value='' class='" & strUnselectedClass & "' ")
					end if
					
					If UCASE(strSelectedOptions) = "FIRST" AND intLoop = 0 Then
						response.write("SELECTED>" & arrTextOptions(intLoop) & "</option>" & vbcrlf)
					Else
						response.write(strSelected & ">" & arrTextOptions(intLoop) & "</option>" & vbcrlf)	
					End if
				end if
			
			next
		end if
	End Sub


'********************************************************************
	'Function: FRM_ReturnTwoVariableOptionList()
	'Arguments: 
	'	strValueOptionList as string - comma delimited string of
	'values
	'	strTextOptionList as string - comma delimited string of
	'text to be displayed.
	'	strSelectedOptions as string - comma delimited string of
	'values that have been selected.
	'	blnMultiSelect as boolean - true = multiselect option box
	'false = NOT multiselect option box
	'Description: This function creates the HTML for and option list, and
	'writes it to the browser immediately.  This function will
	'also highlight the selected options parsed via 
	'strSelectedOptions
	'Output:An option list with the desired values and selected options
	'Special Instructions: If strSelectedOptions is equal to "FIRST" then
	'the first option regardless will be selected.
	'IHT (Oct 05):	additional check in case option lists contain only
	'				one value
	'********************************************************************
	Function FRM_ReturnTwoVariableOptionList(strValueOptionList, strTextOptionList, strSelectedOptions)
		Dim arrValueOptionList
		Dim arrTextOptionList
		Dim arrSelectedOptions
		Dim strCurrentOption
		Dim strSelected
		Dim strFirst
		Dim intLoop
		Dim intI
		dim strTemp
		
		if inStr(strValueOptionList, ",") < 1 then
			if lcase(trim(strSelectedOptions)) = lcase(trim(strValueOptionList)) then
				strTemp = strTemp & "<option SELECTED value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf
			else
				strTemp = strTemp & "<option value='" & strValueOptionList & "'>" & strTextOptionList & "</option>" & vbcrlf
			end if
		else
			arrValueOptionList = Split(strValueOptionList,",",-1)
			arrTextOptionList = Split(strTextOptionList,",",-1)
			if len(strSelectedOptions) > 0 then
				arrSelectedOptions = Split(strSelectedOptions, ",", -1)
			else
				arrSelectedOptions = array()
			end if
		
			for intLoop = 0 to Ubound(arrTextOptionList)
			
				strCurrentOption = arrValueOptionList(intLoop)
				
				if Len(trim(Cstr(strCurrentOption))) > 0 OR Len(trim(arrTextOptionList(intLoop))) > 0 then ' if a vaild string is in either the value or text array
					strSelected = ""
					For intI = 0 To UBound(arrSelectedOptions)
						if lCase(trim(Cstr(strCurrentOption))) = lCase(Trim(Cstr(arrSelectedOptions(intI)))) then 
							strSelected = " selected"
						end if
					Next
					
					If UCASE(strSelectedOptions) = "FIRST" AND intLoop = 0 Then
						strTemp = strTemp & "<option SELECTED value='" & arrValueOptionList(intLoop) & "'>" & arrTextOptionList(intLoop) & "</option>" & vbcrlf
					Else
						strTemp = strTemp & "<option" & strSelected & " value='" & arrValueOptionList(intLoop) & "'>" & arrTextOptionList(intLoop) & "</option>" & vbcrlf
					End if
				end if
			next
		end if
		
		FRM_ReturnTwoVariableOptionList = strTemp
	End Function


	'********************************************************************
	'Function: FRM_GetDBList
	'Description: FRM_GetDBList retrieves two lists from a database to form
	'the arguments supplied to create a option list. The purpose of this is
	'to create dynamic option lists with content from a database.
	'********************************************************************
	Sub FRM_GetDBList(strFirstFieldList, strSecondFieldList, strProcedureName, strProcedureArg)
		Dim connObj, recSet,cmdObj,charDelimiter,paraObj, error
		
		charDelimiter = ""

		'Open the connection
		Set connObj = Server.CreateObject("ADODB.Connection")
		connObj.ConnectionString = Session("ConnectionString")
		connObj.Open

		'Creating the command object
		Set cmdObj = Server.CreateObject("ADODB.Command")
		cmdObj.ActiveConnection = connObj
		cmdObj.CommandText = strProcedureName
		cmdObj.CommandType = AdCmdStoredProc
		
		if strProcedureArg <> "" then
			Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4,strProcedureArg)
			cmdObj.Parameters.Append paraObj
		End if
		
		'DEBUG_dbase(cmdObj)

		On Error Resume Next
		set recSet = cmdObj.Execute
		On Error Goto 0

		if connObj.Errors.Count > 0 then
			for each error in connObj.Errors
				UTL_DisplayMessage "Error: " & Error.Description, strMessageText,AdmErrorColour
			next
		else
			If recSet.state > 0 then
				while recSet.EOF = FALSE
					'build delimited string list - ensuring no "accidental" extra delimiters
					strFirstFieldList = strFirstFieldList & charDelimiter & Replace(recSet.Fields(0), ",", "&#44;")
					strSecondFieldList = strSecondFieldList & charDelimiter & Replace(recSet.Fields(1), ",", "&#44;")
					recSet.MoveNext
					charDelimiter = ","	
				wend
			End If
		end if

		'Destroy all the created objects
		Call DB_closeObject(paraObj)
		Call DB_closeObject(cmdObj)
		Call DB_closeObject(recSet)
		Call DB_closeObject(connObj)
	
	End Sub

	'********************************************************************
	'Function: FRM_GetDBList_Alt
	'Description: FRM_GetDBList_Alt retrieves two lists from a database to form
	'the arguments supplied to create a option list. The purpose of this is
	'to create dynamic option lists with content from a database.
	' differs from FRM_GetDBList in that you can supply 2 arguments
	'********************************************************************
	Sub FRM_GetDBList_Alt(strFirstFieldList, strSecondFieldList, strProcedureName, strIntArgList, charDelimiter)
		Dim connObj, recSet, cmdObj, paraObj, error, charListDelimiter
		dim arrArguments, intParameter
		
		charListDelimiter = ""
		arrArguments = split(strIntArgList, charDelimiter)
		'call debug_array("arrArguments", arrArguments)
		
		'Open the connection
		Set connObj = Server.CreateObject("ADODB.Connection")
		connObj.ConnectionString = Session("ConnectionString")
		connObj.Open

		'Creating the command object
		Set cmdObj = Server.CreateObject("ADODB.Command")
		cmdObj.ActiveConnection = connObj
		cmdObj.CommandText = strProcedureName
		cmdObj.CommandType = AdCmdStoredProc

		For Each intParameter In arrArguments
			If trim(intParameter) = "" or isNull(intParameter) Then
				intParameter = Empty
			End If
			Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4,intParameter)
			cmdObj.Parameters.Append paraObj		
		Next

		On Error Resume Next
		set recSet = cmdObj.Execute
		On Error Goto 0
		
		if connObj.Errors.Count > 0 then
			for each error in connObj.Errors
				UTL_DisplayMessage "Error: " & Error.Description, strMessageText,AdmErrorColour
			next
		else
			If recSet.state > 0 then
				while recSet.EOF = FALSE
					'build delimited string list - ensuring no "accidental" extra delimiters
					strFirstFieldList = strFirstFieldList & charDelimiter & Replace(recSet.Fields(0), ",", "&#44;")
					strSecondFieldList = strSecondFieldList & charDelimiter & Replace(recSet.Fields(1), ",", "&#44;")
					recSet.MoveNext
					charListDelimiter = ","	
				wend
			End If
		end if

		'Destroy all the created objects
		Call DB_closeObject(paraObj)
		Call DB_closeObject(cmdObj)
		Call DB_closeObject(recSet)
		Call DB_closeObject(connObj)
	
	End Sub


'P1**********************************************************************************
'Procedure			:FRM_RadioList
'Arguments			:strValueList as string - a delimited list of option values
'					:strTextList as string - a delimited list of text for each <option>
'					:strRadioName
'					:strSelectedOption as string - a selected option value, if none then ""
'					:strDelim as string - delimiter for the value and text lists
'					:blnLineBreaks
'
'Purpose			:This function builds a radio list with a VALUE attribute
'					:from a list of supplied values and radio button names and 
'					:a selected button if specified. The radio button list is written
'					:directly to the response stream
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:
'P2**********************************************************************************
sub FRM_RadioList(strValueList,strTextList,strRadioName, strSelectedOption, strDelim, blnLineBreaks)
	dim arrValues		' - <input type="radio" VALUE=?> 
	dim arrText			' - <radio>TEXT
	dim strCurrRadio
	dim strChecked
	dim intLoop
	
	arrValues = Split(strValueList, strDelim,-1)
	arrText = Split(strTextList, strDelim,-1)
	
	for intLoop = 0 to ubound(arrValues)
		strCurrRadio = arrValues(intLoop)
		' if a vaild string 
		' is in the array
		if len(cstr(strCurrRadio)) > 0 then 
			strChecked = ""
			if cstr(strCurrRadio) = cstr(strSelectedOption) then
				strChecked = " CHECKED"
			end if
			If blnLineBreaks AND intLoop <> 0 Then
				Response.Write "<BR>"
			End If
			Response.Write "<input type='radio' name='" & strRadioName & "' ID='" & strRadioName & cStr(intLoop+1) & "' value='" & arrValues(intLoop) & "'" & strChecked & "> " & arrText(intLoop) & " &nbsp; " & vbcrlf
		end if
	next
	
end sub




'P1**********************************************************************************
'Procedure			:FRM_ReturnRadioList
'Arguments			:strValueList as string - a delimited list of option values
'					:strTextList as string - a delimited list of text for each <option>
'					:strRadioName
'					:strSelectedOption as string - a selected option value, if none then ""
'					:strDelim as string - delimiter for the value and text lists
'					:blnLineBreaks
'
'Purpose			:This function builds a radio list with a VALUE attribute
'					:from a list of supplied values and radio button names and 
'					:a selected button if specified. The radio button list is returned
'					:in a html string
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:
'P2**********************************************************************************
Function FRM_ReturnRadioList(strValueList,strTextList,strRadioName, strSelectedOption, strDelim, blnLineBreaks)
	dim arrValues		' - <input type="radio" VALUE=?> 
	dim arrText			' - <radio>TEXT
	dim strCurrRadio
	dim strChecked
	dim intLoop
	dim strList
	
	arrValues = Split(strValueList, strDelim,-1)
	arrText = Split(strTextList, strDelim,-1)
	
	for intLoop = 0 to ubound(arrValues)
		strCurrRadio = arrValues(intLoop)
		' if a vaild string 
		' is in the array
		if len(cstr(strCurrRadio)) > 0 then 
			strChecked = ""
			if cstr(strCurrRadio) = cstr(strSelectedOption) then
				strChecked = " CHECKED"
			end if
			If blnLineBreaks AND intLoop <> 0 Then
				strList = strList & "<BR>"
			End If
			strList = strList & "<input type='radio' name='" & strRadioName & "' ID='" & strRadioName & cStr(intLoop+1) & "' value='" & arrValues(intLoop) & "'" & strChecked & "> " & arrText(intLoop) & " &nbsp; " & vbcrlf
		end if
	next
	
	FRM_ReturnRadioList = strList
end Function


'P1**********************************************************************************
'Procedure			:FRM_RadioItem
'Arguments			:strRadioName as string - the name of the radio button
'					:strValue as string - the value of the radio button
'					:strSelectedOption as string - the compare value, to estalish if "checked"
'					:blnReadOnly as boolean - whether the button is read only
'Purpose			:This function builds a radio list with a VALUE attribute
'					:from a list of supplied values and radio button names and 
'					:a selected button if specified. The radio button list is written
'					:directly to the response stream
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:
'P2**********************************************************************************
Sub FRM_RadioItem(strRadioName, strValue, strSelectedOption, blnReadOnly)
	
	If cstr(strValue) = cstr(strSelectedOption) then
		If blnReadOnly Then
			Response.Write "<input type='radio' name='" & strRadioName & "' value='" & strValue & "' CHECKED DISABLED>" & vbcrlf
			
			'to ensure a value is sent regardless if button is disabled
			Response.Write "<input type='hidden' name='" & strRadioName & "' value='" & strValue & "'>" & vbcrlf
		Else
			Response.Write "<input type='radio' name='" & strRadioName & "' value='" & strValue & "' CHECKED>" & vbcrlf
		End If
	Else
		If blnReadOnly Then
			Response.Write "<input type='radio' name='" & strRadioName & "' value='" & strValue & "' DISABLED>" & vbcrlf
		Else
			Response.Write "<input type='radio' name='" & strRadioName & "' value='" & strValue & "'>" & vbcrlf
		End If
	End if
	
End Sub

'P1**********************************************************************************
'Procedure			:FRM_ArrOptionList
'Arguments			:arr as array - option list array
'					:ByVal intLevel as integer - current level
'					:ByVal intParentKey as integer - key of parent option
'					:intSelectedKey as integer - key of selected option.
'Purpose			:This procedure processess an initialised option list array outputting
'					:the neccessary HTML to depict the nested nature of option items.
'Output				:String of html, representing the nested HTML Select List
'Dependancies		:NA
'Special Info		:This function is recursive to achieve the nesting of option
'					:list items.
'P2**********************************************************************************
Function FRM_ArrOptionList(arr, ByVal intLevel, ByVal intParentKey, intSelectedKey, intIndentSize, strIndentStr, strLevelChar)
    Dim intI
    Dim str
    
    If len(intLevel) = 0 Then
        intLevel = 1
    Else
        intLevel = intLevel + 1
    End If
    
    If len(intParentKey) = 0 Then
        intParentKey = 0
    End If
    
    For intI = LBound(arr, 2) To UBound(arr, 2)
        If intParentKey = arr(ARR_OPT_PARENT, intI) Then
            str = str & FRM_ArrAddOption(arr, intI, intLevel, intIndentSize, strIndentStr, strLevelChar, intSelectedKey)
            str = str & FRM_ArrOptionList(arr, intLevel, arr(ARR_OPT_KEY, intI), intSelectedKey, intIndentSize, strIndentStr, strLevelChar)
        End If
    Next
    
    FRM_ArrOptionList = str
End Function


'P1**********************************************************************************
'Procedure			:FRM_ArrAddOption
'Arguments			:arr as array - option array
'					:intI as integer - current array index
'					:intLevel as integer - current level to process
'					:intIndentSize as integer - default indent size (number of times to repeart strIndentStr)
'					:strIndentStr as string - indent string
'					:strLevelChar as string - string to append to start of indented children text
'					:intSelectedKey as integer - key of selected option (optional)
'Purpose			:Returns the html for the current processed option element
'Output				:String
'Dependancies		:UTL_utilities.asp
'Special Info		:NA
'P2**********************************************************************************
Function FRM_ArrAddOption(arr, intI, intLevel, intIndentSize, strIndentStr, strLevelChar, intSelectedKey)
    Dim str
    
    If arr(ARR_OPT_SELECTED, intI) = CInt(True) Or cstr(intSelectedKey) = cstr(arr(ARR_OPT_KEY, intI)) Then
        str = "<option value='" & arr(ARR_OPT_KEY, intI) & "' SELECTED>"
    Else
        str = "<option value='" & arr(ARR_OPT_KEY, intI) & "'>"
    End If
    
    If intLevel > 1 Then
        str = str & UTL_StringRepeat(strIndentStr, intIndentSize * intLevel) & strLevelChar
    End If
    
    str = str & arr(ARR_OPT_TEXT, intI)
    
    str = str & "</option>" & Chr(13)
    
    FRM_ArrAddOption = str
End Function


'P1**********************************************************************************
'Procedure			:FRM_GetValueByKey
'Arguments			:arr as array - option array
'					:intKey as integer - key value of array
'Purpose			:Returns the ARR_OPT_VALUE element for the array row specified
'					:intKey.  If intKey exceeds the array bounds or arr is not an array
'					:false is returned.
'Output				:String
'Dependancies		:NA
'Special Info		:NA
'P2**********************************************************************************
Function FRM_GetValueByKey(arr, intKey)
    If IsArray(arr) Then
        If (intKey - 1) >= LBound(arr, 2) And (intKey - 1) <= UBound(arr, 2) Then
            FRM_GetValueByKey = arr(ARR_OPT_VALUE, intKey - 1)
        Else
            FRM_GetValueByKey = False
        End If
    Else
        FRM_GetValueByKey = False
    End If
End Function


'P1**********************************************************************************
'Procedure			:FRM_GetKeyByValue
'Arguments			:arr as array - option array
'					:strValue as string - value of array option to return key of
'Purpose			:Traverses the array to find a value element that matches strValue.
'					:When the first match is found, the key of the array row is returned.
'Output				:Integer
'Dependancies		:NA
'Special Info		:NA
'P2**********************************************************************************
function FRM_GetKeyByValue(arr, strValue)
	dim blnFound
	dim intI

	blnFound = false
	If IsArray(arr) Then
		for intI = lbound(arr,2) to ubound(arr,2)
			if cstr(arr(ARR_OPT_VALUE, intI)) = cstr(strValue) then
				blnFound = true
				exit for
			end if
		next
       
		if blnFound then
            FRM_GetKeyByValue = arr(ARR_OPT_KEY, intI)
        Else
            FRM_GetKeyByValue = False
        End If
    Else
        FRM_GetKeyByValue = False
    End If
end function


'P1**********************************************************************************
'Procedure			:FRM_ArrAddOptionRow
'Arguments			:arr as array - option array (initialise, not initialised)
'					:strValue as integer - 
'					:intParent as integer - key value of parent option
'					:blnChecked as boolean - true is selected, otherwise not selected
'					:strText as string - 
'Purpose			:This procedure handles the addition of new option items to the option
'					:array.
'Output				:Key field of newly added option row
'Dependancies		:NA
'Special Info		:NA
'P2**********************************************************************************
function FRM_ArrAddOptionRow(arr, strValue, intParent, blnChecked, strText)
    Dim intRow
    If IsArray(arr) Then
        intRow = UBound(arr, 2) + 1
        ReDim Preserve arr(ARR_OPT_MAXFLDS, intRow)
    Else
        intRow = 0
        ReDim arr(ARR_OPT_MAXFLDS, intRow)
    End If

    arr(ARR_OPT_KEY, intRow) = intRow + 1
    arr(ARR_OPT_VALUE, intRow) = strValue
    arr(ARR_OPT_TEXT, intRow) = strText
    arr(ARR_OPT_PARENT, intRow) = intParent
    arr(ARR_OPT_SELECTED, intRow) = CInt(blnChecked)
    
    FRM_ArrAddOptionRow = arr(ARR_OPT_KEY, intRow)
End function


'P1**********************************************************************************
'Procedure			:FRM_ArrRecSetToOptionArray
'Arguments			:arr as array - option array
'					:recSet as ADODB.Recordset
'					:strValueFld as string - field name of recordset to store as value field
'					:strTxtFld as string - field name of recordset to store as text field
'					:intParent as integer - parent key
'					:blnChecked as boolean - selected or not selected.
'Purpose			:This procedure adds new option array rows for each record in the recordset.
'Output				:NA
'Dependancies		:NA
'Special Info		:NA
'P2**********************************************************************************
sub FRM_ArrRecSetToOptionArray(arr, recSet, strValueFld, strTxtFld, intParent, blnChecked)
	recSet.MoveFirst
	do until recSet.EOF
		call FRM_ArrAddOptionRow(arr, recSet(strValueFld), intParent, blnChecked, recSet(strTxtFld))
		recSet.movenext
	loop
	recSet.MoveFirst
end sub

'**********************************************************
'END ARRAY BASED SELECT LIST FUNCTIONS
'**********************************************************

'P1**********************************************************************************
'Procedure			:FRM_OptionList
'Arguments			:strOptionList as string - comma delimited list of option values
'					:strSelectedOption as string - value of selected option
'					:strDelim as string - strOptionList value delimiter
'Purpose			:This function builds an option list from a list of supplied
'					:options and a selected option.  The option list is written
'					:directly to the response stream
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:Note that this function does not return the required <select>
'					:tags to properly display a list of options
'P2**********************************************************************************
sub FRM_OptionList(strOptionList, strSelectedOption, strDelim)
	dim arrOptionList	'Array of option values
	dim strCurrOption
	dim strSelected
	dim intLoop
	
	arrOptionList = Split(strOptionList, strDelim, -1)
	for intLoop = 0 to ubound(arrOptionList)
		strCurrOption = arrOptionList(intLoop)
		strSelected = ""
		if strCurrOption = strSelectedOption then
			strSelected = "selected"
		end if
		response.write("<option " & strSelected & " value='" & strCurrOption & "'>" & _
			strCurrOption & "</option>" & Chr(13))
	next
	
end sub


'P1**********************************************************************************
'Procedure			:FRM_OptionListTwoVar
'Arguments			:strValueList as string - a delimited list of option values
'					:strTextList as string - a delimited list of text for each <option>
'					:strSelectedOption as string - a selected option value, if none then ""
'					:strDelim as string - delimiter for the value and text lists
'
'Purpose			:This function builds an option list with a VALUE attribute
'					:from a list of supplied options values and text values and 
'					:a selected option if specified. The option list is written
'					:directly to the response stream
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:Note that this function does not return the <select> tags
'					:that are required to display a list of options.
'P2**********************************************************************************
sub FRM_OptionListTwoVar(strValueList,strTextList,strSelectedOption, strDelim)
	dim arrValues		' - <option VALUE> 
	dim arrText			' - <option>TEXT</option
	dim strCurrOption
	dim strSelected
	dim intLoop
	
	arrValues = Split(strValueList, strDelim,-1)
	arrText = Split(strTextList, strDelim,-1)
	
	for intLoop = 0 to ubound(arrValues)
		strCurrOption = arrValues(intLoop)
		' if a vaild string 
		' is in the array
		if len(cstr(strCurrOption)) > 0 then 
			strSelected = ""
			if cstr(strCurrOption) = cstr(strSelectedOption) then
				strSelected = " selected"
			end if
			response.write("<option" & strSelected & " value='" & arrValues(intLoop) & _
				"'>" & arrText(intLoop) & "</option>" & Chr(13))
		end if
	next
	
end sub


'P1**********************************************************************************
'Procedure			:FRM_OptionListTwoVarMulti
'Arguments			:strValueList as string - delimited string of values
'					:strTextList as string - delimited string of text to display
'					:strSelectedList as string - delimited string of selected values
'					:strDelim as string - common delimiter for all lists
'
'Purpose			:This function creates the HTML for a two variable option list
'					:that allows multiple values to be selected as specified.  The
'					:option list is written directly to the response stream.
'Output				:NA
'Dependancies		:NA
'Dependants			:NA
'Special Info		:If strSelectedList is equal to "FIRST" then the first
'					:option regardless will be selected.
'P2**********************************************************************************
sub FRM_OptionListTwoVarMulti(strValueList, strTextList, strSelectedList, strDelim)
	dim arrValues		'Array of values
	dim arrText			'Array of text values
	dim arrSelected		'Array of selected option values
	dim strCurrOption	
	dim strSelected
	dim strFirst
	dim intLoop
	dim intI
		
	arrValues = Split(strValueList, strDelim,-1)
	arrText = Split(strTextList, strDelim,-1)
	arrSelected = Split(strSelectedList, strDelim, -1)
	
	for intLoop = 0 to ubound(arrText)
		strCurrOption = arrValues(intLoop)
		
		' if a vaild string is in 
		'either the value or text array	
		if len(cstr(strCurrOption)) > 0 _
			or len(arrText(intLoop)) > 0 then
			
			strSelected = ""
			
			for intI = 0 to ubound(arrSelected)
				if cstr(strCurrOption) = Trim(cstr(arrSelected(intI))) then 
					strSelected = " selected"
				end if
			next
				
			if UCASE(strSelectedList) = "FIRST" _
				and intLoop = 0 then
				
				response.write("<option SELECTED value='" & arrValues(intLoop) & "'>" & _
					arrText(intLoop) & "</option>" & Chr(13))
			else
				response.write("<option" & strSelected & " value='" & arrValues(intLoop) & _
					"'>" & arrText(intLoop) & "</option>" & Chr(13))	
			end if
		end if
		
	next
end sub
%>
