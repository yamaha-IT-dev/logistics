<%
'**************************************************************
'Author: Bronwyn Streeter - Simon Abriani Digital Zoo Pty Ltd
'Copyright Digital Zoo Pty Ltd 2000
'Date: 26/10/2000 Version 1.0
'Revision History:
'Library: utility.asp
'Description: This library consists of various useful tools to concatenate
'strings into various formats, find the minimum/maximum value, and various
'data manipulation tools.
'**************************************************************

'***************************************************************
'Function Name: UTL_TruncateString()
'Arguments: 
'	strVal as string - string to truncate.
'	lngTruncateLen as string - length to truncate to
'	lngMax as string - max length before truncation occurs
'	strTruncateText as string - string to concatenate to 
'	end of string to show that truncation has taken place
'Description: Truncates a string to a desired length if required and concatenates
'the specfied text to show truncation has taken place. Otherwise returns
'value of strVal unaltered.
'Output: String
'Special Instructions:NA
'***************************************************************
Function UTL_TruncateString(strVal, lngTruncateLen, lngMax, strTruncateText)
	If Len(strVal) > Clng(lngMax) Then
		UTL_TruncateString = Left(strVal,lngTruncateLen) & strTruncateText
	Else
		UTL_TruncateString = strVal
	End if
End Function

'***************************************************************
'Function Name: UTL_CombineArraysTwoDim()
'Arguments: 
'	arr1 as array - array to append to
'	arr2 as array - array to arr1
'Description: Combines 2 two dimensional arrays.
'Output: Returns a new array combining the data or arr1 and arr2
'Special Instructions:Note, that only the 2nd dimension can
'be appended to (2 dim array limitation)
'***************************************************************
Function UTL_CombineArraysTwoDim(arr1,arr2)
	Dim lngUBoundArrayOne1, lngUBoundArrayOne2
	Dim lngUBoundArrayTwo1, lngUBoundArrayTwo2
	Dim lngUBoundArrayNew1, lngUBoundArrayNew2
	Dim intOneCount, intTwoCount
	Dim arrNew
	Dim intI
	
	lngUBoundArrayOne1 = UBound(arr1,1)
	lngUBoundArrayOne2 = UBound(arr1,2)
	
	lngUBoundArrayTwo1 = UBound(arr2,1)
	lngUBoundArrayTwo2 = UBound(arr2,2)
	
	lngUBoundArrayNew1 = UBound(arr1,1)
	lngUBoundArrayNew2 = lngUBoundArrayOne2 + lngUBoundArrayTwo2
	'Assign data from first array to the new array
	arrNew = arr1
	
	'Redimension the array to append the data to, while preserving the data
	ReDim Preserve arrNew(lngUBoundArrayNew1, lngUBoundArrayNew2 + 1)
	
	'Now all we have to do is loop through the second array
	'and attached the contents from the second array
	For intOneCount = 0 to lngUBoundArrayNew1 
		intI = 0
		For intTwoCount = lngUBoundArrayOne2 + 1 To lngUBoundArrayNew2 + 1
			arrNew(intOneCount, intTwoCount) = arr2(intOneCount, intI)
			If intI >= lngUBOundArrayTwo2 Then
				Exit For
			Else
				intI = intI + 1
			End if	
		Next
		
	Next
	UTL_CombineArraysTwoDim = arrNew
End Function


'***************************************************************
'Function Name: UTL_ConcatenateString
'Arguments: 
'	strInitialString as string - first string
'	strSecondString as string - second string added to first string
'	strConnectionString as string - sdtring used to join the intital string with
'	the second string
'	strTerminationString as string - string to end the concatenation
'Description: concatenates string1 and string2, placing conneciton_string between
'them and adding termination_string to the end of the new string.  Empty strings
'can be used for the connection_string and termination_string.
'Output: Returns the concatenated string if successful or an empty string
'Special Instructions:NA
'***************************************************************
Function UTL_ConcatenateString(strInitialString, strSecondString, strConnectionString, strTerminationString)
	UTL_ConcatenateString = strInitialString & strConnectionString & strSecondString & strTerminationString
End Function

'***************************************************************
'Function Name: UTL_ConcatenateDate
'Arguments: 
'	intDay as number - number representing a day (1 - 31)
'	intMonth as number - number representing a month (1 - 12)
'	intYear as number - number represnting a year (9999)
'Description: Takes the values supplied and concatenates them into a valid date string
'Output: Returns a string representing a valid date if successful or
'in empty string.
'Special Instructions: For some reason it would seem that the
'built in date functions in VB are in american format. It
'does not look at the local system settings????
'***************************************************************
Function UTL_ConcatenateDate(intDay, intMonth, intYear)
	dim valid
	dim OutputString
	
	OutputString = intMonth & "/" & intDay & "/" & intYear
	'OutputString = intDay & "/" & intMonth & "/" & intYear

	valid = VAL_TestValidDate(OutputString)
	
	if (valid = True) then 
		UTL_ConcatenateDate = OutputString
	else
		UTL_ConcatenateDate = ""
	end if
			
End Function
		
'***************************************************************
'Function Name: UTL_ArrayAsTable
'Arguments: 
'	arrData as array - array of values to output as an HTML table
'	strHeadings as string - a comma seperated list of heading names. They must correspond to
'	the order of the fields in the array. If the string is empty they are
'	ignored.
'	bolBorder as boolean - TRUE to display table border, FALSE otherwise
'Description: Takes an array, determines the number of rows and columns required. EAch
'array row is converted to an HTML row with array values places in each cell.
'Output: String which contains the formatted HTML if successful or an empty
'string otherwise.
'Special Instructions: Only works with a two dimensional array.
'***************************************************************
Function UTL_ArrayAsTable(arrData,strHeadings, bolBorder)
	dim strOutputString, arrHeadings, strHeading
	dim intFirst,intSecond
	if isarray(arrData) then
		'Determine if table border requested
		if bolBorder = TRUE then
			strOutputString = "<TABLE BORDER='1'>" & Chr(10)
		else
			strOutputString = "<TABLE BORDER='0'>" & Chr(10)
		end if
	
		'Determine if headings are required
		if strHeadings <> "" then
			strOutputString = strOutputString & "<TR>"
			arrHeadings = Split(strHeadings,",")
			for each strHeading in arrHeadings
				strOutputString = strOutputString & "<TD>" & strHeading & "</TD>"
			next
			strOutputString = strOutputString & "</TR>"
		end if

		'Create table from array
		for intFirst = LBound(arrData,1) to Ubound(arrData,1)
			strOutputString = strOutputString & "<TR>"
			for intSecond = Lbound(arrData,2) to Ubound(arrData,2)
				if not(isArray(arrData(intFirst, intSecond))) then 'stops arrays with arrays from begin processed.
					if len(arrData(intFirst, intSecond)) > 0 then
						strOutputString = strOutputString & "<TD>" & arrData(intFirst, intSecond) & "</TD>"
					else
						strOutputString = strOutputString & "<TD>&nbsp;</TD>"
					end if
				end if
			next
			strOutputString = strOutputString & chr(10) & "</TR>" & Chr(10)
		next
	
		strOutputString = strOutputString & "</TABLE>"
	else
		strOutputString = "No data in array, or array not initialised."
	end if
	UTL_ArrayAsTable = strOutputString
End Function

'***************************************************************
'Function Name: UTL_StripChars
'Arguments: 
'	strInitialString as string - string to strip characetrs from
'	strCharsToStrip as string - string containing chars to strip
'	chrDelimiter as character - not used
'Description: This function accepts a string from which characters are to be 
'removed. The chars_to_strip argument contains a list of all the characters
'to be removed.
'Output: Returns a string with the required characters removed if successful or an 
'empty string otherwise.
'Special Instructions: Each charater will be stripped. Does not allow for
'more than 1 character patten to be stripped.
'***************************************************************
Function UTL_StripChars(strInitialString, strCharsToStrip, chrDelimiter)
	dim strFinalOutput
	dim intPositionStart, intPositionEnd, intStripCharPosition
	dim chrStripChar
	dim intEscapeFlag
	
	intEscapeFlag = 0
	intStripCharPosition = 1
	while (intEscapeFlag = 0)
		chrStripChar = Mid(strCharsToStrip,intStripCharPosition,1) 'gets the char to strip
		intPositionStart = 1	'resets the starting position
		strFinalOutput = ""
		do while (intEscapeFlag = 0)	
			intPositionEnd = InStr(intPositionStart,strInitialString,chrStripChar) 'looks for the first instance from character 1
			if intPositionEnd = 0 then 'if the string cannot be found
				strFinalOutput = strFinalOutput & Mid(strInitialString, intPositionStart, (len(strInitialString) + 1) - intPositionStart)
				exit do
			end if
			strFinalOutput = strFinalOutput & Mid(strInitialString, intPositionStart, intPositionEnd - intPositionStart)
			intPositionStart = intPositionEnd + 1
		loop
		strInitialString = strFinalOutput
		intStripCharPosition = intStripCharPosition + 2
		if (intStripCharPosition > Len(strCharsToStrip)) then
			intEscapeFlag = 1
		end if
	wend
	UTL_StripChars = strFinalOutput
End Function


'***************************************************************
'Function Name: UTL_CharsSafeForHtml
'Arguments: 
'	strStringToConvert as string - the string to remove characters from
'Description: Strips characters for the provided string. The point of
'this function is to remove URL formating charaters.
'Output: the stripped string
'Special Instructions:NA
'***************************************************************
Function UTL_CharsSafeForHtml(strStringToConvert)
	'NB: needs to be developed!!!!
	
	
	'strStringToConvert = Replace(strStringToConvert,"+"," ")
	'strStringToConvert = Replace(strStringToConvert," ","")
	'strStringToConvert = Replace(strStringToConvert,"%2C","")
'	strStringToConvert = Replace(strStringToConvert,"%2F","/")
'	strStringToConvert = Replace(strStringToConvert,"%24","$")
'	strStringToConvert = Replace(strStringToConvert,"%27","'")
'	strStringToConvert = Replace(strStringToConvert,"%5C","\")
'	strStringToConvert = Replace(strStringToConvert,"%22",Chr(34))

	UTL_CharsSafeForHtml = strStringToConvert
End Function

'***************************************************************
'Function Name: UTL_CharsSafeForJavascript
'Arguments: 
'	strStringToConvert as string - the string to remove characters from
'Description: Convert characters for the provided string. The point of
'	this function is to enable the string not to crash javascript.
'Output: the stripped string
'Special Instructions:NA
'***************************************************************
Function UTL_CharsSafeForJavascript(strStringToConvert)
	
	strStringToConvert = Replace(strStringToConvert,"'","\'")
	strStringToConvert = Replace(strStringToConvert,Chr(34),"\" & chr(34))

	UTL_CharsSafeForJavascript = strStringToConvert
End Function

'***************************************************************
'Function Name: UTL_Max
'Arguments: 
'	arrData as array - single column array 
'Description: Finds the maximum value from a single dimension array of numbers
'Output: the greatest number from the array
'Assumption: only a single dimensional array is passed to the function
'Special Instructions:NA
'***************************************************************
Function UTL_Max(arrData)
	dim intCount
	dim intStoreMax
	
	intStoreMax = arrData(0)
	for intCount = Lbound(arrData,1) + 1 to Ubound(arrData,1)
		if (arrData(intCount) > intStoreMax) then
			intStoreMax = arrData(intCount)
		end if
	next
	if (intStoreMax = "") then
		UTL_Max = 0
	else
		UTL_Max = intStoreMax
	end if
End Function

'***************************************************************
'Function Name: UTL_Min
'Arguments: 
'	arrData as array - single column array
'Description: Finds the Minimum value from a single dimension array of numbers
'Output: The lowest number in the array
'Assumption: only a single dimensional array is passed to the function
'Special Instructions:NA
'***************************************************************
Function UTL_Min(arrData)
	dim intCount
	dim intStoreMin
	
	intStoreMin = arrData(0)
	for intCount = Lbound(arrData, 1) + 1 to Ubound(arrData,1)
		if (arrData(intCount) < intStoreMin) then
			intStoreMin = arrData(intCount)
		end if
	next
	if (intStoreMin = "") then
		UTL_Min = 0
	else
		UTL_Min = intStoreMin
	end if
End Function

'***************************************************************
'Function Name: UTL_Truncate
'Arguments: 
'	strInitialString as string - string representing a floating point number
'	intNumberOfDecimalPlaces as integer - number of decimal places to display
'Description: This function takes a string representation of a floating point
'number and truncates the decimal places bassed upon the intNumberOfDecimalPlaces
'argument.
'Output: String represnting new Truncated number.
'Special Instructions: NA
'***************************************************************
Function UTL_Truncate(strInitialString,intNumberOfDecimalPlaces)
	dim intDecimalPosition
	
	intDecimalPosition = 0
	intDecimalPosition = InStr(1,strInitialString,".")
	if intDecimalPosition <> 0 then 'if there is a decimal point in the string
		if (intNumberOfDecimalPlaces = 0) then
			UTL_Truncate = Mid(strInitialString,1,(intDecimalPosition - 1))
		else
			UTL_Truncate = Mid(strInitialString,1,(intDecimalPosition + intNumberOfDecimalPlaces))
		end if
	else
		UTL_Truncate = strInitialString
	end if
End Function

'***************************************************************
'Function Name: UTL_CreateNumericList
'Arguments: 
'	intStartValue as integer - number to start form
'	intEndValue as integer - number to end at
'	chrSeparator as character - separator between each number
'	bolLeadingZeroas boolean - specifies if numbers < 10 should have a
'	leading zero
'Description: This function takes a start number and an end number
'and creates a string of each number inbetween, using the chrSeparator
'variable as a separator. E.g. intStartValue = 1, intEndValue = 10,
'chrSeparator = "," then the returned string would be "1,2,3,4,5,6,7,8,9,10"
'Output: String containing the required numbers
'Special Instructions: NA - used mainly for the creation of option lists.
'***************************************************************
Function UTL_CreateNumericList(intStartValue,intEndValue,chrSeparator,bolLeadingZero)
	Dim intLoopCount, strNumericList, strCurrentNum
	
	strNumericList = ""
	
	for intLoopCount = intStartValue to intEndValue
		strCurrentNum = CStr(intLoopCount)
		if bolLeadingZero = True then
			if CInt(intLoopCount) < 10 then
				strCurrentNum = "0" & intLoopCount
			end if
		end if
		
		if CInt(intLoopCount) = intEndValue then
			strNumericList = strNumericList & strCurrentNum
		else
			strNumericList = strNumericList &  strCurrentNum & chrSeparator
		end if
	next
	
	UTL_CreateNumericList = strNumericList
	
End Function

'***************************************************************
'Function Name: UTL_FlipArray
'Arguments:
'	arrToFlip as array - array to be flipped.
'Description: This function flips an array by transposing the rows
'and columns from one array to another. The main purpose of the function
'is to flip arrays received from the ADO recordset function GetRows().
'The GetRows() function returns an array but in the wrong order.
'Output: An array of flipped values.
'Special Instructions: Must receive a two dimensional array only.
'***************************************************************
Function UTL_FlipArray(arrToFlip)
	Dim arrFlipped, intRowCount, intColumnCount

	ReDim arrFlipped(Ubound(arrToFlip,2),Ubound(arrToFlip,1))
	
	for intRowCount = lbound(arrToFlip,2) to Ubound(arrToFlip,2)
		for intColumnCount = lbound(arrToFlip,1) to Ubound(arrToFlip,1)
			arrFlipped(intRowCount,intColumnCount) = arrToFlip(intColumnCount,intRowCount)
		next
	next

	UTL_FlipArray = arrFlipped

End Function

'***************************************************************
'Function Name: UTL_GetFileName
'Arguments:
'	strFullPathName as string - full path name to a file.
'Description: This function retrieves the name of a file from the
'supplied full path name.
'Output: A string representing the file name
'Special Instructions: Assumes from a Windows environment using \
'as directory separators.
'***************************************************************
Function UTL_GetFileName(strFullPathName)
	Dim intFieldCount, strFileName, arrPathDetails
	
	arrPathDetails = Split(strFullPathName,"\")
	intFieldCount = UBound(arrPathDetails)
	strFileName = arrPathDetails(intFieldCount)
	
	UTL_GetFileName = strFileName

End Function

'***************************************************************
'Function Name: UTL_GetListDifferences
'Arguments:
'	strList1 as string - a list of values.
'	strList2 as string - a list of values to compare against
'	bolQuoteOutput as boolean - specifies if the returned list should
'	be quoted.
'Description: Determines which elements in list one are not in list
'two.
'Output: A delimited string containing those values in list one
'but not in list two.
'Special Instructions: NA
'***************************************************************
Function UTL_GetListDifferences(strList1,strList2,charDelimiter,bolQuoteOutput)
	Dim arrList1, arrList2, strTempString, bolValueFound
	Dim strCurrentList1Value, strCurrentList2Value

	arrList1 = Split(strList1,charDelimiter)
	arrList2 = Split(strList2,charDelimiter)
	strTempString = ""

	for each strCurrentList1Value in arrList1
		bolValueFound = False
		for each strCurrentList2Value in arrList2
			if Cstr(strCurrentList1Value) = Cstr(strCurrentList2Value) then
				bolValueFound = True
				exit for
			end if
		next
		
		if bolValueFound = False then
			if bolQuoteOutput = False then
				strTempString = strTempString & strCurrentList1Value & charDelimiter
			else
				strTempString = strTempString & "'" & strCurrentList1Value & "'" & charDelimiter
			end if
		end if
	next
	
	if strTempString <> "" then
		strTempString = Left(strTempString,Len(strTempString) -1) 'removes the last delimiter character
	end if

	UTL_GetListDifferences = strTempString
	
End Function

'***************************************************************
'Function Name: UTL_QuoteArray
'Arguments:
'	arrToQuote as array - array of values to put quotes arround.
'Description: This function puts single quotes arround each element
'of the supplied 1 dimension array. The main purpose to prepare
'values that are to be used in an SQL query - especially an IN
'clause.
'Output: The same array with each element quoted.
'Special Instructions: NA
'***************************************************************
Function UTL_QuoteArray(arrToQuote)
	Dim strTempString, intMaxElements, intLoopCount
	
	intMaxElements = UBound(arrToQuote)
	
	for intLoopCount = 0 to intMaxElements
		arrToQuote(intLoopCount) = "'" & arrToQuote(intLoopCount) &  "'"
	next

End Function


'P1********************************************************************************
'Procedure		:UTL_OptionActive
'Arguments		:intBitString as integer - bit string value
'				:intBitMask as integer - bit mask value
'Purpose		:Compares the supplied bit string with the supplied bit mask value
'				:and returns true if found, otherwise false.
'Output			:Boolean.
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_OptionActive(intBitString, intBitMask)
	if cint(intBitString) and cint(intBitMask) then
		UTL_OptionActive = true
	else
		UTL_OptionActive = false
	end if
end function

'P1********************************************************************************
'Procedure		:UTL_StringToBitMask
'Arguments		:strString as string - delimited string of
'				:strDelimiter as string - value delimiter
'Purpose		:Adds the bit mask compatible values of a delimited string
'				:and returns a bit mask.
'Output			:Returns a valid bit mask.  If strString is nothing then 0 is returned.
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_StringToBitMask(strString, strDelimiter)
	dim arrData
	dim intI
	dim intBitMask
	
	intBitMask = 0
	if len(strString) > 0 then
		arrData = split(strString, strDelimiter)
		for intI = lbound(arrData) to ubound(arrData)
			intBitMask = intBitMask + cint(trim(arrData(intI)))
		next
		UTL_StringToBitMask = intBitMask
	else
		UTL_StringToBitMask = 0
	end if
end function

'P1********************************************************************************
'Procedure		:UTL_ModeChecked
'Arguments		:intBitString as integer - value
'				:intBitMask as integer - bit mask
'				:strReturnValue as string - value to return if true
'Purpose		:Checks if the supplied bit mask exists in the bit string.  If true then
'				:strReturnValue is returned.
'Output			:String
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_ModeChecked(intBitString, intBitMask, strReturnValue)
	if UTL_OptionActive(intBitString, intBitMask) = true then
		UTL_ModeChecked = strReturnValue
	end if
end function


'P1**********************************************************************************
'Procedure		:UTL_DelimitRecSetFld
'Arguments		:recSet as ADODB.Recordset - data laiden recordset
'				:strList as string - variable to assign list to
'				:strFldName as string - field of recSet to delimit
'				:lngMaxRows as long - number of rows
'				:strDelimiter as string - delimiter
'Purpose		:Iterates though a recordset to concatentate a field from recordset
'				:into a delimited string
'Output			:True if successful, if not data then false.
'Dependancies	:NA
'Dependants		:NA
'Special Info	:The list position is zero based
'P2**********************************************************************************
function UTL_DelimitRecSetFld(recSet, strList, strFldName, lngMaxRows, strDelimiter)
	dim strRetVal
	dim intRecCount
	
	intRecCount = 0
	if IsObject(recSet) then
		do until recSet.EOF
			UTL_AddListItem strList, Empty, recSet(strFldName), strDelimiter
			recSet.movenext
		loop
		recSet.MoveFirst
		UTL_DelimitRecSetFld = True
	else
		UTL_DelimitRecSetFld = false
	end if
end function


'P1**********************************************************************************
'Procedure		:UTL_AddListItem
'Arguments		:strList as string - a valid list
'				:lngNewPosition as long - list position to add new itm OPTIONAL
'				:strNewVal as string - value of new list item
'				:strDelimiter as string - list delimiter
'Purpose		:Adds a new list item at the specified position, with
'				:the specified value.
'Output			:Boolean, true if new item added.
'Dependancies	:NA
'Dependants		:NA
'Special Info	:If lngNewPosition exceeds the upper bounds of the new list,
'				:then it is just added to the end of the new list, likewise
'				:if lngNewPosition is less than the lower bounds.  If strList
'				:contains no values then a new list is created.
'				:IMPORTANT - the list position is zero based.
'				:lngNewPosition can be set to EMPTY to add the new list item
'				:to the end of the list.
'P2**********************************************************************************
Sub UTL_AddListItem(strList, lngNewPosition, strNewVal, strDelimiter)
	Dim arrNew
	Dim intIndex
	Dim intAltIndex
	Dim arr
	
	intAltIndex = 0
	arr = Split(strList, strDelimiter)
	
	Redim arrNew(Ubound(arr) + 1)
	
	if len(lngNewPosition) > 0 then
		If lngNewPosition > UBound(arrNew) Then
			lngNewPosition = UBound(arrNew)
		End if
	
		If lngNewPosition < LBound(arrNew) Then
			lngNewPosition = LBound(arrNew)
		End if
	end if
		
	If IsArray(arr) Then
		if len(lngNewPosition) > 0 then
			For intIndex = Lbound(arrNew) to Ubound(arrNew)
				If lngNewPosition = intIndex Then
					arrNew(intIndex) = strNewVal
				Else		
					arrNew(intIndex) = arr(intAltIndex)					
					intAltIndex = intAltIndex + 1
				End if
			Next
			strList = Join(arrNew, strDelimiter)
		else
			For intIndex = Lbound(arrNew) to Ubound(arrNew)
				If clng(lngNewPosition) = clng(intIndex) Then
					arrNew(intIndex) = strNewVal
				Else		
					arrNew(intIndex) = arr(intAltIndex)					
					intAltIndex = intAltIndex + 1
				End if
			Next
			strList = Join(arrNew, strDelimiter)			
		end if
	Else
		strList = strNewVal
	End if
end sub


'P1**********************************************************************************
'Procedure		:UTL_FindListItem
'Arguments		:strList as string - a valid list
'				:strValToFind as string - value to find in list
'				:strDelimiter as string - list delimiter
'Purpose		:This function finds the first occurrence of strValToFind and
'				:returns its position in the supplied list.
'Output			:Long - position of first occurrence of value in list
'Dependancies	:NA
'Dependants		:NA
'Special Info	:The list position is zero based
'P2**********************************************************************************
Function UTL_FindListItem(strList, strValToFind, strDelimiter)
	Dim arr
	Dim intIndex
	Dim lngPositionFound
	
	lngPositionFound = -1
	arr = Split(strList, strDelimiter)
	
	
	For intIndex = LBound(arr) to Ubound(arr)

		If trim(cstr(strValToFind)) = trim(cstr(arr(intIndex))) Then
			lngPositionFound = intIndex
			Exit For
		End if
	Next
	UTL_FindListItem = lngPositionFound

End Function


'P1**********************************************************************************
'Procedure		:UTL_GetListItem
'Arguments		:strList as string - a list of values 
'				:lngListPosition as long - list position value to retrieve
'				:strListDelimiter as string - list delimiter
'Purpose		:To return the value specified at specific position in a list
'				:indicated by lngListPostiion
'Output			:String - value at position in list.  If list position does
'				:not exist -1 is returned
'Dependancies	:NA
'Dependants		:NA
'Special Info	:The list position is zero based.  Also note that strList is a list
'				:of values that use the same delimiter as specified by strListDelimiter
'P2**********************************************************************************
Function UTL_GetListItem(strList, lngListPosition, strListDelimiter)
	Dim arr
	
	arr = Split(strList, strListDelimiter)
	If IsArray(arr) Then		
		If lngListPosition < 0 OR lngListPosition > UBound(arr) Then	
			UTL_GetListItem = -1
		Else
			UTL_GetListItem = arr(lngListPosition)
		End if
	Else
		UTL_GetListItem = -1
	End if
End Function


'P1**********************************************************************************
'Procedure      :UTL_ListLength
'Arguments      :strList as string - a valid list
'               :strDelimiter as string - list delimiter
'Purpose        :Returns the number of items in the list
'Output         :Long - length of list
'Dependancies   :NA
'Dependants     :NA
'Special Info   :NA
'P2**********************************************************************************
Function UTL_ListLength(strList, strDelimiter)
    Dim arr
    Dim intIndex
    Dim varRetVal
    If Len(strList) > 0 Then
        arr = Split(strList, strDelimiter)
        If IsArray(arr) Then
            UTL_ListLength = UBound(arr)
        Else
            UTL_ListLength = 0
        End If
    Else
        UTL_ListLength = -1
    End If

End Function

'P1**********************************************************************************
'Procedure      :UTL_NewGUID
'Arguments      :NA
'Purpose        :Returns a new GUID
'Output         :String
'Dependancies   :NA
'Dependants     :NA
'Special Info   :NA
'P2**********************************************************************************
function UTL_NewGUID()
	dim guid
	guid = server.createobject("scriptlet.typelib").guid
	guid = ucase(guid)
	guid = Mid(guid, 1, Len(guid) - 2)
	UTL_NewGUID = guid
end function



'P1**********************************************************************************
'Procedure      :UTL_StringRepeat
'Arguments      :strString as string
'				:intNumber as integer
'Purpose        :Returns the value of strString the number of times specified by intNumber
'Output         :String
'Dependancies   :NA
'Dependants     :NA
'Special Info   :NA
'P2**********************************************************************************
Function UTL_StringRepeat(strString, intNumber)
    Dim intI
    Dim str
    str = ""
    For intI = 1 To intNumber
        str = str & strString
    Next
    UTL_StringRepeat = str
End Function


'P1**********************************************************************************
'Procedure      :UTL_ArrayToString
'Arguments      :arr, strFldDelim, strRecDelim, strBlankReplace
'Purpose        :
'Output         :
'Dependancies   :NA
'Dependants     :NA
'Special Info   :NA
'P2**********************************************************************************
function UTL_ArrayToString(arr, strFldDelim, strRecDelim, strBlankReplace)
	dim intR
	dim intC
	dim str
	dim strRow
	dim strTmp

	str = ""
	for intR = lbound(arr, 1) to ubound(arr, 1)
		
		if str <> "" then
			str = str & strRecDelim
		end if
		
		strRow = ""
		
		for intC = lbound(arr, 2) to ubound(arr, 2)
			if strRow <> "" then
				strRow = strRow & strFldDelim
			end if
		
			strTmp = arr(intR, intC)
		
			if len(trim(strTmp)) = 0 then
				strTmp = strBlankReplace
			end if
		
			strRow = strRow & strTmp
		
		next
		
		str = str & strRow
	
	next
	UTL_ArrayToString = str
end function


'P1**********************************************************************************
'Procedure		:UTL_ReplaceSubString
'Arguments:		:strExp as string - expression to evaluate
'				:strFind as string - string to replace
'				:strReplace as string - text to replace found ocurrences with
'				:intCompareType as integar - 0 = boolean compare, 1 = textual compare
'
'Purpose		:To replace an occurence of a substring within a string with a new 
'				:string.
'Output			:The processed string
'Dependancies	:NA
'Dependants		:NA
'Special Info	:Returns following values with:
'				:strExp = null - returns strReplaceText
'				:strExp = empty - returns strReplaceText
'				:strFind = zero length or empty - returns copy of strExp
'				:strReplace = zero length - copy of strExp with all occurrences of strFind removed.
'				:If strReplace is supplied as null, then an error will result.
'P2**********************************************************************************
function UTL_ReplaceSubString(strExp, strFind, strReplace, intCompareType)
	if isnull(strExp) = true or strExp = empty then
		UTL_ReplaceSubString = strReplace
	else
		UTL_ReplaceSubString = Replace(strExp, strFind, strReplace, 1, -1,intCompareType)
	end if
end function


function UTL_SEArrayAsTable(arr)
	dim str
	dim intI
	str = "<table border=1><tr>"
	for intI = lbound(arr) to ubound(arr)
		str = str & "<td>" & arr(intI) & "</td>"
	next
	str = str & "</tr></table>"
	UTL_SEArrayAsTable = str
end function


'********************************************************************
'Author: Kym Oberauer - Digital Zoo Pty Ltd
'Date: 4/11/1998 Version 1.0
'
'Library: standard.asp
'Description: Library of util functions etc...
'********************************************************************

	'********************************************************************
	'Function: UTL_DisplayMessage
	'Description: Simply concatinates a string to be displayed to the user.
	'********************************************************************
	Sub UTL_DisplayMessage(strMessage, strMessageString,strMessageColour)
		strMessageString = strMessageString & "<span style='color:" & strMessageColour & "'>" & strMessage & "</span> <br>"
	End Sub
	
	'********************************************************************
	'Function:UTL_JoinStrings
	'Description: Simply concatinates a string. This function is not used
	'to display info to the user but for html building.
	'********************************************************************
	Function UTL_JoinStrings(strOne, strTwo)
		strOne = strOne & strTwo
	End Function

	'********************************************************************
	'Function: UTL_QuoteString
	'Description: Adds quotation marks to a string. Used to ensure the 
	'sytax of arguments passed to stored procedures.
	'********************************************************************
	Function UTL_QuoteString(strToQuote)
		UTL_QuoteString = """" & strToQuote & """"
	End Function
	
	'********************************************************************
	'Function: UTL_FormatURLArg
	'Description: This function replaces spaces with '%20'. This is
	'necessary because Netscape does not like space characters in URL args.
	'********************************************************************
	Function UTL_FormatURLArg(strUrlArg)
		Dim arrURLArg,strFormatedArg, strRepVal, strElement
		
		strRepVal = "%20" ' HTTP space character
		arrURLArg = split(strUrlArg," ")
		
		for each strElement in arrURLArg
			strFormatedArg = strFormatedArg & strElement & strRepVal
		next

		UTL_FormatURLArg = strFormatedArg
	End Function

	
'***************************************************************
'Function Name: UTL_NullToZero
'Arguments: 
'	varElement as variant
'Description: This function tests to see if value is either null or an empty string
'Output: Returns the unchanged argument if the value is not null or 0 otherwise
'Special Instructions: Tests for null and empty strings and converts them to 0
'***************************************************************
Function UTL_NullToZero(varElement)
	if (varElement = "") then
		UTL_NullToZero = 0
	elseif (isNull(varElement) = True) then
		UTL_NullToZero = 0
	elseif (isEmpty(varElement) = True) then
		UTL_NullToZero = 0
	else
		UTL_NullToZero = varElement
	end If
End Function



'P1********************************************************************************
'Procedure		:UTL_ArrayGetRowIndex
'Arguments		:arrData as array - a 2 dimensional array of data
'				:intIndexValue as integer - search value
'				:intIndexField as integer - element of array to match intIndexValue to
'Purpose		:Searches the specified field (intIndexField) and tries to find a match
'				:to intIndexValue.  If a match is found, the row index is returned, other
'				:false is returned.
'Output			:Boolean.
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_ArrayGetRowIndex(arrData, intIndexValue, intIndexField)
	dim intR
	if not isarray(arrData) then
		Response.Write "<BR>A valid array must be supplied"
	end if
	for intR = lbound(arrData, 1) to ubound(arrData, 1)
		if cstr(ucase(arrData(intR, intIndexField))) = cstr(ucase(intIndexValue)) then	
			UTL_ArrayGetRowIndex = intR
			exit function
		end if
	next
	UTL_ArrayGetRowIndex = -1
end function

'P1********************************************************************************
'Procedure		:UTL_ArrayPerformFieldOperation
'Arguments		:arrData as array - array of data
'				:intSearchValue - value to match
'				:intIndexField - index to perform match on
'				:intOperationField as integer - field to perform operation on
'				:intStep as integer - amount to perform operation
'				:strOperator as string - either BATCHUPLOAD_DECREMENT or BATCHUPLOAD_INCREMENT
'Purpose		:Searches for the specified field in the array and performs the specified operation on
'				:it.  Returns the resulting value. If element could not be found then false is returned.
'Output			:Boolean.
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_ArrayPerformFieldOperation(arrData, intIndex, intOperationField, intStep, strOperator)
	arrData(intIndex, intOperationField) = eval(clng(arrData(intIndex, intOperationField)) & " " & strOperator & " " & intStep)
	UTL_ArrayPerformFieldOperation = arrData(intIndex, intOperationField)
end function


'P1********************************************************************************
'Procedure		:UTL_FormatTitleCase
'Arguments		:strData as string
'Purpose		:NA
'Output			:Integer
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_FormatTitleCase(strData)
	dim strFirstLetter
	dim strString
	if len(strData) > 0 then	
		strFirstLetter = ucase(left(strData, 1))
		strString = lcase(right(strData, len(strData) - 1))
		UTL_FormatTitleCase = strFirstLetter & strString
	else
		UTL_FormatTitleCase = strData
	end if
end function

    '********************************************************************
	'Function: UTL_removeDupsArray
	'Arguments:
	'	array as string - comma delimited array
	'Description: removes duplcate records
	'Output:    returns string of non duplicate
	'Special Instructions:NA
	'********************************************************************
    Function UTL_removeDupsArray(sList)
    
        Dim sNewList, aList, maxItems, x
        
        aList = split(sList,",")
        maxItems = UBound(aList)
        
        For x = 0 To maxItems
            If InStr(sNewList,(aList(x) & ",")) <= 0 Then
                sNewList = sNewList & aList(x) & ","
            End If
        Next
        
        
        if Len(sNewList) > 0 then
			UTL_removeDupsArray = Left(sNewList,Len(sNewList)-1)
		end if
    End Function


'P1********************************************************************************
'Procedure		:UTL_GetSalt
'Arguments		:intSafID as integer
'Purpose		:Return the salt ID
'Output			:Integer
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_GetSalt(intSafID)
	
	
	Dim objXmlHttp
    Dim strHTML


    ' This is the server safe version from MSXML3.
    Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")


    ' We get teh Salt
    objXmlHttp.open "GET", "http://" & Request.ServerVariables("HTTP_HOST") & "/tools/EncryptService/Default.aspx?safid=" & intSafID, False
    

    ' Send it on it's merry way.
    objXmlHttp.send

    ' Print out the request status:
    ' Response.Write "Status: " & objXmlHttp.status & " " _ & objXmlHttp.statusText & "<br />"

    ' Get the text of the response.
    ' This object is designed to deal with XML so it also has the
    ' following properties: responseBody, responseStream, and
    ' responseXML.  We just want the text so I use:
    if objXmlHttp.Status <> 200 Then
        UTL_GetSalt = ""
    else
        UTL_GetSalt = objXmlHttp.responseText
    end if


    ' Destroy the object after user
    set objXmlHttp = nothing
	
end function

'P1********************************************************************************
'Procedure		:UTL_EncryptValue
'Arguments		:strValue as varchar
'                strSalt as varchar
'Purpose		:Return the encrypted value
'Output			:Integer
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_EncryptValue(strValue, strSalt)
on error resume next

    dim objEncrypt
    dim Context
    dim Key
    dim Blob
    dim strReturnString
    
    strReturnString = ""
    
    'What is the point if the string is blank
    if len(trim(strValue)) > 0 then
    
        set objEncrypt = Server.CreateObject("Persits.CryptoManager")
        Set Context = objEncrypt.OpenContext("", True )
        
        ' generate key from password
	    Set Key = Context.GenerateKeyFromPassword(strSalt)
    	
	    ' Encrypt the value
	    Set Blob = Key.EncryptText(strValue)
    	
	    ' CryptoBlob presents data in three formats: Hex, Base64 and Binary.
	    ' We return the encrypted value
	    strReturnString = Blob.Base64
    	
	    ' If the string is blank then we return the original string
	    if len(trim(strReturnString)) = 0 then
	        strReturnString = trim(strValue)
	    end if
	
	end if
	
	UTL_EncryptValue = strReturnString
	
	
	set Blob = nothing
	set Key = nothing
	set Context = nothing
	set objEncrypt = nothing


end function

'P1********************************************************************************
'Procedure		:UTL_DecryptValue
'Arguments		:strValue as varchar
'                strSalt as varchar
'Purpose		:Return the decrypted value
'Output			:Integer
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_DecryptValue(strValue, strSalt)
on error resume next

    dim objEncrypt
    dim Context
    dim Key
    dim Blob
    dim strReturnString
    
    strReturnString = ""
    
    'What is the point if the string is blank
    if len(trim(strValue)) > 0 then
    
        set objEncrypt = Server.CreateObject("Persits.CryptoManager")
        Set Context = objEncrypt.OpenContext("", True )
        
        
        ' generate key from password
	    Set Key = Context.GenerateKeyFromPassword(strSalt)
        
        ' create an empty blob object and fill it with data
	    Set Blob = objEncrypt.CreateBlob
	    Blob.Base64 = strValue
    	
	    strReturnString = Key.DecryptText(Blob)	
    	
    	
	    ' If the string is blank then we return the original string
	    if len(trim(strReturnString)) = 0 then
	        strReturnString = trim(strValue)
	    end if
	end if
	
	UTL_DecryptValue = strReturnString
	
	set Blob = nothing
	set Key = nothing
	set Context = nothing
	set objEncrypt = nothing


end function

'********************************************************************
'Function:
'	UTL_GetEncryptionValue
'Description:
'       Used to get which field requires to have encryption on
'********************************************************************
function UTL_GetEncryptionValue(recSet)		

    ' Default it to off
    Session("InstUsernameEncrypt")   = 0
	Session("InstPasswordEncrypt")   = 0
	Session("InstEmailEncrypt")      = 0
	Session("InstSSOEncrypt")      = 0
	Session("InstFirstNameEncrypt")      = 0
	Session("InstLastNameEncrypt")      = 0


    if NOT DB_RecSetIsEmpty(recSet) then
        Session("InstUsernameEncrypt")   = Trim(recSet("encrypt_username"))
		Session("InstPasswordEncrypt")   = Trim(recSet("encrypt_password"))
		Session("InstEmailEncrypt")      = Trim(recSet("encrypt_email"))
		Session("InstSSOEncrypt")        = Trim(recSet("sso_encrypt"))
		Session("InstFirstNameEncrypt")      = Trim(recSet("encrypt_first_name"))
    	Session("InstLastNameEncrypt")      = Trim(recSet("encrypt_last_name"))
    
    end if
    
    if trim(Session("InstUsernameEncrypt")) = "1" or trim(Session("InstPasswordEncrypt")) = "1" or trim(Session("InstEmailEncrypt")) = "1" or trim(Session("InstSSOEncrypt")) = "1" _
            or trim(Session("InstFirstNameEncrypt")) = "1" or trim(Session("InstLastNameEncrypt")) = "1" then
        ' We check if we have the salt yet
        if len(trim(session("client_salt"))) = 0 then
            session("client_salt") = UTL_GetSalt(SAF_ID)
        end if
    end if

end function


'********************************************************************
'Function:
'	UTL_GetEncryptionValue
'Description:
'       Used to get which field requires to have encryption on
'********************************************************************
function UTL_GetLineManagerAlias(recSet)		

    Session("InstLineManagerAlias")   = Trim(recSet("line_manager"))
	
end function

'P1********************************************************************************
'Procedure		:CacheBusterQuerystring
'Purpose		:returns a querystring varialbe with a timestamp, for CacheBusting
'Output			:string
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function CacheBusterQuerystring()  
 	Randomize()
 	CacheBusterQuerystring = "cs=" & Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now())  & Second(Now()) & "_" & cint(Rnd*1000)
end function

'P1********************************************************************************
'Procedure		:UTL_validateLogin
'Purpose		:This will validate if the user is still login based on the session
'Output			:string
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
function UTL_validateLogin()
	if len(trim(Session("UsrUserID"))) = 0 or Session("UsrLoginRole") = 0 or Session("UsrUserName") = "" or Session("UsrPassword") = "" then	
        Session.Abandon
		Session.Contents.RemoveAll()		
		response.Redirect("default.asp")
    end if
end function

function UTL_validateAdminLogin()
	if len(trim(Session("AdminUserID"))) = 0 or Session("AdminLoginRole") = 0 or Session("AdminUserName") = "" or Session("AdminPassword") = "" then	
        Session.Abandon
		Session.Contents.RemoveAll()		
		response.Redirect("default.asp")
    end if
end function

function UTL_validateRoadshowLogin()
	if len(trim(Session("UsrUserID"))) = 0 or Session("UsrLoginRole") = 3 or Session("UsrUserName") = "" or Session("UsrPassword") = "" then	
        Session.Abandon
		Session.Contents.RemoveAll()		
		response.Redirect("default.asp")
    end if
end function
%>