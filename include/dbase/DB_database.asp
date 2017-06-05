<%
'H1**********************************************************************************
'Copyright			:Digital Zoo Pty Ltd - 2003
'Library			:DB_database.asp
'Creation Date		:22 Aug 2003
'Version			:1.0
'Author(s)			:Indra Hopleythompson
'
'Library Purpose	:Scripts that are used several times relating to databases
'
'Procedures			:sub DB_closeObject(objADO)
'					:function DB_RecSetIsEmpty(objRecSet)
'					:Function DB_RecSetUse(varUniqueParam, varParamDataType, strSpName, strUseRecSetFunction, strFnErrorMessage)
'					:Function DB_RecSetReturn(varUniqueParam, varParamDataType, strSpName, varResult)
'					:Sub DB_RecSetToArray(arrData, recSet)
'					:function DB_DeleteRecordByID(intUniqueID, strIdDataType, strSpName)
'					:Function DB_NullToEmpty(varElement)
'					:Function DB_ParamLength(strElement)
'
'Special Info		:NA
'
'Revision History	:################################################################
'H2**********************************************************************************
Const DB_PARAM_DELIM	= ","

'P1********************************************************************************
'Procedure		:DB_closeObject
'Arguments		:objADO as any ADO object - 
'Purpose		:check the parameter is an ADO object, then close and set to 
'				:nothing as appropriate
'Output			:NA
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
Sub DB_closeObject(objADO)

	If isObject(objADO) Then
		
		On Error Resume Next
			objADO.Close
		On Error Goto 0
		
		Set objADO = nothing
	
	End If

End Sub

'P1********************************************************************************
'Procedure		:DB_RecSetIsEmpty
'Arguments		:objRecSet as any ADO recordset object - 
'Purpose		:checks if the recordset contains records
'Output			:boolean
'Dependancies	:NA
'Special Info	:NA
'P2********************************************************************************
Function DB_RecSetIsEmpty(objRecSet)

   DB_RecSetIsEmpty = True

   If typename(objRecSet) = "Recordset" then
      'If (objRecSet.ActiveConnection.State = 1) and (objRecSet.State = 1) then
      If objRecSet.State = 1 then
            DB_RecSetIsEmpty = (objRecSet.eof or objRecSet.bof)
         End if
   End If

End Function







'********************************************************************
'Function: 
'	DB_RecSetUse
'Author:
'	Indra, June 06
'Arguments:
'	varUniqueParam			- comma delimited param(s) that will call the single record
'	varParamDataType		- comma delimited param(s) data type
'	strSpName				- name of the Stored Procedure
'	strUseRecSetFunction	- the name of the function that utilises the 
'							produced recordset
'	strFnErrorMessage		- sql error message
'Purpose:
'	Retrieves a recordset from the database with the supplied Params, then uses it in the 
'	provided function call.
'Requires:
'	sqlerror.asp - to produce the error message
'Output:
'	call to function strUseRecSetFunction(recset)
'	boolean result
'********************************************************************
Function DB_RecSetUse(varUniqueParam, varParamDataType, strSpName, strUseRecSetFunction, strFnErrorMessage)
	Dim connObj
	Dim cmdObj
	Dim paraObj
	Dim recSet
	Dim arrParamInput
	Dim arrParamType
	Dim intCounter
	

	'Open the connection
	Set connObj = Server.CreateObject("ADODB.Connection")
	With connObj
		.ConnectionString = Session("ConnectionString")
		.CursorLocation = adUseClient
		.Open
	End With

	'Creating the command object
	Set cmdObj = Server.CreateObject("ADODB.Command")
	With cmdObj
		.ActiveConnection = connObj
		.CommandText = strSpName
		.CommandType = AdCmdStoredProc
	End With

	'establish the parameters & types
	If trim(varUniqueParam) <> "" AND NOT isnull(varUniqueParam) Then
		arrParamInput = Split(varUniqueParam, DB_PARAM_DELIM)
		arrParamType = Split(varParamDataType, DB_PARAM_DELIM)

		'loop through the input arrays and append parameters
		For intCounter = lBound(arrParamInput) To uBound(arrParamInput)

			'Creating parameters
			Select Case arrParamType(intCounter)

				Case "varchar"
					Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,DB_ParamLength(arrParamInput(intCounter)),DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj
					
				Case "char"
					Set paraObj = cmdObj.CreateParameter(,AdChar,AdParamInput,6,DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj

				Case Else
					Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4,DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj

			End Select
		
		Next
	End If

	'DEBUG_dbase(cmdObj)

	'Execute Command
	on error resume next
	set recSet = cmdObj.Execute
	on error goto 0

	'DEBUG_dbase(recSet)

	'Check for errors
	if CheckForSQLError(connObj,"Get",strFnErrorMessage) = TRUE then
		DB_RecSetUse = FALSE
	else
		Execute(strUseRecSetFunction & "(recSet)") 
		DB_RecSetUse = TRUE
	end if

	'Destroy all the created objects
	Call DB_closeObject(paraObj)
	Call DB_closeObject(cmdObj)
	Call DB_closeObject(recSet)
	Call DB_closeObject(connObj)
	
End Function

'********************************************************************
'Function: 
'	DB_RecSetReturn
'Author:
'	Indra, June 06
'Arguments:
'	varUniqueParam		- comma delimited param(s) that will call the single record
'	varParamDataType	- comma delimited param(s) data type
'	strSpName			- name of the Stored Procedure
'	varResult			- ADO recordset object or sql error message
'Purpose:
'	Retrieves a recordset from the database with the supplied Params, then 
'	returns it to the calling script (or the error message).
'	NB: recSet needs to be cleaned up at the caller end
'Requires:
'	sqlerror.asp - to produce the error message
'Output:
'	returns boolean result
'	populates recset or error message
'********************************************************************
Function DB_RecSetReturn(varUniqueParam, varParamDataType, strSpName, varResult)
	Dim connObj
	Dim cmdObj
	Dim paraObj
	dim recSet
	Dim arrParamInput
	Dim arrParamType
	Dim intCounter
	

	'Open the connection
	Set connObj = Server.CreateObject("ADODB.Connection")
	With connObj
		.ConnectionString = Session("ConnectionString")
		.CursorLocation = adUseClient
		.Open
	End With

	'Creating the command object
	Set cmdObj = Server.CreateObject("ADODB.Command")
	With cmdObj
		.ActiveConnection	= connObj
		.CommandText		= strSpName
		.CommandType		= AdCmdStoredProc
	End With

	'establish the parameters & types
	If trim(varUniqueParam) <> "" AND NOT isnull(varUniqueParam) Then
		arrParamInput = Split(varUniqueParam,DB_PARAM_DELIM)
		arrParamType = Split(varParamDataType,DB_PARAM_DELIM)

		'loop through the input arrays and append parameters
		For intCounter = lBound(arrParamInput) To uBound(arrParamInput)

			'Creating parameters
			Select Case arrParamType(intCounter)

				Case "varchar"
					Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,DB_ParamLength(arrParamInput(intCounter)),DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj
					
				Case "char"
					Set paraObj = cmdObj.CreateParameter(,AdChar,AdParamInput,6,DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj

				Case Else
					Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4,DB_NullToEmpty(arrParamInput(intCounter)))
					cmdObj.Parameters.Append paraObj

			End Select
		
		Next
	End If

	'DEBUG_dbase(cmdObj)

	'Execute Command
	on error resume next
	set recSet = cmdObj.Execute
	on error goto 0
	

	'Check for errors
	if CheckForSQLError(connObj,"Get",varResult) = TRUE then
		DB_RecSetReturn = false
	else
		'DEBUG_dbase(recSet)

		set recSet.ActiveConnection = nothing
		set varResult = recSet
		DB_RecSetReturn = true
	end if

	'Destroy all the created objects
	Call DB_closeObject(paraObj)
	Call DB_closeObject(cmdObj)
	'NB: recSet needs to be cleaned up at the caller end - DB_closeObject(recSet)
	Call DB_closeObject(connObj)
	
End Function



'********************************************************************
'Function: 
'	DB_RecSetToArray
'Author:
'	Indra, Oct 05
'Description:
'	Assigns the contents of the recordset to a new 2 dimensional array
'Requires:
'********************************************************************
Sub DB_RecSetToArray(arrData, recSet)
	dim intR
	dim intC
	redim arrData(recSet.RecordCount - 1, recSet.Fields.Count - 1)
	
	intR = 0
	do until recSet.EOF
		for intC = 0 to recSet.Fields.Count -1
			arrData(intR, intC) = recSet.Fields(intC).value
		next
		recSet.movenext
		intR = intR + 1
	loop
End Sub



'********************************************************************
'Function: 
'	DB_DeleteRecordByID
'Author:
'	Indra, June 03
'Description:
'	This function is used to remove a record from the database.
'Requires:
'	sqlerror.asp - to produce the error message
'	global dim of "strMessageText"
'********************************************************************
Function DB_DeleteRecordByID(intUniqueID, strIdDataType, strSpName)

	Dim connObj
	Dim cmdObj
	Dim paraObj
	Dim arrParamInput
	Dim arrParamType
	Dim intCounter
	
	'establish the parameters & types
	 arrParamInput = Split(intUniqueID,DB_PARAM_DELIM)
	 arrParamType = Split(strIdDataType,DB_PARAM_DELIM)

	'Open the connection
	Set connObj = Server.CreateObject("ADODB.Connection")
	With connObj
		.ConnectionString = Session("ConnectionString")
		.Open
	End With

	'Creating the command object
	Set cmdObj = Server.CreateObject("ADODB.Command")
	With cmdObj
		.ActiveConnection = connObj
		.CommandText = strSpName
		.CommandType = AdCmdStoredProc
	End With


	'loop through the input arrays and append parameters
	For intCounter = 0 To uBound(arrParamInput)

		'Creating parameters
		Select Case arrParamType(intCounter)

			Case "char"
				Set paraObj = cmdObj.CreateParameter(,AdChar,AdParamInput,6,arrParamInput(intCounter))
				cmdObj.Parameters.Append paraObj

			Case Else
				Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4,arrParamInput(intCounter))
				cmdObj.Parameters.Append paraObj

		End Select
	Next

	'debug_dbase cmdobj

	'Execute Command
	On Error Resume Next
	cmdObj.Execute
	On Error Goto 0
	
	'Check for errors
	if CheckForSQLError(connObj,"Delete",strMessageText) = TRUE then
		DB_DeleteRecordByID = FALSE
	else
		DB_DeleteRecordByID = TRUE
	end if

	'Destroy all the created objects
	Call DB_closeObject(paraObj)
	Call DB_closeObject(cmdObj)
	Call DB_closeObject(connObj)

End Function


'***************************************************************
'Function Name: DB_NullToEmpty
'Arguments: 
'	varElement as variant
'Description: This function tests to see if value is either null or an empty string
'Output: Returns the unchanged argument if the value is not null or Empty otherwise
'Special Instructions: Tests for null and empty strings and converts them to Empty
'***************************************************************
Function DB_NullToEmpty(varElement)
	if (trim(varElement) = "") then
		DB_NullToEmpty = Empty
	elseif (isNull(varElement) = True) then
		DB_NullToEmpty = Empty
	elseif (isEmpty(varElement) = True) then
		DB_NullToEmpty = Empty
	else
		DB_NullToEmpty = varElement
	end If
End Function



'***************************************************************
'Procedure      :DB_ParamLength
'Arguments      :strElement
'Purpose        :to return a positive value for the length of 
'				:a string, whilst ensuring no errors
'Output         :
'Dependancies   :NA
'Dependants     :NA
'Special Info   :NA
'***************************************************************
Function DB_ParamLength(strElement)
	if trim(strElement = "") then
		DB_ParamLength = 1
	elseif isNull(strElement) then
		DB_ParamLength = 1
	elseif isEmpty(strElement) then
		DB_ParamLength = 1
	else
		DB_ParamLength = Len(strElement)
		If DB_ParamLength < 1 Then
			DB_ParamLength = 1
		End If
	end If
End Function

%>