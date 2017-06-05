<%
'********************************************************************
'Author: Kym Oberauer - Digital Zoo Pty Ltd
'Date: 16/12/1998 Version 1.0
'
'Library: sqlerror.asp
'Description: This library consists of functions used to handle errors
'which occur using the ADO objects and connections with ODBC databases.
'********************************************************************

	'****************************************************************
	'Function:CheckForSQLError
	'Description:
	'****************************************************************
	Function CheckForSQLError(connObj,strSQLType,strMessageText)
		Dim intNumErrors, errorObj
		
		CheckForSQLError = FALSE
		intNumErrors = connObj.Errors.Count
		UTL_DisplayMessage "Errors: " & intNumErrors,strMessageText,AdmErrorColour
		if intNumErrors > 0 then
			CheckForSQLError = TRUE
			UTL_DisplayMessage "The following errors have occured:<br>", strMessageText,AdmErrorColour
			for each errorObj in connObj.Errors
				if TestADOError(errorObj,strSQLType,strMessageText) = FALSE then
					if TestMSQLError(errorObj,strSQLType,strMessageText) = FALSE then
						TestCustomError errorObj,strSQLType,strMessageText 
					end if
				end if
			next
		end if
		
	End Function


	'****************************************************************
	'Function:TestADOError
	'Description:
	'****************************************************************
	Function TestADOError(errorObj,strSQLType,strMessageText)
		TestADOError = TRUE
		Select Case errorObj.Number
			Case adErrInvalidArgument
				UTL_DisplayMessage "Invalid Argument",strMessageText,AdmErrorColour
			Case adErrNoCurrentRecord
				UTL_DisplayMessage "No Current Record",strMessageText,AdmErrorColour
			Case adErrIllegalOperation
				UTL_DisplayMessage "Illegal Operation",strMessageText,AdmErrorColour
			Case adErrInTransaction
				UTL_DisplayMessage "In Transaction",strMessageText,AdmErrorColour
			Case adErrFeatureNotAvailable
				UTL_DisplayMessage "Feature Not Available",strMessageText,AdmErrorColour
			Case adErrItemNotFound
				UTL_DisplayMessage "Item Not Found",strMessageText,AdmErrorColour
			Case adErrObjectInCollection
				UTL_DisplayMessage "No Object In Collection",strMessageText,AdmErrorColour
			Case adErrObjectNotSet
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrDataConversion
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrObjectClosed
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrObjectOpen
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrProviderNotFound
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrBoundToCommand
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrInvalidParamInfo
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrInvalidConnection
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrStillExecuting
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case adErrStillConnecting
				UTL_DisplayMessage "Standard Error",strMessageText,AdmErrorColour
			Case else
				TestADOError = FALSE
		End Select
	End Function

	'****************************************************************
	'Function:TestMSQLError
	'Description:
	'****************************************************************
	Function TestMSQLError(errorObj,strSQLType,strMessageText)
		
		TestMSQLError = TRUE
		Select Case errorObj.NativeError
			Case 547 'primary/foreign key relationship violation
				UTL_DisplayMessage "Error Num: " & errorObj.NativeError & " Record may not be deleted because of relationship with other records.",strMessageText,AdmErrorColour
				UTL_DisplayMessage "Remove the relationship with other records before attempting to delete.",strMessageText,AdmErrorColour
			Case 2627 'tried to duplicate unqiue value - eg user name
				UTL_DisplayMessage "The name already exists. Please try another name.",strMessageText,AdmErrorColour
			Case else
				TestMSQLError = FALSE
		End Select
	End Function

	'****************************************************************
	'Function:TestCustomError
	'Description:
	'****************************************************************
	Function TestCustomError(errorObj,strSQLType,strMessageText)
		
		UTL_DisplayMessage "Error Num: " & errorObj.NativeError & " " & Replace(errorObj.Description, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""),strMessageText,AdmErrorColour
	
	End Function
	

%>
