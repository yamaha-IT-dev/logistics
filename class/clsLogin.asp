<%
'********************************************************************
'Function: 		SetSessionVariables
'Description: 	This function either collects the form variables and sets the required session variables or clears the previous values.
'				The session variables are used to maintain state between requests to add records. 
'Parameters:	blnSetType - to fill with form variables or clear
'********************************************************************
Function SetSessionVariables(blnSetType)
    if blnSetType = TRUE then
        'save the form values
        Session("UsrUserName")      = Server.HTMLEncode(Lcase(Request("txtUsername")))
        Session("UsrPassword")      = Server.HTMLEncode(Request("txtPassword"))
    elseif blnSetType = FALSE then
        'clear the session variables and set defaults
        Session("UsrUserName")      = ""
        Session("UsrPassword")      = ""
        Session("UsrLoginRole")     = 0
        Session("UsrUserID")        = 0
        Session("UsrDivision")      = ""
    else
        'error
        SetSessionVariables = FALSE
        Exit Function
    end if

    Session.Timeout = 180 'number of minutes

    SetSessionVariables = TRUE
End Function

function testUserLogin
    Dim cmdObj, paraObj

    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spUserLogin"
    cmdObj.CommandType = AdCmdStoredProc

    Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,50,Session("UsrUserName"))
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,50,Session("UsrPassword"))
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("intUserID",AdInteger,adParamInputOutput,4,0)
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("intLogin",AdInteger,adParamInputOutput,4,0)
    cmdObj.Parameters.Append paraObj

    On Error Resume Next
    cmdObj.Execute
    On error Goto 0

    if CheckForSQLError(conn,"Update",strMessageText) = TRUE then
        testUserLogin = FALSE
    else
        Session("UsrLoginRole")  = cmdObj("intLogin")
        Session("UsrUserID")     = cmdObj("intUserID")
        'Session("UsrDivision")   = cmdObj("strDivision")
        UTL_validateLogin
        AntiFixationInit()
        'if Request.Cookies("current_URL_cookie_logistics") = "" then
        Select Case Session("UsrLoginRole")
            Case 4
                Response.Redirect("list_item-maintenance.asp")
            Case 5
                Response.Redirect("list_stockmod.asp")
            Case 6
                Response.Redirect("list_changeover.asp")
            Case 7
                Response.Redirect("list_gra.asp")
            Case 9
                Response.Redirect("list_item-maintenance.asp")
            Case 10
                Response.Redirect("list_freights.asp")
            Case 11
                Response.Redirect("list_item-maintenance.asp")
            Case 12
                Response.Redirect("list_changeover.asp")
            Case 13
                Response.Redirect("list_transfer.asp")
            Case Else
                Response.Redirect("list_shipment.asp")
        End Select
        'else 
        '    Response.Redirect(Request.Cookies("current_URL_cookie_logistics"))
        'end if
        testUserLogin = TRUE
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function
%>