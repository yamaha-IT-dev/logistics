<!--#include file="include/connection_it.asp " -->
<!--#include file="include/FRM_build_form.asp " -->
<% strSection = "" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Dealer</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function validatePassword(fld) {
    var error = "";

    if (fld.value.length == 0) {
        fld.style.background = 'Yellow';
        error = "Please enter your new password.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function trim(s)
{
  return s.replace(/^\s+|\s+$/, '');
}

function validateFormOnSubmit(theForm) {	var reason = "";
	var blnSubmit = true;

  	reason += validatePassword(theForm.txtPassword);

  	if (reason != "") {
    	alert("" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
		//alert("Yea");
        theForm.Action.value = 'Update';
  		theForm.submit();

		return true;
    }
}
</script>
</head>
<body>
<%
function saveUserPassword
    Dim cmdObj, paraObj

    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spUpdateUserPassword"
    cmdObj.CommandType = AdCmdStoredProc

	'session("password") 		= request("txtPassword")

	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4, session("UsrUserId"))
	cmdObj.Parameters.Append paraObj

	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,255, request("txtPassword"))
	cmdObj.Parameters.Append paraObj

	On Error Resume Next
	cmdObj.Execute
    On error Goto 0

    if CheckForSQLError(conn,"Update",strMessageText) = TRUE then
	    saveDealer = FALSE
    else
        strMessageText = "Your password has been successfully updated."
		saveDealer = TRUE
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

sub main

    ' We check if we are still login
    call UTL_validateLogin

    ' We check the action of the form
    if Trim(Request("Action")) = "Update" then
        call saveUserPassword
    'else
        'intDealerId = request("dealer_id")
        'session("dealer_id") = request("dealer_id")

        'if len(trim(intDealerId)) > 0 then
         '   call getDealerDetails
        'end if
    end if

	'call getUser
	'call getState
end sub

dim strMessageText
call main

%>
<table cellspacing="0" cellpadding="0" align="center" width="100%">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><h2>Change your password</h2>
      <p><font color="green"><%= strMessageText %></font></p>
      <form action="" method="post" name="form_update_password" id="form_update_password" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" width="500">
          <tr>
            <td width="30%">Your username:</td>
            <td width="70%"><%= session("UsrUserName") %></td>
          </tr>
          <tr>
            <td>New password<span class="mandatory">*</span>:</td>
            <td><input type="password" id="txtPassword" name="txtPassword" maxlength="25" size="35" /></td>
          </tr>
        </table>
        <br />
        <input type="hidden" name="Action" />
        <input type="submit" value="Change Password" />
      </form></td>
  </tr>
</table>
</body>
</html>
