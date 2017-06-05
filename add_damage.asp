<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsDamageType.asp" -->
<% strSection = "damage" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Warehouse Damage</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtItemName,"Item name");
	reason += validateEmptyField(theForm.txtSerialNo,"Serial no");	
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
		
		return true;
    }
}
</script>
<%
sub addDamagedItem
	dim strSQL
	
	dim strItemName
	dim strSerialNo
	dim strDamageType
	dim strDamageLocation
	dim strDamageConnote
	dim strCourseDamage
	dim intSentExcel
	dim strSentExcelDate
	dim strComments
	
	strItemName 		= trim(Request.Form("txtItemName"))
	strSerialNo 		= trim(Request.Form("txtSerialNo"))
	strDamageType 		= trim(Request.Form("cboDamageType"))
	strDamageLocation 	= trim(Request.Form("txtLocation"))
	strDamageConnote 	= trim(Request.Form("txtConnote"))
	strCourseDamage 	= trim(Request.Form("cboCourseDamage"))
	intSentExcel 		= trim(Request.Form("cboSentExcel"))
	strSentExcelDate 	= trim(Request.Form("txtSentExcelDate"))
	strComments 		= trim(Request.Form("txtComments"))
	
	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_damage ("
	strSQL = strSQL & "damage_item, "
	strSQL = strSQL & "damage_serial_no, "
	strSQL = strSQL & "damage_type, "
	strSQL = strSQL & "damage_location, "
	strSQL = strSQL & "damage_connote, "
	strSQL = strSQL & "course_damage, "
	strSQL = strSQL & "sent_excel, "
	strSQL = strSQL & "sent_excel_date, "
	strSQL = strSQL & "date_created, "
	strSQL = strSQL & "created_by, "
	strSQL = strSQL & "damage_comments"
	strSQL = strSQL & ") VALUES ( "
	strSQL = strSQL & "'" & Server.HTMLEncode(strItemName) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & "'" & strDamageType & "',"
	strSQL = strSQL & "'" & strDamageLocation & "',"
	strSQL = strSQL & "'" & strDamageConnote & "',"
	strSQL = strSQL & "'" & strCourseDamage & "',"
	strSQL = strSQL & "'" & intSentExcel & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strSentExcelDate & "',103),getdate(),"
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "')"
	
	'response.Write strSQL	
	  
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else		
		Response.Redirect("list_damage.asp")
	end if 
	
	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin  
	call getDamageTypeList
	call getCourseDamageList
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then		
			call addDamagedItem
		end if
	end if
end sub

call main

dim strDamageTypeList
dim strCourseDamageList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_damage.asp"><img src="images/icon_damaged.jpg" border="0" alt="Damage Items" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_damage.asp">Back to List</a>
            <h2>Add Warehouse Damage</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_damaged_item" id="form_add_damaged_item" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">Warehouse Damage</td>
          </tr>
          <tr>
            <td width="50%">Item name<span class="mandatory">*</span>:<br />
              <input type="text" id="txtItemName" name="txtItemName" maxlength="20" size="30" /></td>
            <td width="50%">Serial no<span class="mandatory">*</span>:<br />
              <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td>Damage type:<br />
              <select name="cboDamageType">
                <%= strDamageTypeList %>
              </select></td>
            <td>Cause:<br />
              <select name="cboCourseDamage">
                <%= strCourseDamageList %>
              </select></td>
          </tr>
          <tr>
            <td>Location: <img src="images/icon_new.gif" border="0" align="top" /><br />
              <input type="text" id="txtLocation" name="txtLocation" maxlength="20" size="30" /></td>
            <td>Con-note: <img src="images/icon_new.gif" border="0" align="top" /><br />
              <input type="text" id="txtConnote" name="txtConnote" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td colspan="2">Sent to Excel:
              <select name="cboSentExcel">
                <option value="0" rel="none">No</option>
                <option value="1" rel="date">Yes</option>
              </select></td>
          </tr>
          <tr rel="date">
            <td colspan="2">Date:
              <input type="text" id="txtSentExcelDate" name="txtSentExcelDate" maxlength="10" size="10" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td colspan="2">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="55" rows="3"></textarea></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
          <input type="submit" value="Add Warehouse Damage" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>