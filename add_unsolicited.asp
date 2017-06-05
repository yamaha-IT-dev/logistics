<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsUnsolicited.asp " -->
<% strSection = "unsolicited" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Incomplete Goods to Excel</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" language="JavaScript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtItemCode,"Item Code");
	reason += validateSpecialCharacters(theForm.txtItemCode,"Item Code");
	
	reason += validateEmptyField(theForm.txtDescription,"Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Description");
	
	reason += validateSpecialCharacters(theForm.txtDealer,"Dealer");
	
	reason += validateEmptyField(theForm.txtShipmentNo,"Shipment No");
	reason += validateSpecialCharacters(theForm.txtShipmentNo,"Shipment No");
	
	reason += validateNumeric(theForm.txtQty,"Qty");
	
	reason += validateEmptyField(theForm.txtConnote,"Return Connote");
	reason += validateSpecialCharacters(theForm.txtConnote,"Return Connote");
	
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
sub main
	call UTL_validateLogin	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then
			unsDepartment 	= Request.Form("cboDepartment")
			unsItemCode 	= Request.Form("txtItemCode")
			unsDescription 	= Request.Form("txtDescription")
			unsConnote 		= Request.Form("txtConnote")
			unsGRA 			= Request.Form("txtGRA")
			unsDealer 		= Request.Form("txtDealer")
			unsShipmentNo 	= Request.Form("txtShipmentNo")
			unsQty 			= Request.Form("txtQty")
			unsInstruction 	= Request.Form("cboInstruction")
			unsComments 	= Request.Form("txtComments")
				
			call addUnsolicited(unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, session("UsrUserName"))
		end if
	end if
end sub

call main

Dim unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, unsStatus, unsCreatedBy
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_quarantine.asp"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_unsolicited.asp">Back to List</a>
            <h2>Add Incomplete Goods to Excel</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_return" id="form_add_return" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" class="item_maintenance_header">Unsolicited Goods to Excel</td>
          </tr>
          <tr>
            <td>Department:<br />
              <select name="cboDepartment">
                <option value="AV">AV</option>
                <option value="MPD">MPD</option>
                <option value="Other">Other</option>
            </select></td>
            <td colspan="2">Item code<span class="mandatory">*</span>:<br />
              <input type="text" id="txtItemCode" name="txtItemCode" maxlength="20" size="25" /></td>
          </tr>
          <tr>
            <td colspan="3">Description<span class="mandatory">*</span>:<br />
              <input type="text" id="txtDescription" name="txtDescription" maxlength="50" size="60" /></td>
          </tr>
          <tr>
            <td>Con-note<span class="mandatory">*</span>:<br />
              <input type="text" id="txtConnote" name="txtConnote" maxlength="15" size="20" /></td>
            <td colspan="2">GRA:<br />
            <input type="text" id="txtGRA" name="txtGRA" maxlength="7" size="10" /></td>
          </tr>
          <tr>
            <td width="60%">Dealer:<br />
              <input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="35" /></td>
            <td width="20%">Shipment no<span class="mandatory">*</span>:<br />
              <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="10" size="12" /></td>
            <td width="20%">Qty<span class="mandatory">*</span>:<br />
              <input type="text" id="txtQty" name="txtQty" maxlength="4" size="5" /></td>
          </tr>
          <tr>
            <td>Instruction:<br />
              <select name="cboInstruction">
                <option value="1">Move to 3XL</option>
                <option value="2">Move to 3S</option>
                <option value="3">Investigate</option>
              </select></td>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="3"></textarea></td>
          </tr>
          <tr>
            <td colspan="3"><input type="hidden" name="Action" />
              <input type="submit" value="Add Unsolicited Goods" />
              <input type="reset" value="Reset" /></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
</body>
</html>