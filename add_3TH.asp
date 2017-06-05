<!--#include file="include/connection_it.asp " -->
<!--#include file="class/cls3thReturn.asp " -->
<!--#include file="class/clsWarehouseReturn.asp " -->
<% strSection = "3TH" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add 3TH</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" language="JavaScript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.cboReturnType,"Type");
	
	reason += validateEmptyField(theForm.cboDepartment,"Department");
	
	reason += validateEmptyField(theForm.txtItemCode,"Item Code");
	reason += validateSpecialCharacters(theForm.txtItemCode,"Item Code");
	
	reason += validateEmptyField(theForm.txtDescription,"Item Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Item Description");
	
	reason += validateSpecialCharacters(theForm.txtDealer,"Dealer");
	
	reason += validateEmptyField(theForm.txtShipmentNo,"Shipment No");	
	reason += validateSpecialCharacters(theForm.txtShipmentNo,"Shipment No");
	
	reason += validateNumeric(theForm.txtQty,"Qty");
	
	reason += validateSpecialCharacters(theForm.txtLabelNo,"Label no");
	
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
	session("reason_code") = ""
	
	call UTL_validateLogin
	call getReasonCode
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim intReturnType
		dim strDepartment
		dim strItemCode
		dim strShipmentNo
		dim intQty
		dim strDescription
		dim strGRA
		dim strCarrier
		dim strLabelNo
		dim strOriginalConnote
		dim strDealer
		dim intInstruction
		dim strSerialNo
		dim intStockType
		dim strDateReceived
		dim strComments
		dim intStatus
			
		intReturnType 		= Request.Form("cboReturnType")	
		strDepartment 		= Request.Form("cboDepartment")	
		strItemCode 		= Replace(Trim(Request.Form("txtItemCode")),"'","''")
		strShipmentNo 		= Replace(Trim(Request.Form("txtShipmentNo")),"'","''")
		intQty 				= Replace(Trim(Request.Form("txtQty")),"'","''")
		strDescription 		= Replace(Trim(Request.Form("txtDescription")),"'","''")		
		strGRA 				= Replace(Trim(Request.Form("txtGRA")),"'","''")
		strCarrier 			= Request.Form("cboCarrier")
		strLabelNo 			= Replace(Trim(Request.Form("txtLabelNo")),"'","''")
		strOriginalConnote 	= Replace(Trim(Request.Form("txtOriginalConnote")),"'","''")
		strDealer 			= Replace(Trim(Request.Form("txtDealer")),"'","''")	
		intInstruction 		= Request.Form("cboInstruction")
		strSerialNo 		= Replace(Trim(Request.Form("txtSerialNo")),"'","''")
		intStockType		= Request.Form("cboStockType")	
		strDateReceived		= Request.Form("txtDateReceived")
		strComments 		= Replace(Trim(Request.Form("txtComments")),"'","''")
	
		if Trim(Request("Action")) = "Add" then
			call add3thReturn(intReturnType, strDepartment, strItemCode, strShipmentNo, intQty, strDescription, strGRA, strCarrier, strLabelNo, strOriginalConnote, strDealer, intInstruction, strSerialNo, intStockType, strDateReceived, strComments, session("UsrUserName"))
		end if
	end if
end sub

call main

dim strReasonCodeList
dim strMessageText
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_quarantine.asp"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_3TH.asp">Back to List</a>
            <h2>Add 3TH</h2>
            <%= strMessageText %></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_return" id="form_add_return" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3">Type<span class="mandatory">*</span>:
              <select name="cboReturnType" onchange="disableList(this);">
                <option value="" rel="none">...</option>
                <option value="1" rel="none">Lost in Warehouse</option>
                <option value="2" rel="none">Lost by Carrier</option>
                <option value="3" rel="none">Packaging Issue</option>
                <option value="4" rel="none">Warehouse Variance</option>
                <option value="5" rel="none">Display Stock</option>
              </select></td>
          </tr>
          <tr>
            <td>Department<span class="mandatory">*</span>:<br />
              <select name="cboDepartment">
                <option value="">...</option>
                <option value="AV">AV</option>
                <option value="MPD">MPD</option>
                <option value="Other">Other</option>
              </select></td>
            <td colspan="2">Item code<span class="mandatory">*</span>:<br />
              <input type="text" id="txtItemCode" name="txtItemCode" maxlength="20" size="25" /></td>
          </tr>
          <tr>
            <td colspan="3">Item description<span class="mandatory">*</span>:<br />
              <input type="text" id="txtDescription" name="txtDescription" maxlength="50" size="60" /></td>
          </tr>
          <tr>
            <td>Label #:<br />
              <input type="text" id="txtLabelNo" name="txtLabelNo" maxlength="15" size="20" /></td>
            <td colspan="2">Con-note:<br />
              <input type="text" id="txtOriginalConnote" name="txtOriginalConnote" maxlength="15" size="20" /></td>
          </tr>
          <tr>
            <td width="60%">Dealer:<br />
              <input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="35" /></td>
            <td width="25%">Shipment #<span class="mandatory">*</span>:<br />
              <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="10" size="12" /></td>
            <td width="15%">Qty<span class="mandatory">*</span>:<br />
              <input type="text" id="txtQty" name="txtQty" maxlength="4" size="5" /></td>
          </tr>
          <tr>
            <td>Carrier:<br />
              <select name="cboCarrier">
                <option value="">...</option>
                <option value="Cope">Cope</option>
                <option value="StarTrack">StarTrack</option>
                <option value="Schenker">Schenker</option>
                <option value="Kings">Kings</option>
              </select></td>
            <td colspan="2">Date received:<br />
              <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="15"  /></td>
          </tr>
          <tr>
            <td colspan="3">Serial #:<br />
              <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="50" size="55" /></td>
          </tr>
          <tr>
            <td>Instruction:<br />
              <select name="cboInstruction">
                <option value="">...</option>
                <option value="1">Update GRA</option>
                <option value="2">Writeoff Approval Required</option>
              </select></td>
            <td colspan="2">GRA:<br />
              <input type="text" id="txtGRA" name="txtGRA" maxlength="7" size="10" /></td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="3"></textarea></td>
          </tr>
          <tr>
            <td colspan="3"><input type="hidden" name="Action" />
              <input type="submit" value="Add 3TH" />
              <input type="reset" value="Reset" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateReceived'),
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2013,2020],
		format: 'DD/MM/YYYY'
    });		
</script>
</body>
</html>