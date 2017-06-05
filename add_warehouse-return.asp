<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsWarehouseReturn.asp " -->
<% strSection = "quarantine" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Warehouse Return</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" language="JavaScript">
function disableList(thisForm) {
    if (thisForm.cboReturnType.value == 1) {
        alert("Managed return");
    }
}

function validateFormOnSubmit(theForm) {
    var reason      = "";
    var blnSubmit   = true;

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
    reason += validateEmptyField(theForm.txtReturnConnote,"Return Connote");
    reason += validateSpecialCharacters(theForm.txtReturnConnote,"Return Connote");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Add';

        return true;
    }
}
</script>
<%
sub addReturn
    dim strSQL

    dim intReturnType
    dim strDepartment
    dim strItemCode
    dim strShipmentNo
    dim intQty
    dim strDescription
    dim intPhotos
    dim strGRA
    dim strReturnCarrier
    dim strReturnConnote
    dim strOriginalConnote
    dim strDealer
    dim intReasonCode
    dim intInstruction
    dim strSerialNo
    dim intStockType
    dim strDateReceived
    dim strComments
    dim intStatus


    intReturnType       = Request.Form("cboReturnType")
    strDepartment       = Request.Form("cboDepartment")
    strItemCode         = Request.Form("txtItemCode")
    strShipmentNo       = Request.Form("txtShipmentNo")
    intQty              = Request.Form("txtQty")
    strDescription      = Replace(Replace(Request.Form("txtDescription"), "'", "''"), " ", " ")
    intPhotos           = Request.Form("chkPhotos")
    strGRA              = Request.Form("txtGRA")
    strReturnCarrier    = Request.Form("cboReturnCarrier")
    strReturnConnote    = Request.Form("txtReturnConnote")
    strOriginalConnote  = Request.Form("txtOriginalConnote")
    strDealer           = Replace(Replace(Request.Form("txtDealer"), "'", "''"), " ", " ")
    intReasonCode       = Request.Form("cboReasonCode")
    intInstruction      = Request.Form("cboInstruction")
    strSerialNo         = Request.Form("txtSerialNo")
    intStockType        = Request.Form("cboStockType")
    strDateReceived     = Request.Form("txtDateReceived")
    strComments         = Replace(Replace(Request.Form("txtComments"), "'", "''"), " ", " ")

    call OpenDataBase()

    strSQL = "INSERT INTO yma_quarantines ("
    strSQL = strSQL & " return_type, "
    strSQL = strSQL & " department, "
    strSQL = strSQL & " item_code, "
    strSQL = strSQL & " shipment_no, "
    strSQL = strSQL & " qty, "
    strSQL = strSQL & " description, "
    strSQL = strSQL & " photos, "
    strSQL = strSQL & " gra, "
    strSQL = strSQL & " return_carrier, "
    strSQL = strSQL & " return_connote, "
    strSQL = strSQL & " original_connote, "
    strSQL = strSQL & " dealer, "
    strSQL = strSQL & " reason_code, "
    strSQL = strSQL & " instruction, "
    strSQL = strSQL & " serial_no, "
    strSQL = strSQL & " stock_type, "
    strSQL = strSQL & " date_received, "
    strSQL = strSQL & " comments, "
    strSQL = strSQL & " created_by"
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & "'" & intReturnType & "',"
    strSQL = strSQL & "'" & strDepartment & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strItemCode) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strShipmentNo) & "',"
    strSQL = strSQL & "'" & intQty & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strDescription) & "',"
    strSQL = strSQL & "'" & intPhotos & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strGRA) & "',"
    strSQL = strSQL & "'" & strReturnCarrier & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strReturnConnote) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strOriginalConnote) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strDealer) & "',"
    strSQL = strSQL & "'" & intReasonCode & "',"
    strSQL = strSQL & "'" & intInstruction & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strSerialNo) & "',"
    strSQL = strSQL & "'" & intStockType & "',"
    strSQL = strSQL & " CONVERT(datetime,'" & strDateReceived & "',103),"
    strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "')"

    response.Write strSQL
    on error resume next
    conn.Execute strSQL
Response.Write "Test here "
    if err <> 0 then
        strMessageText = err.description
		Response.Write "Tes here 1"
    else
	Response.Write "Test here 2"
        Response.Redirect("thank-you_quarantine.asp")
    end if

    call CloseDataBase()
end sub

sub main
    session("reason_code") = ""
    call UTL_validateLogin
    call getReasonCode

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if Trim(Request("Action")) = "Add" then
            call addReturn
        end if
    end if
end sub

call main

dim strReasonCodeList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_quarantine.asp"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_warehouse-return.asp">Back to List</a>
            <h2>Add Warehouse Return</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_return" id="form_add_return" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" class="item_maintenance_header">Warehouse Return</td>
          </tr>
          <tr>
            <td colspan="3">Return type<span class="mandatory">*</span>:
              <select name="cboReturnType" onchange="disableList(this);">
                <option value="" rel="none">...</option>
                <option value="1" rel="none">Managed (GRA)</option>
                <option value="0" rel="unmanaged">Un-managed</option>
                <option value="2" rel="none">Un-addressed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="3">&nbsp;</td>
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
            <td>Return con-note<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReturnConnote" name="txtReturnConnote" maxlength="15" size="20" /></td>
            <td colspan="2">Original con-note:<br />
              <input type="text" id="txtOriginalConnote" name="txtOriginalConnote" maxlength="15" size="20" /></td>
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
            <td colspan="3"><input type="checkbox" name="chkPhotos" id="chkPhotos" value="1" />
              <img src="images/camera_icon.gif" alt="Photo" border="0" /> Photo</td>
          </tr>
          <tr rel="unmanaged">
            <td colspan="3" bgcolor="#66CCFF"><table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr>
                  <td>Return carrier:<br />
                    <select name="cboReturnCarrier">
                      <option value="">...</option>
                      <option value="Cope">Cope</option>
                      <option value="StarTrack">StarTrack</option>
                      <option value="Schenker">Schenker</option>
                      <option value="Kings">Kings</option>
                    </select></td>
                  <td>Serial no(s): (If damaged)<br />
                    <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="50" size="55" /></td>
                </tr>
                <tr>
                  <td colspan="2" align="right">Stock: <img src="images/icon_new.gif" border="0" align="top" />
                    <select name="cboStockType">
                      <option value="">...</option>
                      <option value="1">Damaged</option>
                      <option value="2">Partial</option>
                    </select></td>
                </tr>
                <tr>
                  <td colspan="2">Date stock received:<br />
                    <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="15" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td>Reason: <img src="images/icon_new.gif" border="0" /><br />
              <select name="cboReasonCode">
                <%= strReasonCodeList %>
              </select></td>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td>Instruction:<br />
              <select name="cboInstruction">
                <option value="">...</option>
                <option value="1">Return to good stock 3T</option>
                <option value="2">Transfer to Excel 3XL</option>
                <option value="3">Resend to customer</option>
                <option value="4">Damaged item to Excel - good stock to 3T</option>
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
              <input type="submit" value="Add Warehouse Return" />
              <input type="reset" value="Reset" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>