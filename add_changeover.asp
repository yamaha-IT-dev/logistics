<!--#include file="include/connection_it.asp " -->
<% strSection = "changeover" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Changeover</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtCustomer,"Customer");
    reason += validateSpecialCharacters(theForm.txtCustomer,"Customer");
    reason += validateEmptyField(theForm.txtPhone,"Phone");
    reason += validateSpecialCharacters(theForm.txtPhone,"Phone");
    reason += validateSpecialCharacters(theForm.txtContactPerson,"Contact person");
    reason += validateEmptyField(theForm.txtAddress,"Address");
    reason += validateEmptyField(theForm.txtCity,"City");
    reason += validateNumeric(theForm.txtPostcode,"Postcode");
    reason += validateEmptyField(theForm.txtOldModel,"Old model");
    reason += validateEmptyField(theForm.txtOldModelSerial,"Old model serial no");
    reason += validateEmptyField(theForm.txtReplacementModel,"Replacement model");
    reason += validateSpecialCharacters(theForm.txtComments,"Comments");

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

sub addChangeover
    dim strSQL

    dim strCustomer
    dim strContactPerson
    dim strPhone
    dim strMobile
    dim strAddress
    dim strCity
    dim strState
    dim strPostcode
    dim strOldModel
    dim strOldModelSerial
    dim intProof
    dim intWarranty
    dim intDestroy
    dim strReplacementModel
    dim strMakeUpCost
    dim strReplacementDestination
    dim strDateReceived
    dim strDatePayment
    dim strInvoiceNo
    dim strConnote
    dim strComments
    dim intStatus

    strCustomer         = Replace(Request.Form("txtCustomer"),"'","''")
    strContactPerson    = Replace(Request.Form("txtContactPerson"),"'","''")
    strPhone            = Replace(Request.Form("txtPhone"),"'","''")
    strMobile           = Replace(Request.Form("txtMobile"),"'","''")
    strAddress          = Replace(Request.Form("txtAddress"),"'","''")
    strCity             = Replace(Request.Form("txtCity"),"'","''")
    strState            = Replace(Request.Form("cboState"),"'","''")
    strPostcode         = Replace(Request.Form("txtPostcode"),"'","''")
    strOldModel         = Replace(Request.Form("txtOldModel"),"'","''")
    strOldModelSerial   = Replace(Request.Form("txtOldModelSerial"),"'","''")
    intProof            = Request.Form("chkProof")
    intWarranty         = Request.Form("chkWarranty")
    intDestroy          = Request.Form("chkDestroy")
    strReplacementModel = Replace(Request.Form("txtReplacementModel"),"'","''")
    strMakeUpCost       = Replace(Request.Form("txtMakeUpCost"),"'","''")
    strReplacementDestination = Replace(Request.Form("txtReplacementDestination"),"'","''")
    strDateReceived     = Replace(Request.Form("txtDateReceived"),"'","''")
    strDatePayment      = Replace(Request.Form("txtDatePayment"),"'","''")
    strInvoiceNo        = Replace(Request.Form("txtInvoiceNo"),"'","''")
    strConnote          = Replace(Request.Form("txtConnote"),"'","''")
    strComments         = Replace(Request.Form("txtComments"),"'","''")
    intStatus           = Request.Form("cboStatus")

    call OpenDataBase()

    strSQL = "INSERT INTO yma_changeover (customer, contact_person, phone, mobile, address, city, state, postcode, old_model, old_model_serial, proof, warranty, destroy, replacement_model, make_up_cost, replacement_destination, date_received, date_payment, invoice_no, connote, date_created, created_by, comments) VALUES ( "
    strSQL = strSQL & "'" & strCustomer & "',"
    strSQL = strSQL & "'" & strContactPerson & "',"
    strSQL = strSQL & "'" & strPhone & "',"
    strSQL = strSQL & "'" & strMobile & "',"
    strSQL = strSQL & "'" & strAddress & "',"
    strSQL = strSQL & "'" & strCity & "',"
    strSQL = strSQL & "'" & strState & "',"
    strSQL = strSQL & "'" & strPostcode & "',"
    strSQL = strSQL & "'" & strOldModel & "',"
    strSQL = strSQL & "'" & strOldModelSerial & "',"
    strSQL = strSQL & "'" & intProof & "',"
    strSQL = strSQL & "'" & intWarranty & "',"
    strSQL = strSQL & "'" & intDestroy & "',"
    strSQL = strSQL & "'" & strReplacementModel & "',"
    strSQL = strSQL & "'" & strMakeUpCost & "',"
    strSQL = strSQL & "'" & strReplacementDestination & "',"
    strSQL = strSQL & "CONVERT(datetime,'" & strDateReceived & "',103),"
    strSQL = strSQL & "CONVERT(datetime,'" & strDatePayment & "',103),"
    strSQL = strSQL & "'" & strInvoiceNo & "',"	
    strSQL = strSQL & "'" & strConnote & "',getdate(),"	
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "'" & strComments & "')"

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    'On error Goto 0  

    if err <> 0 then
        strMessageText = err.description
    else
        Response.Redirect("thank-you_changeover.asp")
    end if

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if Trim(Request("Action")) = "Add" then
            call addChangeover
        end if
    end if
end sub

call main
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_changeover.asp"><img src="images/icon_changeover.jpg" border="0" alt="Changeover" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_changeover.asp">Back to List</a>
            <h2>Add Changeover Log</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_changeover" id="form_add_changeover" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" width="1024">
          <tr>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="3" class="item_maintenance_header">Customer Info</td>
                </tr>
                <tr>
                  <td colspan="3">Customer<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCustomer" name="txtCustomer" maxlength="30" size="40" /></td>
                </tr>
                <tr>
                  <td colspan="3">Contact person:<br />
                    <input type="text" id="txtContactPerson" name="txtContactPerson" maxlength="30" size="40" /></td>
                </tr>
                <tr>
                  <td>Phone no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPhone" name="txtPhone" maxlength="12" size="15" /></td>
                  <td colspan="2">Mobile phone:<br />
                    <input type="text" id="txtMobile" name="txtMobile" maxlength="12" size="15" /></td>
                </tr>
                <tr>
                  <td colspan="3">Address<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtAddress" name="txtAddress" maxlength="50" size="60" /></td>
                </tr>
                <tr>
                  <td width="50%">City<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCity" name="txtCity" maxlength="30" size="35" /></td>
                  <td width="20%">State:<br />
                    <select name="cboState">
                      <option value="VIC">VIC</option>
                      <option value="NSW">NSW</option>
                      <option value="ACT">ACT</option>
                      <option value="QLD">QLD</option>
                      <option value="NT">NT</option>
                      <option value="WA">WA</option>
                      <option value="SA">SA</option>
                      <option value="TAS">TAS</option>
                    </select></td>
                  <td width="30%">Postcode<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPostcode" name="txtPostcode" maxlength="4" size="5" /></td>
                </tr>
              </table>
              <br />
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header">Additional Info</td>
                </tr>
                <tr>
                  <td>Comments:<br />
                    <textarea name="txtComments" id="txtComments" cols="50" rows="4"></textarea></td>
                </tr>
              </table></td>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Changeover Info</td>
                </tr>
                <tr>
                  <td colspan="2">Old model<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtOldModel" name="txtOldModel" maxlength="30" size="35" /></td>
                </tr>
                <tr>
                  <td colspan="2">Old model serial no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtOldModelSerial" name="txtOldModelSerial" maxlength="20" size="20" /></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="checkbox" name="chkProof" id="chkProof" value="1" />
                    Proof of purchase<br />
                    <input type="checkbox" name="chkWarranty" id="chkWarranty" value="1" />
                    Warranty<br />
                    <input type="checkbox" name="chkDestroy" id="chkDestroy" value="1" />
                    Destroy
                    </td>
                </tr>
                <tr>
                  <td colspan="2">Replacement model<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtReplacementModel" name="txtReplacementModel" maxlength="20" size="20" /></td>
                </tr>
                <tr>
                  <td colspan="2">Make up cost:<br />
                    $
                    <input type="text" id="txtMakeUpCost" name="txtMakeUpCost" maxlength="6" size="8" /></td>
                </tr>
                <tr>
                  <td colspan="2">Replacement going to:<br />
                    <input type="text" id="txtReplacementDestination" name="txtReplacementDestination" maxlength="30" size="40" /></td>
                </tr>
                <tr>
                  <td width="50%">Date received:<br />
                    <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em></td>
                  <td width="50%">Date payment:<br />
                    <input type="text" id="txtDatePayment" name="txtDatePayment" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
                <tr>
                  <td>Invoice no:<br />
                    <input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" /></td>
                  <td>Connote:<br />
                    <input type="text" id="txtConnote" name="txtConnote" maxlength="20" size="20" /></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <p><input type="hidden" name="Action" />
              <input type="submit" value="Add Changeover" /></p>
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">
    var picker = new Pikaday({
        field: document.getElementById('txtDateReceived'),
        firstDay: 1,
        minDate: new Date('2000-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2000,2020],
        format: 'DD/MM/YYYY'
    });

    var picker = new Pikaday({
        field: document.getElementById('txtDatePayment'),
        firstDay: 1,
        minDate: new Date('2000-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2000,2020],
        format: 'DD/MM/YYYY'
    });
</script>
</body>
</html>