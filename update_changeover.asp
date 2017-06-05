<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "changeover" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Changeover</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtCustomer,"Customer");
	reason += validateSpecialCharacters(theForm.txtCustomer,"Customer");	
	reason += validateEmptyField(theForm.txtPhone,"Phone");
	reason += validateSpecialCharacters(theForm.txtContactPerson,"Contact Person");
	reason += validateEmptyField(theForm.txtAddress,"Address");
	reason += validateEmptyField(theForm.txtCity,"City");
	reason += validateEmptyField(theForm.txtPostcode,"Postcode");
	reason += validateEmptyField(theForm.txtOldModel,"Old Model");
	reason += validateEmptyField(theForm.txtOldModelSerial,"Old Model Serial");
	reason += validateEmptyField(theForm.txtReplacementModel,"Replacement Model");

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';

		return true;
    }
}

function submitComment(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtComment,"Comment");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}
</script>
<%

Sub getChangeover
	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_changeover WHERE changeover_id = " & intID

	rs.Open strSQL, conn
	
    if not DB_RecSetIsEmpty(rs) Then
		session("customer") 		= rs("customer")
		session("contact_person") 	= rs("contact_person")
		session("phone") 			= rs("phone")
		session("mobile") 			= rs("mobile")
		session("address") 			= rs("address")
		session("city") 			= rs("city")
		session("state") 			= rs("state")
		session("postcode") 		= rs("postcode")
		session("old_model") 		= rs("old_model")
		session("old_model_serial") = rs("old_model_serial")
		session("proof") 			= rs("proof")
		session("warranty") 		= rs("warranty")
		session("destroy") 			= rs("destroy")
		session("replacement_model") = rs("replacement_model")
		session("make_up_cost") 	= rs("make_up_cost")
		session("replacement_destination") = rs("replacement_destination")
		session("date_received") 	= rs("date_received")
		session("date_payment") 	= rs("date_payment")
		session("status") 			= rs("status")
		session("invoice_no") 		= rs("invoice_no")
		session("connote") 			= rs("connote")
		session("date_created") 	= rs("date_created")
		session("created_by") 		= rs("created_by")
		session("date_modified") 	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
		session("comments") 		= rs("comments")
    end if

    call CloseDataBase()

end sub

sub updateChangeover
	dim strSQL
	dim intID
	intID = request("id")
	
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
	
	strCustomer			= Replace(Request.Form("txtCustomer"),"'","''")
	strContactPerson	= Replace(Request.Form("txtContactPerson"),"'","''")
	strPhone			= Replace(Request.Form("txtPhone"),"'","''")
	strMobile			= Replace(Request.Form("txtMobile"),"'","''")
	strAddress			= Replace(Request.Form("txtAddress"),"'","''")
	strCity				= Replace(Request.Form("txtCity"),"'","''")
	strState			= Replace(Request.Form("cboState"),"'","''")
	strPostcode			= Replace(Request.Form("txtPostcode"),"'","''")
	strOldModel			= Replace(Request.Form("txtOldModel"),"'","''")
	strOldModelSerial	= Replace(Request.Form("txtOldModelSerial"),"'","''")
	intProof			= Request.Form("chkProof")
	intWarranty			= Request.Form("chkWarranty")
	intDestroy			= Request.Form("chkDestroy")
	strReplacementModel = Replace(Request.Form("txtReplacementModel"),"'","''")
	strMakeUpCost		= Replace(Request.Form("txtMakeUpCost"),"'","''")
	strReplacementDestination = Replace(Request.Form("txtReplacementDestination"),"'","''")
	strDateReceived		= Replace(Request.Form("txtDateReceived"),"'","''")
	strDatePayment		= Replace(Request.Form("txtDatePayment"),"'","''")
	strInvoiceNo		= Replace(Request.Form("txtInvoiceNo"),"'","''")
	strConnote			= Replace(Request.Form("txtConnote"),"'","''")
	strComments			= Replace(Request.Form("txtComments"),"'","''")
	intStatus			= Request.Form("cboStatus")

	Call OpenDataBase()

	strSQL = "UPDATE yma_changeover SET "
	strSQL = strSQL & "customer = '" & strCustomer & "',"
	strSQL = strSQL & "contact_person = '" & strContactPerson & "',"
	strSQL = strSQL & "phone = '" & strPhone & "',"
	strSQL = strSQL & "mobile = '" & strMobile & "',"
	strSQL = strSQL & "address = '" & strAddress & "',"
	strSQL = strSQL & "city = '" & strCity & "',"
	strSQL = strSQL & "state = '" & strState & "',"
	strSQL = strSQL & "postcode = '" & strPostcode & "',"
	strSQL = strSQL & "old_model = '" & strOldModel & "',"
	strSQL = strSQL & "old_model_serial = '" & strOldModelSerial & "',"
	strSQL = strSQL & "proof = '" & intProof & "',"
	strSQL = strSQL & "warranty = '" & intWarranty & "',"
	strSQL = strSQL & "destroy = '" & intDestroy & "',"
	strSQL = strSQL & "replacement_model = '" & strReplacementModel & "',"
	strSQL = strSQL & "make_up_cost = '" & strMakeUpCost & "',"
	strSQL = strSQL & "replacement_destination = '" & strReplacementDestination & "',"
	strSQL = strSQL & "date_received = CONVERT(datetime,'" & strDateReceived & "',103),"
	strSQL = strSQL & "date_payment = CONVERT(datetime,'" & strDatePayment & "',103),"
	strSQL = strSQL & "invoice_no = '" & strInvoiceNo & "',"
	strSQL = strSQL & "connote = '" & strConnote & "',"
	strSQL = strSQL & "comments = '" & strComments & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & intStatus & "' WHERE changeover_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The changeover log has been updated."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	dim intID
	intID 	= request("id")
	
	call getChangeover
	call listComments(intID,changeoverModuleID)
		
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateChangeover
				call getChangeover
			case "Comment"
				call addComment(intID,changeoverModuleID)
				call listComments(intID,changeoverModuleID)
		end select
	end if
end sub

call main

dim strMessageText
dim strCommentsList
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
            <h2>Update Changeover Log</h2>
            <h3><u>Claim no: <%= request("id") %></u></h3>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("modified_by") %></td>
                <td><%= displayDateFormatted(session("date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_changeover" id="form_update_changeover" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" width="1024">
          <tr>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="3" class="item_maintenance_header">Customer Info</td>
                </tr>
                <tr>
                  <td colspan="3">Customer<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCustomer" name="txtCustomer" maxlength="30" size="40" value="<%= session("customer") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3">Contact person:<br />
                    <input type="text" id="txtContactPerson" name="txtContactPerson" maxlength="30" size="40" value="<%= session("contact_person") %>" /></td>
                </tr>
                <tr>
                  <td>Phone no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPhone" name="txtPhone" maxlength="12" size="15" value="<%= session("phone") %>" /></td>
                  <td colspan="2">Mobile phone:<br />
                    <input type="text" id="txtMobile" name="txtMobile" maxlength="12" size="15" value="<%= session("mobile") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3">Address<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtAddress" name="txtAddress" maxlength="50" size="60" value="<%= session("address") %>" /></td>
                </tr>
                <tr>
                  <td width="50%">City<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCity" name="txtCity" maxlength="30" size="35" value="<%= session("city") %>" /></td>
                  <td width="20%">State:<br />
                    <select name="cboState">
                      <option <% if session("state") = "VIC" then Response.Write " selected" end if%> value="VIC">VIC</option>
                      <option <% if session("state") = "NSW" then Response.Write " selected" end if%> value="NSW">NSW</option>
                      <option <% if session("state") = "ACT" then Response.Write " selected" end if%> value="ACT">ACT</option>
                      <option <% if session("state") = "QLD" then Response.Write " selected" end if%> value="QLD">QLD</option>
                      <option <% if session("state") = "NT" then Response.Write " selected" end if%> value="NT">NT</option>
                      <option <% if session("state") = "WA" then Response.Write " selected" end if%> value="WA">WA</option>
                      <option <% if session("state") = "SA" then Response.Write " selected" end if%> value="SA">SA</option>
                      <option <% if session("state") = "TAS" then Response.Write " selected" end if%> value="TAS">TAS</option>
                    </select></td>
                  <td width="30%">Postcode<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPostcode" name="txtPostcode" maxlength="4" size="8" value="<%= session("postcode") %>" /></td>
                </tr>
              </table>
              <br />
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header">Additional Info</td>
                </tr>
                <tr>
                  <td><textarea name="txtComments" id="txtComments" cols="50" rows="3"><%= session("comments") %></textarea></td>
                </tr>
                <tr class="status_row">
                  <td>Status:
                    <select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select></td>
                </tr>
              </table></td>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Changeover Info</td>
                </tr>
                <tr>
                  <td colspan="2">Old model<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtOldModel" name="txtOldModel" maxlength="30" size="35" value="<%= session("old_model") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Old model serial no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtOldModelSerial" name="txtOldModelSerial" maxlength="20" size="20" value="<%= session("old_model_serial") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="checkbox" name="chkProof" id="chkProof" value="1" <% if session("proof") = "1" then Response.Write " checked" end if%> />
                    Proof of purchase<br />
                    <input type="checkbox" name="chkWarranty" id="chkWarranty" value="1" <% if session("warranty") = "1" then Response.Write " checked" end if%> />
                    Warranty<br />
                    <input type="checkbox" name="chkDestroy" id="chkDestroy" value="1" <% if session("destroy") = "1" then Response.Write " checked" end if%> />
                    Destroy
                    </td>
                </tr>               
                <tr>
                  <td colspan="2">Replacement model<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtReplacementModel" name="txtReplacementModel" maxlength="20" size="20" value="<%= session("replacement_model") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Make up cost:<br />
                    $
                    <input type="text" id="txtMakeUpCost" name="txtMakeUpCost" maxlength="6" size="8" value="<%= session("make_up_cost") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Replacement going to:<br />
                    <input type="text" id="txtReplacementDestination" name="txtReplacementDestination" maxlength="30" size="40" value="<%= session("replacement_destination") %>" /></td>
                </tr>
                <tr>
                  <td width="50%">Date received:<br />
                    <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10" value="<%= session("date_received") %>" />
                    <em>DD/MM/YYYY</em></td>
                  <td width="50%">Date payment:<br />
                    <input type="text" id="txtDatePayment" name="txtDatePayment" maxlength="10" size="10" value="<%= session("date_payment") %>" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
                <tr>
                  <td>Invoice no:<br />
                    <input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" value="<%= session("invoice_no") %>" /></td>
                  <td>Con-note:<br />
                    <input type="text" id="txtConnote" name="txtConnote" maxlength="20" size="20" value="<%= session("connote") %>" /></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Update Changeover" />
        </p>
      </form>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">        
        <%= strCommentsList %>
        <tr>
          <td><form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
              <p><input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
              <input type="hidden" name="Action" />
              <input type="submit" value="Add Comment" /></p>
            </form></td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateReceived'),		
        firstDay: 1,
        minDate: new Date('2000-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2000,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
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