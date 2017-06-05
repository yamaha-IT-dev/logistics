<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "recall" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Customer Recall</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script language="JavaScript" type="text/javascript">
function validateDealer(fld) {
    var error = "";

    if (fld.value.length == 0) {
        fld.style.background = 'Yellow';
        error = "- Dealer has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateCustomer(fld) {
    var error = "";

    if (fld.value.length == 0) {
        fld.style.background = 'Yellow';
        error = "- Customer has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateAddress(fld) {
    var error = "";

    if (fld.value.length == 0) {
        fld.style.background = 'Yellow';
        error = "- Address has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function trim(s)
{
  return s.replace(/^\s+|\s+$/, '');
}

function validateEmail(fld) {
    var error="";
    var tfld = trim(fld.value);                        // value of field with whitespace trimmed off
    var emailFilter = /^[^@]+@[^@.]+\.[^@]*\w\w$/ ;
    var illegalChars= /[\(\)\<\>\,\;\:\\\"\[\]]/ ;

    if (fld.value == "") {
        fld.style.background = 'Yellow';
        error = "- Email address has not been filled in.\n";
    } else if (!emailFilter.test(tfld)) {              //test email for illegal characters
        fld.style.background = 'Yellow';
        error = "- Please enter a valid email address.\n";
    } else if (fld.value.match(illegalChars)) {
        fld.style.background = 'Yellow';
        error = "- Email address contains illegal characters.\n";
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateWords(fld) {
	var error = "";

	    var iChars = "@#%^&*+=[];'";
        for (var i = 0; i < fld.value.length; i++) {
                if (iChars.indexOf(fld.value.charAt(i)) != -1) {
					fld.style.background = 'Yellow';
                	error = "- Special characters are not allowed. Please remove them. \n";
        		}
        }
	return error;
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateDealer(theForm.txtDealer);

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

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
<%

Sub getStockMod

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_customer_recall WHERE recall_id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("dealer") 			= rs("dealer")
		session("product") 			= rs("product")
		session("qty") 				= rs("qty")
		session("customer_name") 	= rs("customer_name")
		session("customer_address") = rs("customer_address")
		session("customer_city") 	= rs("customer_city")
		session("customer_state") 	= rs("customer_state")
		session("customer_postcode") = rs("customer_postcode")
		session("customer_phone") 	= rs("customer_phone")
		session("customer_mobile") 	= rs("customer_mobile")
		session("customer_email") 	= rs("customer_email")
		session("tested_by") 		= rs("tested_by")
		session("site_visit") 		= rs("site_visit")
		session("status") 			= rs("status")
		session("date_created") 	= rs("date_created")
		session("created_by") 		= rs("created_by")
		session("date_modified") 	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
		session("comments") 		= rs("comments")
    end if

	rs.close
	set rs = nothing
    call CloseDataBase()

end sub

sub updateStockMod

	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_customer_recall SET "
	strSQL = strSQL & "dealer = '" & trim(Request.Form("txtDealer")) & "',"
	strSQL = strSQL & "product = '" & trim(Request.Form("cboProduct")) & "',"
	strSQL = strSQL & "qty = '" & trim(Request.Form("txtQty")) & "',"
	strSQL = strSQL & "customer_name = '" & trim(Request.Form("txtCustomer")) & "',"
	strSQL = strSQL & "customer_address = '" & trim(Request.Form("txtAddress")) & "',"
	strSQL = strSQL & "customer_city = '" & trim(Request.Form("txtCity")) & "',"
	strSQL = strSQL & "customer_state = '" & trim(Request.Form("cboState")) & "',"
	strSQL = strSQL & "customer_postcode = '" & trim(Request.Form("txtPostcode")) & "',"
	strSQL = strSQL & "customer_email = '" & trim(Request.Form("txtEmail")) & "',"
	strSQL = strSQL & "customer_phone = '" & trim(Request.Form("txtPhone")) & "',"
	strSQL = strSQL & "customer_mobile = '" & trim(Request.Form("txtMobile")) & "',"
	strSQL = strSQL & "tested_by = '" & trim(Request.Form("txtTestedBy")) & "',"
	strSQL = strSQL & "site_visit = '" & trim(Request.Form("cboSiteVisit")) & "',"
	strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComments"),"'","''") & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & trim(Request.Form("cboStatus")) & "' WHERE recall_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The customer recall has been updated."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin

	'response.Write strRef
	if Trim(Request("Action")) = "Update" then
		call updateStockMod
		call getStockMod
	else
		call getStockMod
	end if

end sub

call main

dim strMessageText
%>
</head>
<body>
<form action="" method="post" name="form_update_recall" id="form_update_recall" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table border="0" width="800">
          <tr>
            <td colspan="2" align="left"><font color="green"><%= strMessageText %></font><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_recall.asp">Back To Customer Recall List</a>
              <h2>Update Customer Recall</h2></td>
          </tr>
          <tr>
            <td>ID:</td>
            <td><%= request("id") %></td>
          </tr>
          <tr>
            <td width="30%">Dealer<span class="mandatory">*</span>:</td>
            <td width="70%"><input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="40" value="<%= session("dealer") %>" /></td>
          </tr>
          <tr>
            <td>Product:</td>
            <td><select name="cboProduct">
                <option <% if session("product") = "P2500S" then Response.Write " selected" end if%> value="P2500S">P2500S</option>
                <option <% if session("product") = "P3500S" then Response.Write " selected" end if%> value="P3500S">P3500S</option>
                <option <% if session("product") = "P5000SJ" then Response.Write " selected" end if%> value="P5000SJ">P5000SJ</option>
                <option <% if session("product") = "P7000S" then Response.Write " selected" end if%> value="P7000S">P7000S</option>
                <option <% if session("product") = "DSR112" then Response.Write " selected" end if%> value="DSR112">DSR112</option>
                <option <% if session("product") = "DSR115" then Response.Write " selected" end if%> value="DSR115">DSR115</option>
                <option <% if session("product") = "DSR118W" then Response.Write " selected" end if%> value="DSR118W">DSR118W</option>
                <option <% if session("product") = "MS101III" then Response.Write " selected" end if%> value="MS101III">MS101III</option>
                <option <% if session("product") = "MS50DR" then Response.Write " selected" end if%> value="MS50DR">MS50DR</option>
                <option <% if session("product") = "MS100DR" then Response.Write " selected" end if%> value="MS100DR">MS100DR</option>
                <option <% if session("product") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td>Qty:</td>
            <td><input type="text" id="txtQty" name="txtQty" maxlength="5" size="5" value="<%= session("qty") %>" /></td>
          </tr>
          <tr>
            <td colspan="2"><hr /><h3>Customer Details</h3></td>
          </tr>
          <tr>
            <td>Customer:</td>
            <td><input type="text" id="txtCustomer" name="txtCustomer" maxlength="50" size="50" value="<%= session("customer_name") %>" /></td>
          </tr>
          <tr>
            <td>Address:</td>
            <td><input type="text" id="txtAddress" name="txtAddress" maxlength="50" size="60" value="<%= session("customer_address") %>" /></td>
          </tr>
          <tr>
            <td>City:</td>
            <td><input type="text" id="txtCity" name="txtCity" maxlength="15" size="20" value="<%= session("customer_city") %>" /></td>
          </tr>
          <tr>
            <td>State:</td>
            <td><select name="cboState">
                <option <% if session("customer_state") = "VIC" then Response.Write " selected" end if%> value="VIC">VIC</option>
                <option <% if session("customer_state") = "NSW" then Response.Write " selected" end if%> value="NSW">NSW</option>
                <option <% if session("customer_state") = "ACT" then Response.Write " selected" end if%> value="ACT">ACT</option>
                <option <% if session("customer_state") = "QLD" then Response.Write " selected" end if%> value="QLD">QLD</option>
                <option <% if session("customer_state") = "NT" then Response.Write " selected" end if%> value="NT">NT</option>
                <option <% if session("customer_state") = "WA" then Response.Write " selected" end if%> value="WA">WA</option>
                <option <% if session("customer_state") = "SA" then Response.Write " selected" end if%> value="SA">SA</option>
                <option <% if session("customer_state") = "TAS" then Response.Write " selected" end if%> value="TAS">TAS</option>
              </select></td>
          </tr>
          <tr>
            <td>Postcode:</td>
            <td><input type="text" id="txtPostcode" name="txtPostcode" maxlength="4" size="8" value="<%= session("customer_postcode") %>" /></td>
          </tr>
          <tr>
            <td>Email:</td>
            <td><input type="text" id="txtEmail" name="txtEmail" maxlength="40" size="50" value="<%= session("customer_email") %>" /></td>
          </tr>
          <tr>
            <td>Phone:</td>
            <td><input type="text" id="txtPhone" name="txtPhone" maxlength="15" size="25" value="<%= session("customer_phone") %>" /></td>
          </tr>
          <tr>
            <td>Mobile:</td>
            <td><input type="text" id="txtMobile" name="txtMobile" maxlength="15" size="25" value="<%= session("customer_mobile") %>" /></td>
          </tr>
          <tr>
            <td colspan="2"><hr /></td>
          </tr>
          <tr>
            <td>Tested By:</td>
            <td><input type="text" id="txtTestedBy" name="txtTestedBy" maxlength="20" size="20" value="<%= session("tested_by") %>" /></td>
          </tr>
          <tr>
            <td>Site Visit:</td>
            <td><select name="cboSiteVisit">
                <option <% if session("site_visit") = "1" then Response.Write " selected" end if%> value="1">Yes</option>
                <option <% if session("site_visit") = "0" then Response.Write " selected" end if%> value="0">No</option>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="60" rows="5"><%= session("comments") %></textarea></td>
          </tr>
          <tr>
            <td>Status:</td>
            <td><select name="cboStatus">
                <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
              </select></td>
          </tr>
          <tr>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Update" />
              <p><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_recall.asp">Back To Customer Recall List</a></p></td>
            <td><table width="300" cellpadding="4" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td width="50%">Date Created:</td>
                  <td width="50%"><%= session("date_created") %></td>
                </tr>
                <tr>
                  <td>Created By:</td>
                  <td><%= session("created_by") %></td>
                </tr>
                <tr>
                  <td>Last Modified Date:</td>
                  <td><%= session("date_modified") %></td>
                </tr>
                <tr>
                  <td>Last Modified By:</td>
                  <td><%= session("modified_by") %></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    
  </table>
</form>
</body>
</html>