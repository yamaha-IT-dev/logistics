<!--#include file="include/connection_it.asp " -->
<% strSection = "recall" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Customer Recall</title>
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
        theForm.Action.value = 'Add';
  		//theForm.submit();
		
		return true;
    }
}
</script>
<%

sub addRecall
	dim strSQL

	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_customer_recall (dealer, product, qty, customer_name, customer_address, customer_city, customer_state, customer_postcode, customer_email, customer_phone, customer_mobile, tested_by, site_visit, date_created, created_by, comments, status) VALUES ( "
	strSQL = strSQL & "'" & trim(Request.Form("txtDealer")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboProduct")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtQty")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtCustomer")) & "',"		
	strSQL = strSQL & "'" & trim(Request.Form("txtAddress")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtCity")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboState")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtPostcode")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtEmail")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtPhone")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtMobile")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtTestedBy")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboSiteVisit")) & "',getdate(),"	
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & Replace(Request.Form("txtComments"),"'","''") & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboStatus")) & "')"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	'On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		'Response.Write "Success!"
		Response.Redirect("thank-you_recall.asp")
	end if 
	
	'conn.close
	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin  
	if Trim(Request("Action")) = "Add" then		
		call addRecall
	end if
end sub

call main
%>
</head>
<body>
<form action="" method="post" name="form_add_recall" id="form_add_recall" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table border="0" width="800">
          <tr>
            <td colspan="2" align="left"><font color="green"><%= strMessageText %></font>
              <h2>Add NEW Customer Recall</h2></td>
          </tr>
          <tr>
            <td width="30%">Dealer<span class="mandatory">*</span>:</td>
            <td width="70%"><input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="40" /></td>
          </tr>
          <tr>
            <td>Product:</td>
            <td><select name="cboProduct">
            	<option value="P2500S">P2500S</option>
                <option value="P3500S">P3500S</option>
                <option value="P5000SJ">P5000SJ</option>
                <option value="P7000S">P7000S</option>
                <option value="DSR112">DSR112</option>
                <option value="DSR115">DSR115</option>
                <option value="DSR118W">DSR118W</option>                
                <option value="MS101III">MS101III</option>
                <option value="MS50DR">MS50DR</option>
				<option value="MS100DR">MS100DR</option>
                <option value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td>Qty:</td>
            <td><input type="text" id="txtQty" name="txtQty" maxlength="5" size="5" /></td>
          </tr>
          <tr>
            <td colspan="2"><hr /><h3>Customer Details</h3></td>
          </tr>
          <tr>
            <td>Customer:</td>
            <td><input type="text" id="txtCustomer" name="txtCustomer" maxlength="50" size="50" /></td>
          </tr>
          <tr>
            <td>Address:</td>
            <td><input type="text" id="txtAddress" name="txtAddress" maxlength="50" size="60" /></td>
          </tr>
          <tr>
            <td>City:</td>
            <td><input type="text" id="txtCity" name="txtCity" maxlength="15" size="20" /></td>
          </tr>
          <tr>
            <td>State:</td>
            <td><select name="cboState">
                <option value="VIC">VIC</option>
                <option value="NSW">NSW</option>
                <option value="ACT">ACT</option>
                <option value="QLD">QLD</option>
                <option value="NT">NT</option>
                <option value="WA">WA</option>
                <option value="SA">SA</option>
                <option value="TAS">TAS</option>
              </select></td>
          </tr>
          <tr>
            <td>Postcode:</td>
            <td><input type="text" id="txtPostcode" name="txtPostcode" maxlength="4" size="8" /></td>
          </tr>
          <tr>
            <td>Email:</td>
            <td><input type="text" id="txtEmail" name="txtEmail" maxlength="40" size="50" /></td>
          </tr>
          <tr>
            <td>Phone:</td>
            <td><input type="text" id="txtPhone" name="txtPhone" maxlength="15" size="25" /></td>
          </tr>
          <tr>
            <td>Mobile:</td>
            <td><input type="text" id="txtMobile" name="txtMobile" maxlength="15" size="25" /></td>
          </tr>
          
          <tr>
            <td colspan="2"><hr /></td>
          </tr>
          <tr>
            <td>Tested by:</td>
            <td><input type="text" id="txtTestedBy" name="txtTestedBy" maxlength="20" size="20" /></td>
          </tr>
          <tr>
            <td>Site Visit:</td>
            <td><select name="cboSiteVisit">
                <option value="1">Yes</option>
                <option value="0">No</option>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="60" rows="5"></textarea></td>
          </tr>
          <tr>
            <td>Status:</td>
            <td><select name="cboStatus">
                <option value="1">Open</option>
                <option value="0">Completed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
              <input type="submit" value="Add" />
              <p><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_recall.asp">Back To Customer Recall List</a></p></td>
          </tr>
        </table></td>
    </tr>
    
  </table>
</form>
</body>
</html>