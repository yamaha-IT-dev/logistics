<!--#include file="include/connection_it.asp " -->
<% strSection = "return" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Return Log</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script language="JavaScript" type="text/javascript">
function validateModelName(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Model Name has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}
function validatePartNoBase(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Part No Base has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateWords(fld) {
	var error = "";
	
	    var iChars = "@#$%^&*+=[]\\\;/{}|\<>'";
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
	
	reason += validateModelName(theForm.txtModelName); 
	reason += validateWords(theForm.txtModelName);
	reason += validatePartNoBase(theForm.txtPartNoBase);    	
	reason += validateWords(theForm.txtPartNoBase);
	reason += validateWords(theForm.txtVendorModelNo);
	
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

sub addStockMod
	dim strSQL

	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_stock_mod (model_name, model_type, part_no_base, vendor_model_no, hardwired, document, date_created, created_by, comments, status) VALUES ( "
	strSQL = strSQL & "'" & trim(Request.Form("txtModelName")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboModelType")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("txtPartNoBase")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("txtVendorModelNo")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboHardwired")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboDocument")) & "',getdate(),"	
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtComments")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboStatus")) & "')" 	
	
	'response.Write strSQL	  
	on error resume next
	conn.Execute strSQL
	
	'On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		'Response.Write "Success!"			
		Response.Redirect("thank-you_return.asp")
	end if 
	
	'conn.close
	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then		
			call addStockMod
		end if
	end if
end sub

call main
%>
</head>
<body>
<form action="" method="post" name="form_add_stockmod" id="form_add_stockmod" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table border="0" width="800">
          <tr>
            <td colspan="2" align="left"><font color="green"><%= strMessageText %></font>
              <h2>Add NEW Return Log</h2>
              <p><strong>Progress: 50%</strong></p></td>
          </tr>
          <tr>
            <td>Department:</td>
            <td><select name="cboDepartment">
                <option value="AV">AV</option>
                <option value="MPD">MPD</option>
                <option value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td width="30%">Date Received<span class="mandatory">*</span>:</td>
            <td width="70%"><input type="text" id="txtModelName" name="txtModelName" maxlength="20" size="20" />
              <em>(DD/MM/YYYY)</em></td>
          </tr>
          <tr>
            <td>Return Con Note<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtPartNoBase" name="txtPartNoBase" maxlength="20" size="20" /></td>
          </tr>
          <tr>
            <td>Carrier:</td>
            <td><select name="cboModelType">
                <option value="-">...</option>
                <option value="LEAD">LEAD</option>
                <option value="PLUG">PLUG</option>
                <option value="ADAPTOR">ADAPTOR</option>
                <option value="ADAPTOR and DVD">ADAPTOR and DVD</option>
                <option value="ADAPTOR and LEAD">ADAPTOR and LEAD</option>
              </select></td>
          </tr>
          <tr>
            <td>Ship No:</td>
            <td><input type="text" id="txtVendorModelNo" name="txtVendorModelNo" maxlength="20" size="20" /></td>
          </tr>
          <tr>
            <td>Dealer Code:</td>
            <td><input type="text" id="txtVendorModelNo2" name="txtVendorModelNo2" maxlength="20" size="20" /></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="45" rows="5"></textarea></td>
          </tr>          
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
              <input type="submit" value="Add" />
              <p><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_stockmod.asp">Back To Return Logs Listing</a></p></td>
          </tr>
        </table></td>
    </tr>
    
  </table>
</form>
</body>
</html>