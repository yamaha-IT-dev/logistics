<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Roadshow Item</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script language="JavaScript" type="text/javascript">
function validateProductCode(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Item Code has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateItemGroup(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Item Group has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateDescription(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Description has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateQuantity(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "- Quantity has not been filled in.\n"
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
	
	reason += validateProductCode(theForm.txtProductCode);  
	reason += validateWords(theForm.txtProductCode);
	//reason += validateItemGroup(theForm.txtItemGroup);  
	//reason += validateWords(theForm.txtItemGroup);	
	reason += validateDescription(theForm.txtDescription);  
	reason += validateWords(theForm.txtDescription);
	reason += validateQuantity(theForm.txtQuantity);  
	reason += validateWords(theForm.txtQuantity);
  	//reason += validateWords(theForm.txtComment);
	
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
'Call OpenDataBase

sub addItem
	dim strSQL

	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_roadshow_aug2011 (product_code, item_group, description, category, department, rrp, sku_type, prototype, quantity, packaging, source, origin, available, transit, invoice_no, pallet_no, loading_sequence, pre_sold, comments, return_to, owner, date_created, created_by, status) VALUES ( "
	strSQL = strSQL & "'" & trim(Request.Form("txtProductCode")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("txtItemGroup")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("txtDescription")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboCategory")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboDepartment")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtRRP")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboSkuType")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboPrototype")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtQuantity")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboPackaging")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboSource")) & "',"		
	strSQL = strSQL & "'" & trim(Request.Form("cboOrigin")) & "',"	
	strSQL = strSQL & "'" & trim(Request.Form("cboAvailable")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboTransit")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtPalletNo")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("txtLoadingSequence")) & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboPreSold")) & "',"		
	strSQL = strSQL & "'" & Replace(Request.Form("txtComment"),"'","''") & "',"
	strSQL = strSQL & "'" & trim(Request.Form("cboReturnTo")) & "',"		
	'strSQL = strSQL & " CONVERT(datetime,'" & trim(Request.Form("txtBumpInDate")) & "',103),"
	'strSQL = strSQL & " CONVERT(datetime,'" & trim(Request.Form("txtBumpOutDate")) & "',103),"
	strSQL = strSQL & "'" & trim(Request.Form("cboOwner")) & "',"		
	strSQL = strSQL & "getdate(),"	
	strSQL = strSQL & "'" & session("UsrUserName") & "',1)"
	'strSQL = strSQL & "'" & trim(Request.Form("cboStatus")) & "')" 	
	
	response.Write strSQL	  
	on error resume next
	conn.Execute strSQL
	
	'On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		'Response.Write "Success!"			
		Response.Redirect("thank-you_item.asp")
	end if 
	
	Call CloseDataBase()

end sub

sub main
	call UTL_validateLogin
	'call UTL_validateRoadshowLogin  
	if Trim(Request("Action")) = "Add" then		
		call addItem
	end if
end sub

call main
%>
</head>
<body>
<form action="" method="post" name="form_add_shipment" id="form_add_shipment" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table border="0" width="800" cellpadding="3">
          <tr>
            <td colspan="2" align="left"><font color="green"><%= strMessageText %></font>
              <h2>Add NEW Roadshow Item</h2></td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="#f0f0f0" align="center"><strong>Item Details</strong></td>
          </tr>
          <tr>
            <td width="25%">Owner:</td>
            <td width="75%"><select name="cboOwner">
                <option value="cameront">Cameron Tait</option>
                <option value="felixe">Felix Elliot</option>
                <option value="jamesh">James Harvey</option>
                <option value="jamieg">Jamie Goff</option>
                <option value="nathanb">Nathan Biggin</option>
                <option value="shaunm">Shaun McMahon</option>
                <option value="stevenv">Steven Vranch</option>
              </select></td>
          </tr>
          <tr>
            <td>Item Code<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtProductCode" name="txtProductCode" maxlength="30" size="30" /></td>
          </tr>
          <tr>
            <td>Item Group:</td>
            <td><input type="text" id="txtItemGroup" name="txtItemGroup" maxlength="30" size="30" /></td>
          </tr>
          <tr>
            <td>Description<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDescription" name="txtDescription" maxlength="80" size="85" /></td>
          </tr>
          <tr>
            <td>Category:</td>
            <td><select name="cboCategory">
                <option value="Acoustic Drums">Acoustic Drums</option>
                <option value="Brass &amp; Woodwind">Brass &amp; Woodwind</option>
                <option value="CA">CA</option>
                <option value="Digital Pianos">Digital Pianos</option>
                <option value="Electronic Drums">Electronic Drums</option>
                <option value="Guitars">Guitars</option>
                <option value="MPP">MPP</option>
                <option value="Paiste">Paiste</option>
                <option value="Percussions &amp; Strings">Percussions &amp; Strings</option>
                <option value="Pianos">Pianos</option>
                <option value="POS">POS</option>
                <option value="Portable Keyboards">Portable Keyboards</option>
                <option value="Pro Audio">Pro Audio</option>
                <option value="SYDE">SYDE</option>
                <option value="VOX">VOX</option>
                <option value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td>Department:</td>
            <td><select name="cboDepartment">
                <option value="Pro">Pro</option>
                <option value="Trad">Trad</option>
              </select></td>
          </tr>
          <tr>
            <td>RRP:</td>
            <td>$
              <input type="text" id="txtRRP" name="txtRRP" maxlength="8" size="8" /></td>
          </tr>
          <tr>
            <td>SKU Type:</td>
            <td><select name="cboSkuType">
                <option value="A Sku">A Sku</option>
                <option value="B Sku">B Sku</option>
                <option value="New">New</option>
              </select></td>
          </tr>
          <tr>
            <td>Prototype:</td>
            <td><select name="cboPrototype">
                <option value="Yes">Yes</option>
                <option value="No">No</option>
              </select></td>
          </tr>
          <tr>
            <td>Quantity<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtQuantity" name="txtQuantity" maxlength="8" size="8" /></td>
          </tr>
          <tr>
            <td>Packaging:</td>
            <td><select name="cboPackaging">
                <option value="Carton">Carton</option>
                <option value="Roadcase">Roadcase</option>
                <option value="Unboxed">Unboxed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="#f0f0f0" align="center"><strong>Logistics Details</strong></td>
          </tr>
          <tr>
            <td>Source:</td>
            <td><select name="cboSource">
                <option value="Asset">Asset</option>
                <option value="Loan Stock">Loan Stock</option>
                <option value="New Pick">New Pick</option>
                <option value="Promo Stock">Promo Stock</option>
                <option value="3S">3S</option>
              </select></td>
          </tr>
          <tr>
            <td>Origin:</td>
            <td><select name="cboOrigin">
                <option value="Head Office">Head Office</option>
                <option value="Kagan">Kagan</option>
                <option value="Direct from YCJ">Direct from YCJ</option>
                <option value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td>Available?</td>
            <td><select name="cboAvailable">
                <option value="In Transit">In Transit</option>
                <option value="No">No</option>
                <option value="Yes">Yes</option>
              </select></td>
          </tr>
          <tr>
            <td>Transit:</td>
            <td><select name="cboTransit">
                <option value="Air Freight">Air Freight</option>
                <option value="Sea Freight">Sea Freight</option>
                <option value="NA">NA</option>
              </select></td>
          </tr>
          <tr>
            <td>Invoice No:</td>
            <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="30" size="30" /></td>
          </tr>
          <tr>
            <td>Pallet No:</td>
            <td><input type="text" id="txtPalletNo" name="txtPalletNo" maxlength="30" size="30" /></td>
          </tr>
          <tr>
            <td>Loading Sequence:</td>
            <td><input type="text" id="txtLoadingSequence" name="txtLoadingSequence" maxlength="30" size="30" /></td>
          </tr>
          <tr>
            <td>Type:</td>
            <td><select name="cboType">
                <option value="Existing Asset">Existing Asset</option>
                <option value="Existing Loan">Existing Loan</option>
                <option value="New Loan">New Loan</option>
                <option value="New Order">New Order</option>
              </select></td>
          </tr>
          <tr>
            <td>Available for sale?</td>
            <td><select name="cboPreSold">
                <option value="No">No</option>
                <option value="Yes">Yes</option>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComment" id="txtComment" cols="45" rows="5"></textarea></td>
          </tr>
          <tr>
            <td>After-show Return to:</td>
            <td><select name="cboReturnTo">
                <option value="3S">3S</option>
                <option value="Dealer">Dealer</option>
                <option value="Head Office">Head Office</option>
                <option value="Kagan">Kagan</option>
              </select></td>
          </tr>
          <tr>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Add" /></td>
          </tr>
        </table>
        <p><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_item.asp">Back To Roadshow Product List</a></p></td>
    </tr>
    
  </table>
</form>
</body>
</html>