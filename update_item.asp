<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Roadshow Item</title>
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

	    var iChars = "@#$%^&*+=[]\\\;{}|\<>'";
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
	//reason += validateWords(theForm.txtProductCode);
	//reason += validateItemGroup(theForm.txtItemGroup);
	//reason += validateWords(theForm.txtItemGroup);
	reason += validateDescription(theForm.txtDescription);
	//reason += validateWords(theForm.txtDescription);
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
        theForm.Action.value = 'Update';
  		//theForm.submit();

		return true;
    }
}

</script>
<%

Sub getItem

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_item WHERE item_id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("product_code") = rs("product_code")
		session("item_group") = rs("item_group")
		session("description") = rs("description")
		session("category") = rs("category")
		session("department") = rs("department")
		session("rrp") = rs("rrp")
		session("sku_type") = rs("sku_type")
		session("prototype") = rs("prototype")
		session("quantity") = rs("quantity")
		session("packaging") = rs("packaging")
		session("source") = rs("source")
		session("origin") = rs("origin")
		session("available") = rs("available")
		session("transit") = rs("transit")
		session("invoice_no") = rs("invoice_no")
		session("pallet_no") = rs("pallet_no")
		session("loading_sequence") = rs("loading_sequence")
		session("type") = rs("type")
		session("pre_sold") = rs("pre_sold")
		session("comments") = rs("comments")
		'session("bump_in_date") = rs("bump_in_date")
		'session("bump_out_date") = rs("bump_out_date")
		session("return_to") = rs("return_to")
		session("status") = rs("status")
		session("date_modified") = rs("date_modified")
		session("modified_by") = rs("modified_by")
		session("item_date_created") = rs("date_created")
		session("item_created_by") = rs("created_by")
		session("owner") = rs("owner")
    end if

	rs.close
	set rs = nothing
    call CloseDataBase()

end sub

sub updateItem
	dim strSQL
	dim intID
	intID = request("id")

	OpenDataBase()

	strSQL = "UPDATE yma_item SET "
	strSQL = strSQL & "product_code = '" & trim(Request.Form("txtProductCode")) & "',"
	strSQL = strSQL & "item_group = '" & trim(Request.Form("txtItemGroup")) & "',"
	strSQL = strSQL & "description = '" & Replace(Request.Form("txtDescription"),"'","''") & "',"
	strSQL = strSQL & "category = '" & trim(Request.Form("cboCategory")) & "',"
	strSQL = strSQL & "department = '" & trim(Request.Form("cboDepartment")) & "',"
	strSQL = strSQL & "rrp = '" & trim(Request.Form("txtRRP")) & "',"
	strSQL = strSQL & "sku_type = '" & trim(Request.Form("cboSkuType")) & "',"
	strSQL = strSQL & "prototype = '" & trim(Request.Form("cboPrototype")) & "',"
	strSQL = strSQL & "quantity = '" & trim(Request.Form("txtQuantity")) & "',"
	strSQL = strSQL & "packaging = '" & trim(Request.Form("cboPackaging")) & "',"
	strSQL = strSQL & "source = '" & trim(Request.Form("cboSource")) & "',"
	strSQL = strSQL & "origin = '" & trim(Request.Form("cboOrigin")) & "',"
	strSQL = strSQL & "available = '" & trim(Request.Form("cboAvailable")) & "',"
	strSQL = strSQL & "transit = '" & trim(Request.Form("cboInTransit")) & "',"
	strSQL = strSQL & "invoice_no = '" & trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "pallet_no = '" & trim(Request.Form("txtPalletNo")) & "',"
	strSQL = strSQL & "loading_sequence = '" & trim(Request.Form("txtLoadingSequence")) & "',"
	strSQL = strSQL & "type = '" & trim(Request.Form("cboType")) & "',"
	strSQL = strSQL & "pre_sold = '" & trim(Request.Form("cboPreSold")) & "',"
	strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComment"),"'","''") & "',"
	'strSQL = strSQL & "bump_in_date = CONVERT(datetime,'" &	trim(Request.Form("txtBumpInDate")) & "',103),"
	'strSQL = strSQL & "bump_out_date = CONVERT(datetime,'" &	trim(Request.Form("txtBumpOutDate")) & "',103),"
	strSQL = strSQL & "return_to = '" & trim(Request.Form("cboReturnTo")) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & trim(Request.Form("cboStatus")) & "' WHERE item_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	conn.close

end sub

sub main
	call UTL_validateLogin

	if Trim(Request("Action")) = "Update" then
		call updateItem
		call getItem
	else
		call getItem
	end if

end sub

dim strMessageText
call main
%>
</head>
<body>
<form action="" method="post" name="form_update_shipment" id="form_update_shipment" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table border="0" width="800">
          <tr>
            <td colspan="2" align="left"><font color="green"><%= strMessageText %></font>
              <h2>Update Roadshow Item</h2></td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="#f0f0f0" align="center"><strong>Item Details</strong></td>
          </tr>
          <tr>
            <td width="25%">Owner:</td>
            <td width="75%"><%= session("owner") %></td>
          </tr>
          <tr>
            <td>Item Code<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtProductCode" name="txtProductCode" maxlength="20" size="20" value="<%= session("product_code") %>" /></td>
          </tr>
          <tr>
            <td>Item Group:</td>
            <td><input type="text" id="txtItemGroup" name="txtItemGroup" maxlength="20" size="20" value="<%= session("item_group") %>" /></td>
          </tr>
          <tr>
            <td>Description<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDescription" name="txtDescription" maxlength="80" size="85" value="<%= session("description") %>" /></td>
          </tr>
          <tr>
            <td>Category:</td>
            <td><select name="cboCategory">
                <option <% if session("category") = "Acoustic Drums" then Response.Write " selected" end if%> value="Acoustic Drums">Acoustic Drums</option>
                <option <% if session("category") = "Brass & Woodwind" then Response.Write " selected" end if%> value="Brass & Woodwind">Brass & Woodwind</option>
                <option <% if session("category") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
                <option <% if session("category") = "Digital Pianos" then Response.Write " selected" end if%> value="Digital Pianos">Digital Pianos</option>
                <option <% if session("category") = "Electronic Drums" then Response.Write " selected" end if%> value="Electronic Drums">Electronic Drums</option>
                <option <% if session("category") = "Guitars" then Response.Write " selected" end if%> value="Guitars">Guitars</option>
                <option <% if session("category") = "MPP" then Response.Write " selected" end if%> value="MPP">MPP</option>
                <option <% if session("category") = "Paiste" then Response.Write " selected" end if%> value="Paiste">Paiste</option>
                <option <% if session("category") = "Percussions & Strings" then Response.Write " selected" end if%> value="Percussions & Strings">Percussions & Strings</option>
                <option <% if session("category") = "Pianos" then Response.Write " selected" end if%> value="Pianos">Pianos</option>
                <option <% if session("category") = "POS" then Response.Write " selected" end if%> value="POS">POS</option>
                <option <% if session("category") = "Portable Keyboards" then Response.Write " selected" end if%> value="Portable Keyboards">Portable Keyboards</option>
                <option <% if session("category") = "Pro Audio" then Response.Write " selected" end if%> value="Pro Audio">Pro Audio</option>
                <option <% if session("category") = "SYDE" then Response.Write " selected" end if%> value="SYDE">SYDE</option>
                <option <% if session("category") = "VOX" then Response.Write " selected" end if%> value="VOX">VOX</option>
                <option <% if session("category") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
              </select></td>
          </tr>
          <tr>
            <td>Department:</td>
            <td><select name="cboDepartment">
                <option <% if session("department") = "Pro" then Response.Write " selected" end if%> value="Pro">Pro</option>
                <option <% if session("department") = "Trad" then Response.Write " selected" end if%> value="Trad">Trad</option>
              </select></td>
          </tr>
          <tr>
            <td>RRP:</td>
            <td>$
              <input type="text" id="txtRRP" name="txtRRP" maxlength="8" size="8" value="<%= session("rrp") %>" /></td>
          </tr>
          <tr>
            <td>SKU Type:</td>
            <td><select name="cboSkuType">
                <option <% if session("sku_type") = "" then Response.Write " selected" end if%> value="">...</option>
                <option <% if session("sku_type") = "A Sku" then Response.Write " selected" end if%> value="A Sku">A Sku</option>
                <option <% if session("sku_type") = "B Sku" then Response.Write " selected" end if%> value="B Sku">B Sku</option>
                <option <% if session("sku_type") = "New" then Response.Write " selected" end if%> value="New">New</option>
              </select></td>
          </tr>
          <tr>
            <td>Prototype:</td>
            <td><select name="cboPrototype">
                <option <% if session("prototype") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                <option <% if session("prototype") = "No" then Response.Write " selected" end if%> value="No">No</option>
              </select></td>
          </tr>
          <tr>
            <td>Quantity<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtQuantity" name="txtQuantity" maxlength="8" size="8" value="<%= session("quantity") %>" /></td>
          </tr>
          <tr>
            <td>Packaging:</td>
            <td><select name="cboPackaging">
                <option <% if session("packaging") = "Carton" then Response.Write " selected" end if%> value="Carton">Carton</option>
                <option <% if session("packaging") = "Roadcase" then Response.Write " selected" end if%> value="Roadcase">Roadcase</option>
                <option <% if session("packaging") = "Unboxed" then Response.Write " selected" end if%> value="Unboxed">Unboxed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="#f0f0f0" align="center"><strong>Logistics Details</strong></td>
          </tr>
          <tr>
            <td>Source:</td>
            <td><select name="cboSource">
                <option <% if session("source") = "" then Response.Write " selected" end if%> value="">...</option>
                <option <% if session("source") = "Asset" then Response.Write " selected" end if%> value="Asset">Asset</option>
                <option <% if session("source") = "Loan Stock" then Response.Write " selected" end if%> value="Loan Stock">Loan Stock</option>
                <option <% if session("source") = "New Pick" then Response.Write " selected" end if%> value="New Pick">New Pick</option>
                <option <% if session("source") = "Promo Stock" then Response.Write " selected" end if%> value="Promo Stock">Promo Stock</option>
                <option <% if session("source") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
              </select></td>
          </tr>
          <tr>
            <td>Origin:</td>
            <td><select name="cboOrigin">
                <option <% if session("origin") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
                <option <% if session("origin") = "Kagan" then Response.Write " selected" end if%> value="Kagan">Kagan</option>
                <option <% if session("origin") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                <option <% if session("origin") = "3K" then Response.Write " selected" end if%> value="3K">3K</option>
                <option <% if session("origin") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                <option <% if session("origin") = "Direct from YCJ" then Response.Write " selected" end if%> value="Direct from YCJ">Direct from YCJ</option>
              </select></td>
          </tr>
          <tr>
            <td>Available?</td>
            <td><select name="cboAvailable">
                <option <% if session("available") = "In Transit" then Response.Write " selected" end if%> value="In Transit">In Transit</option>
                <option <% if session("available") = "No" then Response.Write " selected" end if%> value="No">No</option>
                <option <% if session("available") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
              </select></td>
          </tr>
          <tr>
            <td>Transit:</td>
            <td><select name="cboInTransit">
                <option <% if session("transit") = "" then Response.Write " selected" end if%> value="">...</option>
                <option <% if session("transit") = "Air Freight" then Response.Write " selected" end if%> value="Air Freight">Air Freight</option>
                <option <% if session("transit") = "Sea Freight" then Response.Write " selected" end if%> value="Sea Freight">Sea Freight</option>
                <option <% if session("transit") = "NA" then Response.Write " selected" end if%> value="NA">NA</option>
              </select></td>
          </tr>
          <tr>
            <td bgcolor="#FFFF00">Invoice No:</td>
            <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" value="<%= session("invoice_no") %>" /></td>
          </tr>
          <tr>
            <td bgcolor="#FFFF00">Pallet No:</td>
            <td><input type="text" id="txtPalletNo" name="txtPalletNo" maxlength="20" size="20" value="<%= session("pallet_no") %>" /></td>
          </tr>
          <tr>
            <td bgcolor="#FFFF00">Loading Sequence:</td>
            <td><input type="text" id="txtLoadingSequence" name="txtLoadingSequence" maxlength="20" size="20" value="<%= session("loading_sequence") %>" /></td>
          </tr>
          <tr>
            <td>Type:</td>
            <td><select name="cboType">
                <option <% if session("type") = "Existing Asset" then Response.Write " selected" end if%> value="Existing Asset">Existing Asset</option>
                <option <% if session("type") = "Existing Loan" then Response.Write " selected" end if%> value="Existing Loan">Existing Loan</option>
                <option <% if session("type") = "New Loan" then Response.Write " selected" end if%> value="New Loan">New Loan</option>
                <option <% if session("type") = "New Order" then Response.Write " selected" end if%> value="New Order">New Order</option>
              </select></td>
          </tr>
          <tr>
            <td>Available for sale?</td>
            <td><select name="cboPreSold">
                <option <% if session("pre_sold") = "No" then Response.Write " selected" end if%> value="No">No</option>
                <option <% if session("pre_sold") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComment" id="txtComment" cols="60" rows="8"><%= session("comments") %></textarea></td>
          </tr>
          <tr>
            <td>After-show Return to<span class="mandatory">*</span>:</td>
            <td><select name="cboReturnTo">
                <option <% if session("return_to") = "NA" then Response.Write " selected" end if%> value="NA">...</option>
                <option <% if session("return_to") = "Dealer" then Response.Write " selected" end if%> value="Dealer">Dealer</option>
                <option <% if session("return_to") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
                <option <% if session("return_to") = "Kagan" then Response.Write " selected" end if%> value=" Kagan">Kagan</option>
                <option <% if session("return_to") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
              </select></td>
          </tr>
          <tr>
            <td>Status:</td>
            <td><select name="cboStatus">
                <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
              </select></td>
          </tr>
          <tr>
            <td valign="top"><input type="hidden" name="Action" />
              <input type="submit" value="Update" />
              <p><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_item.asp">Back To Roadshow List</a></p></td>
            <td><table width="300" cellpadding="4" cellspacing="0" bgcolor="#CCCCCC">
            	<tr>
                <td width="50%">Date Created:</td>
                <td width="50%"><%= session("item_date_created") %></td>
              </tr>
              <tr>
                <td>Created By:</td>
                <td><%= session("item_created_by") %></td>
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