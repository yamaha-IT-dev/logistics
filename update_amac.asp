<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update AMAC Item</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtProductCode,"Product Code");
	//reason += validateSpecialCharacters(theForm.txtProductCode,"Product Code");
	
	reason += validateEmptyField(theForm.txtDescription,"Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Description");
	
	reason += validateNumeric(theForm.txtQuantity,"Quantity");

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

</script>
<%

Sub getItem

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_amac WHERE item_id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("department") 		= rs("department")
		session("item_group") 		= rs("item_group")
		session("category") 		= rs("category")
		session("product_code") 	= rs("product_code")		
		session("description") 		= rs("description")				
		session("rrp") 				= rs("rrp")
		session("quantity") 		= rs("quantity")
		session("sku_type") 		= rs("sku_type")
		session("prototype") 		= rs("prototype")		
		session("packaging") 		= rs("packaging")
		session("source") 			= rs("source")
		session("origin") 			= rs("origin")
		session("available") 		= rs("available")
		session("transit") 			= rs("transit")
		session("invoice_no") 		= rs("invoice_no")
		session("type") 			= rs("type")
		session("available_for_sale") = rs("available_for_sale")
		session("pre_sold") 		= rs("pre_sold")
		session("return_to") 		= rs("return_to")		
		session("displayed") 		= rs("displayed")
		session("fb_completed")		= rs("fb_completed")
		session("status") 			= rs("status")
		session("date_created")		= rs("date_created")
		session("created_by") 		= rs("created_by")
		session("date_modified") 	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
		session("logistics_action") = rs("logistics_action")
		session("pallet_no") 		= rs("pallet_no")
		session("loading_sequence") = rs("loading_sequence")
    end if

    call CloseDataBase()
end sub

sub updateItem
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE tbl_amac SET "
	strSQL = strSQL & "product_code = '" & Server.HTMLEncode(Request.Form("txtProductCode")) & "',"
	strSQL = strSQL & "item_group = '" & Trim(Request.Form("txtItemGroup")) & "',"
	strSQL = strSQL & "description = '" & Replace(Request.Form("txtDescription"),"'","''") & "',"
	strSQL = strSQL & "category = '" & Trim(Request.Form("cboCategory")) & "',"
	strSQL = strSQL & "department = '" & Trim(Request.Form("cboDepartment")) & "',"
	strSQL = strSQL & "rrp = '" & Trim(Request.Form("txtRRP")) & "',"
	strSQL = strSQL & "sku_type = '" & Trim(Request.Form("cboSkuType")) & "',"
	strSQL = strSQL & "prototype = '" & Trim(Request.Form("cboPrototype")) & "',"
	strSQL = strSQL & "quantity = '" & Trim(Request.Form("txtQuantity")) & "',"
	strSQL = strSQL & "packaging = '" & Trim(Request.Form("cboPackaging")) & "',"
	strSQL = strSQL & "source = '" & Trim(Request.Form("cboSource")) & "',"
	strSQL = strSQL & "origin = '" & Trim(Request.Form("cboOrigin")) & "',"
	strSQL = strSQL & "available = '" & Trim(Request.Form("cboAvailable")) & "',"
	strSQL = strSQL & "transit = '" & Trim(Request.Form("cboInTransit")) & "',"
	strSQL = strSQL & "invoice_no = '" & Trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "pallet_no = '" & Trim(Request.Form("txtPalletNo")) & "',"
	strSQL = strSQL & "loading_sequence = '" & Trim(Request.Form("txtLoadingSequence")) & "',"
	strSQL = strSQL & "type = '" & Trim(Request.Form("cboType")) & "',"
	strSQL = strSQL & "available_for_sale = '" & Trim(Request.Form("cboAvailableForSale")) & "',"
	strSQL = strSQL & "pre_sold = '" & Trim(Request.Form("cboPreSold")) & "',"
	'strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComment"),"'","''") & "',"
	'strSQL = strSQL & "bump_in_date = CONVERT(datetime,'" &	Trim(Request.Form("txtBumpInDate")) & "',103),"
	'strSQL = strSQL & "bump_out_date = CONVERT(datetime,'" &	Trim(Request.Form("txtBumpOutDate")) & "',103),"
	strSQL = strSQL & "return_to = '" & Trim(Request.Form("cboReturnTo")) & "',"
	strSQL = strSQL & "displayed = '" & Trim(Request.Form("txtDisplay")) & "',"
	strSQL = strSQL & "fb_completed = '" & Trim(Request.Form("cboFB")) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "logistics_action = '" & Trim(Request.Form("chkLogisticsAction")) & "',"
	strSQL = strSQL & "status = '" & Trim(Request.Form("cboStatus")) & "' WHERE item_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Update" then
			call updateItem
		end if
	end if
	
	call getItem
end sub

dim strMessageText
call main
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_amac.asp">Back to List</a>
            <h2>Update AMAC<u></u></h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Last modified:</strong></td>
                <td class="created_column_2"><%= session("modified_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_roadshow" id="form_update_roadshow" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" width="1024">
          <tr>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Item Details</td>
                </tr>
                <tr>
                  <td width="25%" align="right">Dept:</td>
                  <td width="75%"><select name="cboDepartment">
                      <option <% if session("department") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
                      <option <% if session("department") = "Pro" then Response.Write " selected" end if%> value="Pro">Pro</option>
                      <option <% if session("department") = "Trad" then Response.Write " selected" end if%> value="Trad">Trad</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Group:</td>
                  <td><input type="text" id="txtItemGroup" name="txtItemGroup" maxlength="5" size="5" value="<%= session("item_group") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Category:</td>
                  <td><select name="cboCategory">
                      <option <% if session("category") = "Acoustic Drums" then Response.Write " selected" end if%> value="Acoustic Drums">Acoustic Drums</option>
                      <option <% if session("category") = "Brass & Woodwind" then Response.Write " selected" end if%> value="Brass & Woodwind">Brass & Woodwind</option>
                      <option <% if session("category") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
                      <option <% if session("category") = "Digital Pianos" then Response.Write " selected" end if%> value="Digital Pianos">Digital Pianos</option>
                      <option <% if session("category") = "Electronic Drums" then Response.Write " selected" end if%> value="Electronic Drums">Electronic Drums</option>
                      <option <% if session("category") = "Guitars" then Response.Write " selected" end if%> value="Guitars">Guitars</option>
                      <option <% if session("category") = "MPP" then Response.Write " selected" end if%> value="MPP">MPP</option>
                      <option <% if session("category") = "Paiste" then Response.Write " selected" end if%> value="Paiste">Paiste</option>
                      <option <% if session("category") = "Percussion & Strings" then Response.Write " selected" end if%> value="Percussion & Strings">Percussion & Strings</option>
                      <option <% if session("category") = "Pianos" then Response.Write " selected" end if%> value="Pianos">Pianos</option>
                      <option <% if session("category") = "Portable Keyboards" then Response.Write " selected" end if%> value="Portable Keyboards">Portable Keyboards</option>
                      <option <% if session("category") = "Pro Audio" then Response.Write " selected" end if%> value="Pro Audio">Pro Audio</option>
                      <option <% if session("category") = "SYDE" then Response.Write " selected" end if%> value="SYDE">SYDE</option>
                      <option <% if session("category") = "VOX" then Response.Write " selected" end if%> value="VOX">VOX</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Product code<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtProductCode" name="txtProductCode" maxlength="20" size="20" value="<%= session("product_code") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Description<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtDescription" name="txtDescription" maxlength="60" size="55" value="<%= Server.HTMLEncode(session("description")) %>" /></td>
                </tr>
                <tr>
                  <td align="right">RRP:</td>
                  <td>$
                    <input type="text" id="txtRRP" name="txtRRP" maxlength="8" size="8" value="<%= session("rrp") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Qty<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtQuantity" name="txtQuantity" maxlength="4" size="4" value="<%= session("quantity") %>" /></td>
                </tr>
                <tr>
                  <td align="right">SKU:</td>
                  <td><select name="cboSkuType">
                      <option <% if session("sku_type") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("sku_type") = "A Sku" then Response.Write " selected" end if%> value="A Sku">A Sku</option>
                      <option <% if session("sku_type") = "B Sku" then Response.Write " selected" end if%> value="B Sku">B Sku</option>
                      <option <% if session("sku_type") = "New" then Response.Write " selected" end if%> value="New">New</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Prototype:</td>
                  <td><select name="cboPrototype">
                      <option <% if session("prototype") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                      <option <% if session("prototype") = "No" then Response.Write " selected" end if%> value="No">No</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Packaging:</td>
                  <td><select name="cboPackaging">
                      <option <% if session("packaging") = "Carton" then Response.Write " selected" end if%> value="Carton">Carton</option>
                      <option <% if session("packaging") = "Roadcase" then Response.Write " selected" end if%> value="Roadcase">Roadcase</option>
                      <option <% if session("packaging") = "Unboxed" then Response.Write " selected" end if%> value="Unboxed">Unboxed</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Source:</td>
                  <td><select name="cboSource">
                      <option <% if session("source") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("source") = "Asset" then Response.Write " selected" end if%> value="Asset">Asset</option>
                      <option <% if session("source") = "Loan Stock" then Response.Write " selected" end if%> value="Loan Stock">Loan Stock</option>
                      <option <% if session("source") = "New Pick" then Response.Write " selected" end if%> value="New Pick">New Pick</option>
                      <option <% if session("source") = "Promo Stock" then Response.Write " selected" end if%> value="Promo Stock">Promo Stock</option>
                      <option <% if session("source") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                    </select></td>
                </tr>
              </table></td>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Logistics Details</td>
                </tr>
                <tr>
                  <td width="30%" align="right">Origin:</td>
                  <td width="70%"><select name="cboOrigin">
                      <option <% if session("origin") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                      <option <% if session("origin") = "Direct from VOX" then Response.Write " selected" end if%> value="Direct from VOX">Direct from VOX</option>
                      <option <% if session("origin") = "Direct from YCJ" then Response.Write " selected" end if%> value="Direct from YCJ">Direct from YCJ</option>
                      <option <% if session("origin") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Available?</td>
                  <td><select name="cboAvailable">
                      <option <% if session("available") = "In Transit" then Response.Write " selected" end if%> value="In Transit">In Transit</option>
                      <option <% if session("available") = "No" then Response.Write " selected" end if%> value="No">No</option>
                      <option <% if session("available") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Transit:</td>
                  <td><select name="cboInTransit">
                      <option <% if session("transit") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("transit") = "Air Freight" then Response.Write " selected" end if%> value="Air Freight">Air Freight</option>
                      <option <% if session("transit") = "Sea Freight" then Response.Write " selected" end if%> value="Sea Freight">Sea Freight</option>
                      <option <% if session("transit") = "NA" then Response.Write " selected" end if%> value="NA">NA</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Type:</td>
                  <td><select name="cboType">
                      <option <% if session("type") = "Existing Asset" then Response.Write " selected" end if%> value="Existing Asset">Existing Asset</option>
                      <option <% if session("type") = "Existing Loan" then Response.Write " selected" end if%> value="Existing Loan">Existing Loan</option>
                      <option <% if session("type") = "New Loan" then Response.Write " selected" end if%> value="New Loan">New Loan</option>
                      <option <% if session("type") = "New Order" then Response.Write " selected" end if%> value="New Order">New Order</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Available for sale?</td>
                  <td><select name="cboAvailableForSale">
                      <option <% if session("available_for_sale") = "No" then Response.Write " selected" end if%> value="No">No</option>
                      <option <% if session("available_for_sale") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Pre-sold?</td>
                  <td><select name="cboPreSold">
                      <option <% if session("pre_sold") = "No" then Response.Write " selected" end if%> value="No">No</option>
                      <option <% if session("pre_sold") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">After-show Return to<span class="mandatory">*</span>:</td>
                  <td><select name="cboReturnTo">
                      <option <% if session("return_to") = "NA" then Response.Write " selected" end if%> value="NA">...</option>
                      <option <% if session("return_to") = "Dealer" then Response.Write " selected" end if%> value="Dealer">Dealer</option>
                      <option <% if session("return_to") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
                      <option <% if session("return_to") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                      <option <% if session("return_to") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Display at:</td>
                  <td><input type="text" id="txtDisplay" name="txtDisplay" maxlength="30" size="30" value="<%= session("displayed") %>" /></td>
                </tr>
                <tr>
                  <td align="right">F &amp; B completed?</td>
                  <td><select name="cboFB">
                      <option <% if session("fb_completed") = "No" then Response.Write " selected" end if%> value="No">No</option>
                      <option <% if session("fb_completed") = "Yes" then Response.Write " selected" end if%> value="Yes">Yes</option>
                    </select></td>
                </tr>
                <tr align="right" bgcolor="#99FF66">
                  <td colspan="2"><input type="checkbox" name="chkLogisticsAction" id="chkLogisticsAction" value="1" <% if session("logistics_action") = "1" then Response.Write " checked" end if%> />
                    Logistics Actioned</td>
                </tr>
                <tr class="status_row">
                  <td align="right">Invoice no:</td>
                  <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" value="<%= session("invoice_no") %>" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Pallet no:</td>
                  <td><input type="text" id="txtPalletNo" name="txtPalletNo" maxlength="20" size="20" value="<%= session("pallet_no") %>" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Loading sequence:</td>
                  <td><input type="text" id="txtLoadingSequence" name="txtLoadingSequence" maxlength="20" size="20" value="<%= session("loading_sequence") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Status:</td>
                  <td><select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="2" valign="top" align="center"><input type="hidden" name="Action" />
              <input type="submit" value="Update" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>